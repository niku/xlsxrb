# frozen_string_literal: true

# rbs_inline: enabled

require "zlib"
require "stringio"
require "tempfile"

module Xlsxrb
  module Ooxml
    # Reads ZIP archives using only stdlib (zlib).
    # Supports streaming entry decompression with bounded memory consumption.
    class ZipReader
      LOCAL_HEADER_SIG = "PK\x03\x04".b
      MAX_UNCOMPRESSED_SIZE = 500 * 1024 * 1024 # 500MB per file limit
      attr_accessor :max_uncompressed_size

      # Metadata record for a single archive entry.
      class Entry
        attr_reader :name, :method, :compressed_size, :uncompressed_size, :crc32, :local_header_offset, :data_offset

        #: (name: String, method: Integer, compressed_size: Integer, uncompressed_size: Integer, crc32: Integer, local_header_offset: Integer, data_offset: Integer) -> void
        def initialize(name:, method:, compressed_size:, uncompressed_size:, crc32:, local_header_offset:, data_offset:)
          @name = name
          @method = method
          @compressed_size = compressed_size
          @uncompressed_size = uncompressed_size
          @crc32 = crc32
          @local_header_offset = local_header_offset
          @data_offset = data_offset
        end
      end

      # IO-like stream wrapper that decompresses entry chunks on demand without full buffering.
      class EntryStreamIO
        DEFAULT_CHUNK_SIZE = 65_536

        attr_reader :data_offset, :compressed_size, :compression_method

        #: (untyped io, Integer data_offset, Integer compressed_size, Integer compression_method, ?max_uncompressed_size: Integer?) -> void
        def initialize(io, data_offset, compressed_size, compression_method, max_uncompressed_size: nil)
          @io = io
          @data_offset = data_offset
          @compressed_size = compressed_size
          @compression_method = compression_method
          @max_uncompressed_size = max_uncompressed_size || ZipReader::MAX_UNCOMPRESSED_SIZE
          @buffer = +""
          @bytes_read_compressed = 0
          @total_uncompressed = 0
          @eof = false
          @closed = false
          @inflater = @compression_method == 8 ? Zlib::Inflate.new(-Zlib::MAX_WBITS) : nil
        end

        #: (?Integer? length) -> String?
        def read(length = nil)
          raise IOError, "closed stream" if @closed

          if length.nil?
            result = @buffer.dup
            @buffer.clear
            while (chunk = fetch_chunk(DEFAULT_CHUNK_SIZE))
              result << chunk
            end
            return result
          end

          return "" if length.zero?

          while @buffer.bytesize < length && !@eof
            chunk = fetch_chunk([length - @buffer.bytesize, DEFAULT_CHUNK_SIZE].max)
            break unless chunk

            @buffer << chunk
          end

          return nil if @buffer.empty? && @eof

          @buffer.slice!(0, length)
        end

        #: (?Integer chunk_size) ?{ (String) -> void } -> (void | Enumerator[String, void])
        def each_chunk(chunk_size = DEFAULT_CHUNK_SIZE, &block)
          return enum_for(:each_chunk, chunk_size) unless block
          raise IOError, "closed stream" if @closed

          while (chunk = read(chunk_size))
            break if chunk.nil? || chunk.empty?

            yield chunk
          end
        end

        #: () -> void
        def close
          return if @closed

          @closed = true
          @inflater&.close
          @inflater = nil
          @buffer.clear
          nil
        end

        #: () -> bool
        def closed?
          @closed
        end

        #: () -> bool
        def eof?
          @eof && @buffer.empty?
        end

        private

        def fetch_chunk(chunk_size)
          return nil if @eof || @closed

          if @compression_method.zero?
            remaining = @compressed_size - @bytes_read_compressed
            if remaining <= 0
              @eof = true
              return nil
            end

            read_len = [remaining, chunk_size].min
            @io.seek(@data_offset + @bytes_read_compressed, IO::SEEK_SET)
            raw = @io.read(read_len)
            if raw.nil? || raw.empty?
              @eof = true
              return nil
            end

            @bytes_read_compressed += raw.bytesize
            @total_uncompressed += raw.bytesize
            raise ArgumentError, "ZIP bomb detected: Uncompressed size exceeds #{@max_uncompressed_size} bytes" if @total_uncompressed > @max_uncompressed_size

            raw
          elsif @compression_method == 8
            inflated = +""
            while inflated.empty? && !@eof
              if @bytes_read_compressed < @compressed_size
                remaining_compressed = @compressed_size - @bytes_read_compressed
                read_len = [remaining_compressed, chunk_size].min
                @io.seek(@data_offset + @bytes_read_compressed, IO::SEEK_SET)
                raw = @io.read(read_len)
                if raw.nil? || raw.empty?
                  @eof = true
                else
                  @bytes_read_compressed += raw.bytesize
                  inflater = @inflater
                  inflated << inflater.inflate(raw) if inflater
                end
              else
                inflater = @inflater
                if inflater && !inflater.finished?
                  begin
                    inflated << inflater.finish
                  rescue Zlib::BufError, Zlib::DataError
                    # Stream ended
                  end
                end
                @eof = true
              end
            end

            @eof = true if @inflater&.finished?

            @total_uncompressed += inflated.bytesize
            raise ArgumentError, "ZIP bomb detected: Uncompressed size exceeds #{@max_uncompressed_size} bytes" if @total_uncompressed > @max_uncompressed_size

            inflated.empty? && @eof ? nil : inflated
          else
            @eof = true
            nil
          end
        end
      end

      # Opens a ZIP from a file path or IO and yields the reader.
      #: (untyped source, ?max_uncompressed_size: Integer) ?{ (ZipReader) -> untyped } -> (ZipReader | untyped)
      def self.open(source, max_uncompressed_size: MAX_UNCOMPRESSED_SIZE)
        should_close = source.is_a?(String) && File.file?(source)
        io = if source.is_a?(String)
               if source.start_with?(LOCAL_HEADER_SIG) || source.include?("\x00")
                 StringIO.new(source.b)
               else
                 File.open(source, "rb")
               end
             else
               source
             end
        reader = new(io, max_uncompressed_size: max_uncompressed_size, should_close_io: should_close)
        if block_given?
          begin
            yield reader
          ensure
            reader.close
          end
        else
          reader
        end
      end

      #: (untyped io, ?max_uncompressed_size: Integer, ?should_close_io: bool) -> void
      def initialize(io, max_uncompressed_size: MAX_UNCOMPRESSED_SIZE, should_close_io: false)
        @raw_io = io
        @max_uncompressed_size = max_uncompressed_size
        @should_close_io = should_close_io
        @tempfile = nil
        @io = nil
        @catalog = nil
        @cached_entries = nil
      end

      # Returns a Hash { entry_name => raw_bytes } for all entries.
      #: () -> Hash[String, String?]
      def read_all
        result = {}
        catalog.each_key do |name|
          result[name] = read_entry(name)
        end
        result
      end

      # Returns raw bytes for a single entry, or nil if not found.
      #: (String name) -> String?
      def read_entry(name)
        return @cached_entries[name] if @cached_entries&.key?(name)

        entry = catalog[name]
        return nil unless entry

        io.seek(entry.data_offset, IO::SEEK_SET)
        raw = io.read(entry.compressed_size)
        decompress(raw, entry.method)
      end

      # Yields (entry_name, data_string) for each file in the archive.
      #: () ?{ (String, String?) -> void } -> (void | Enumerator[[String, String?], void])
      def each_entry(&block)
        return enum_for(:each_entry) unless block

        catalog.each_key do |name|
          block.call(name, read_entry(name))
        end
      end

      # Opens an EntryStreamIO for an entry by name.
      #: (String entry_name) -> EntryStreamIO
      def open_entry_io(entry_name)
        entry = catalog[entry_name]
        raise ArgumentError, "Entry not found in archive: #{entry_name}" unless entry

        EntryStreamIO.new(io, entry.data_offset, entry.compressed_size, entry.method, max_uncompressed_size: @max_uncompressed_size)
      end

      # Yields inflated binary chunks directly from the compressed stream.
      #: (String entry_name, ?chunk_size: Integer) ?{ (String) -> void } -> (void | Enumerator[String, void])
      def each_entry_chunk(entry_name, chunk_size: 65_536, &block)
        return enum_for(:each_entry_chunk, entry_name, chunk_size: chunk_size) unless block

        stream = open_entry_io(entry_name)
        begin
          stream.each_chunk(chunk_size, &block)
        ensure
          stream.close
        end
      end

      # Returns true if the archive contains the specified entry.
      #: (String name) -> bool
      def entry?(name)
        catalog.key?(name)
      end

      # Returns entry metadata catalog.
      #: () -> Hash[String, Entry]
      def entry_catalog
        catalog
      end

      # Returns array of all entry names.
      #: () -> Array[String]
      def entry_names
        catalog.keys
      end

      # Copies the specified entry to a ZipWriter directly without decompression.
      #: (String entry_name, ZipWriter writer, ?target_path: String?) -> void
      def copy_to_writer(entry_name, writer, target_path: nil)
        entry = catalog[entry_name]
        raise ArgumentError, "Entry not found in archive: #{entry_name}" unless entry

        if entry.compressed_size.zero? && entry.uncompressed_size.positive?
          data = read_entry(entry_name)
          writer.add_binary_entry(target_path || entry_name, data || "")
          return
        end

        writer.copy_raw_entry(
          target_path || entry_name,
          io,
          data_offset: entry.data_offset,
          compressed_size: entry.compressed_size,
          uncompressed_size: entry.uncompressed_size,
          crc32: entry.crc32,
          method: entry.method
        )
      end

      # Closes the archive reader and cleans up any temporary resources.
      #: () -> void
      def close
        @raw_io.close if @should_close_io && @raw_io.respond_to?(:close) && !@raw_io.closed?
        if @tempfile
          @tempfile.close
          @tempfile.unlink
          @tempfile = nil
        end
        nil
      end

      private

      def io
        @io ||= begin
          prepared = if @raw_io.is_a?(StringIO) || (@raw_io.respond_to?(:read) && @raw_io.respond_to?(:seek) && @raw_io.respond_to?(:pos))
                       @raw_io
                     elsif @raw_io.respond_to?(:each) || @raw_io.is_a?(Enumerator)
                       @tempfile = Tempfile.new("xlsxrb_zip")
                       @tempfile.binmode
                       @raw_io.each { |chunk| @tempfile.write(chunk) }
                       @tempfile.rewind
                       @tempfile
                     else
                       StringIO.new(@raw_io.read.b)
                     end
          prepared.binmode if prepared.respond_to?(:binmode)
          prepared
        end
      end

      def catalog
        @catalog ||= parse_catalog
      end

      def parse_catalog
        first_sig = io.read(4)
        raise ArgumentError, "Invalid magic number: Expected a valid ZIP/XLSX file format (PK\\x03\\x04)" unless first_sig == LOCAL_HEADER_SIG

        # Try Central Directory first
        cd_catalog = parse_catalog_from_central_directory(io)
        return cd_catalog if cd_catalog && !cd_catalog.empty?

        io.seek(-4, IO::SEEK_CUR) if io.respond_to?(:seek)
        parse_catalog_from_local_headers(io)
      end

      def parse_catalog_from_central_directory(io)
        return nil unless io.respond_to?(:seek) && io.respond_to?(:pos)

        io.seek(0, IO::SEEK_END)
        file_size = io.pos
        return nil if file_size < 22

        max_search = [file_size, 65_557].min
        search_offset = file_size - max_search
        io.seek(search_offset, IO::SEEK_SET)
        search_buf = io.read(max_search)
        return nil unless search_buf

        eocd_index = search_buf.rindex("PK\x05\x06")
        return nil unless eocd_index

        eocd_pos = search_offset + eocd_index
        io.seek(eocd_pos + 4, IO::SEEK_SET)
        eocd_data = io.read(18)
        return nil unless eocd_data && eocd_data.bytesize == 18

        _disk_num, _cd_disk, _disk_entries, total_entries, _cd_size, cd_offset, _comment_len = eocd_data.unpack("vvvvVVv")
        return nil if cd_offset > file_size

        io.seek(cd_offset, IO::SEEK_SET)
        result = {}

        total_entries.times do
          sig = io.read(4)
          break unless sig == "PK\x01\x02"

          cd_header = io.read(42)
          break unless cd_header && cd_header.bytesize == 42

          _v_made, _v_need, _gp, method, _time, _date, crc, csize, usize, nlen, elen, clen, _dnum, _iattr, _eattr, offset = cd_header.unpack("vvvvvvVVVvvvvvVV")
          entry_name = io.read(nlen).force_encoding("UTF-8")
          io.read(elen + clen)

          next if entry_name.end_with?("/")

          current_cd_pos = io.pos

          io.seek(offset, IO::SEEK_SET)
          local_sig = io.read(4)
          next unless local_sig == LOCAL_HEADER_SIG

          local_header = io.read(26)
          next unless local_header && local_header.bytesize == 26

          local_nlen = local_header[22, 2].unpack1("v")
          local_elen = local_header[24, 2].unpack1("v")
          data_offset = offset + 30 + local_nlen + local_elen

          result[entry_name] = Entry.new(
            name: entry_name,
            method: method,
            compressed_size: csize,
            uncompressed_size: usize,
            crc32: crc,
            local_header_offset: offset,
            data_offset: data_offset
          )

          io.seek(current_cd_pos, IO::SEEK_SET)
        end

        result
      rescue ArgumentError => e
        raise e
      rescue StandardError
        nil
      end

      def parse_catalog_from_local_headers(io)
        result = {}
        loop do
          header_pos = io.pos
          sig = io.read(4)
          break unless sig == LOCAL_HEADER_SIG

          header = io.read(26)
          break unless header && header.bytesize == 26

          gp_flag         = header[2, 2].unpack1("v")
          method          = header[4, 2].unpack1("v")
          crc             = header[10, 4].unpack1("V")
          compressed_size = header[14, 4].unpack1("V")
          uncompressed_size = header[18, 4].unpack1("V")
          name_len        = header[22, 2].unpack1("v")
          extra_len       = header[24, 2].unpack1("v")

          entry_name = io.read(name_len).force_encoding("UTF-8")
          io.read(extra_len)
          data_offset = io.pos

          next if entry_name.end_with?("/")

          has_data_descriptor = gp_flag.anybits?(0x08)

          if has_data_descriptor && compressed_size.zero?
            entry_data, = find_data_descriptor_stream(io, method)
            @cached_entries ||= {}
            @cached_entries[entry_name] = entry_data
            result[entry_name] = Entry.new(
              name: entry_name,
              method: method,
              compressed_size: entry_data.bytesize,
              uncompressed_size: entry_data.bytesize,
              crc32: crc,
              local_header_offset: header_pos,
              data_offset: data_offset
            )
          else
            result[entry_name] = Entry.new(
              name: entry_name,
              method: method,
              compressed_size: compressed_size,
              uncompressed_size: uncompressed_size,
              crc32: crc,
              local_header_offset: header_pos,
              data_offset: data_offset
            )

            io.seek(compressed_size, IO::SEEK_CUR) if io.respond_to?(:seek)
          end
        end
        result
      end

      def find_data_descriptor_stream(io, method)
        if method == 8
          inflater = Zlib::Inflate.new(-Zlib::MAX_WBITS)
          result = +""
          chunk_size = 4096
          begin
            while !inflater.finished? && (chunk = io.read(chunk_size))
              break if chunk.empty?

              inflated = inflater.inflate(chunk)
              limit = @max_uncompressed_size || MAX_UNCOMPRESSED_SIZE
              raise ArgumentError, "ZIP bomb detected: Uncompressed size exceeds #{limit} bytes" if result.bytesize + inflated.bytesize > limit

              result << inflated
            end
          rescue Zlib::BufError, Zlib::DataError
            # Inflation ended
          ensure
            remaining_bytes = inflater.avail_in
            inflater.close
          end

          io.seek(-remaining_bytes, IO::SEEK_CUR) if io.respond_to?(:seek) && remaining_bytes.positive?

          desc_sig = io.read(4)
          if desc_sig == [0x50, 0x4B, 0x07, 0x08].pack("C4")
            io.read(12)
          elsif desc_sig == [0x50, 0x4B, 0x03, 0x04].pack("C4") || desc_sig == [0x50, 0x4B, 0x01, 0x02].pack("C4")
            io.seek(-4, IO::SEEK_CUR) if io.respond_to?(:seek)
          elsif desc_sig
            io.read(8)
          end

          [result, 0]
        else
          [io.read(0), 0]
        end
      end

      def decompress(raw, method)
        return raw&.dup&.force_encoding("UTF-8") || "" if method.zero?

        safe_inflate(raw || "", -Zlib::MAX_WBITS)
      rescue Zlib::DataError
        safe_inflate(raw || "", -Zlib::MAX_WBITS)
      end

      def safe_inflate(raw, wbits)
        inflater = Zlib::Inflate.new(wbits)
        result = +""
        chunk_size = 32_768
        offset = 0
        raw_len = raw.bytesize
        limit = @max_uncompressed_size || MAX_UNCOMPRESSED_SIZE

        begin
          while offset < raw_len
            chunk = raw.byteslice(offset, chunk_size)
            inflated = inflater.inflate(chunk)
            raise ArgumentError, "ZIP bomb detected: Uncompressed size exceeds #{limit} bytes" if result.bytesize + inflated.bytesize > limit

            result << inflated
            offset += chunk_size
          end
          inflated = inflater.finish
          raise ArgumentError, "ZIP bomb detected: Uncompressed size exceeds #{limit} bytes" if result.bytesize + inflated.bytesize > limit

          result << inflated
        ensure
          inflater.close
        end

        result.force_encoding("UTF-8")
      end
    end
  end
end
