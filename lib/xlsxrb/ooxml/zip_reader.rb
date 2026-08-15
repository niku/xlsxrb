# frozen_string_literal: true

# rbs_inline: enabled

require "zlib"
require "stringio"
require "tempfile"

module Xlsxrb
  module Ooxml
    # Reads ZIP archives using only stdlib (zlib).
    # Scans local file headers sequentially — works with non-seekable IO.
    class ZipReader
      LOCAL_HEADER_SIG = "PK\x03\x04".b
      MAX_UNCOMPRESSED_SIZE = 500 * 1024 * 1024 # 500MB per file limit
      attr_accessor :max_uncompressed_size

      # Opens a ZIP from a file path or IO and yields the reader.
      #: (untyped source, ?max_uncompressed_size: Integer) ?{ (ZipReader) -> untyped } -> (ZipReader | untyped)
      def self.open(source, max_uncompressed_size: MAX_UNCOMPRESSED_SIZE)
        io = source.is_a?(String) ? File.open(source, "rb") : source
        reader = new(io, max_uncompressed_size: max_uncompressed_size)
        if block_given?
          begin
            yield reader
          ensure
            io.close if source.is_a?(String)
          end
        else
          reader
        end
      end

      #: (untyped io, ?max_uncompressed_size: Integer) -> void
      def initialize(io, max_uncompressed_size: MAX_UNCOMPRESSED_SIZE)
        @io = io
        @max_uncompressed_size = max_uncompressed_size
        @entries = nil
      end

      # Returns a Hash { entry_name => raw_bytes } for all entries.
      def read_all
        result = {}
        each_entry { |name, data| result[name] = data }
        result
      end

      # Returns raw bytes for a single entry, or nil if not found.
      def read_entry(name)
        entries[name]
      end

      # Yields (entry_name, data_string) for each file in the archive.
      def each_entry(&block)
        return enum_for(:each_entry) unless block

        entries.each_pair(&block)
      end

      private

      def entries
        @entries ||= parse_entries
      end

      def parse_entries
        io = if @io.is_a?(StringIO) || (@io.respond_to?(:read) && @io.respond_to?(:seek) && @io.respond_to?(:pos))
               @io
             elsif @io.respond_to?(:each) || @io.is_a?(Enumerator)
               tf = Tempfile.new("xlsxrb_zip")
               tf.binmode
               @io.each { |chunk| tf.write(chunk) }
               tf.rewind
               tf
             else
               StringIO.new(@io.read.b)
             end

        io.binmode if io.respond_to?(:binmode)

        first_sig = io.read(4)
        raise ArgumentError, "Invalid magic number: Expected a valid ZIP/XLSX file format (PK\\x03\\x04)" unless first_sig == LOCAL_HEADER_SIG

        # Try parsing via Central Directory first (standard and handles Data Descriptors perfectly)
        cd_result = parse_from_central_directory(io)
        if cd_result && !cd_result.empty?
          if io.is_a?(Tempfile)
            io.close
            io.unlink
          end
          return cd_result
        end

        io.seek(-4, IO::SEEK_CUR) if io.respond_to?(:seek)
        io = StringIO.new(first_sig + io.read.b) unless io.respond_to?(:seek)

        result = {}

        loop do
          sig = io.read(4)
          break unless sig == LOCAL_HEADER_SIG

          header = io.read(26)
          break unless header && header.bytesize == 26

          gp_flag         = header[2, 2].unpack1("v")
          method          = header[4, 2].unpack1("v")
          compressed_size = header[14, 4].unpack1("V")
          name_len        = header[22, 2].unpack1("v")
          extra_len       = header[24, 2].unpack1("v")

          entry_name = io.read(name_len).force_encoding("UTF-8")
          io.read(extra_len) # skip extra

          has_data_descriptor = gp_flag.anybits?(0x08)

          if has_data_descriptor && compressed_size.zero?
            entry_data, = find_data_descriptor_stream(io, method)
          else
            raw = io.read(compressed_size)
            entry_data = decompress(raw, method)
          end
          result[entry_name] = entry_data unless entry_name.end_with?("/")
        end

        if io.is_a?(Tempfile)
          io.close
          io.unlink
        end

        result
      end

      def parse_from_central_directory(io)
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

          _v_made, _v_need, _gp, method, _time, _date, _crc, csize, _usize, nlen, elen, clen, _dnum, _iattr, _eattr, offset = cd_header.unpack("vvvvvvVVVvvvvvVV")
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
          io.seek(offset + 30 + local_nlen + local_elen, IO::SEEK_SET)

          raw = io.read(csize)
          entry_data = decompress(raw, method)
          result[entry_name] = entry_data

          io.seek(current_cd_pos, IO::SEEK_SET)
        end

        result
      rescue ArgumentError => e
        raise e
      rescue StandardError
        nil
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
        return raw&.dup&.force_encoding("UTF-8") || "" if method.zero? # stored

        # Deflated
        safe_inflate(raw || "", -Zlib::MAX_WBITS)
      rescue Zlib::DataError
        # Try with raw deflate (no header)
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
          # Finish inflation
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
