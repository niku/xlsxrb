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

      # Opens a ZIP from a file path or IO and yields the reader.
      def self.open(source)
        io = source.is_a?(String) ? File.open(source, "rb") : source
        reader = new(io)
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

      def initialize(io)
        @io = io
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
        result = {}

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

        io.seek(-4, IO::SEEK_CUR) if io.respond_to?(:seek)
        io = StringIO.new(first_sig + io.read.b) unless io.respond_to?(:seek)

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
            raw, = find_data_descriptor_stream(io, method)
          else
            raw = io.read(compressed_size)
          end
          entry_data = decompress(raw, method)
          result[entry_name] = entry_data unless entry_name.end_with?("/")
        end

        if io.is_a?(Tempfile)
          io.close
          io.unlink
        end

        result
      end

      def find_data_descriptor_stream(io, method)
        if method == 8
          inflater = Zlib::Inflate.new(-Zlib::MAX_WBITS)
          result = +""
          consumed = 0
          chunk_size = 4096
          begin
            while (chunk = io.read(chunk_size))
              break if chunk.empty?

              inflated = inflater.inflate(chunk)
              raise ArgumentError, "ZIP bomb detected: Uncompressed size exceeds #{MAX_UNCOMPRESSED_SIZE} bytes" if result.bytesize + inflated.bytesize > MAX_UNCOMPRESSED_SIZE

              result << inflated
              consumed += chunk.bytesize
            end
          rescue Zlib::BufError, Zlib::DataError
            # Inflation ended
          ensure
            consumed -= inflater.avail_in
            inflater.close
          end

          io.seek(-inflater.avail_in, IO::SEEK_CUR) if io.respond_to?(:seek)

          desc_sig = io.read(4)
          io.seek(-4, IO::SEEK_CUR) unless desc_sig == [0x50, 0x4B, 0x07, 0x08].pack("C4")
          io.read(12)

          [result, consumed]
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

        begin
          while offset < raw_len
            chunk = raw.byteslice(offset, chunk_size)
            inflated = inflater.inflate(chunk)
            raise ArgumentError, "ZIP bomb detected: Uncompressed size exceeds #{MAX_UNCOMPRESSED_SIZE} bytes" if result.bytesize + inflated.bytesize > MAX_UNCOMPRESSED_SIZE

            result << inflated
            offset += chunk_size
          end
          # Finish inflation
          inflated = inflater.finish
          raise ArgumentError, "ZIP bomb detected: Uncompressed size exceeds #{MAX_UNCOMPRESSED_SIZE} bytes" if result.bytesize + inflated.bytesize > MAX_UNCOMPRESSED_SIZE

          result << inflated
        ensure
          inflater.close
        end

        result.force_encoding("UTF-8")
      end
    end
  end
end
