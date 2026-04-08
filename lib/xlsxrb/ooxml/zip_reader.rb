# frozen_string_literal: true

require "zlib"
require "stringio"

module Xlsxrb
  module Ooxml
    # Reads ZIP archives using only stdlib (zlib).
    # Scans local file headers sequentially — works with non-seekable IO.
    class ZipReader
      LOCAL_HEADER_SIG = [0x50, 0x4B, 0x03, 0x04].pack("C4")

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
        data = @io.is_a?(StringIO) ? @io.string : @io.read
        data = data.b
        pos = 0

        while pos + 4 <= data.bytesize
          sig = data[pos, 4]
          break unless sig == LOCAL_HEADER_SIG

          # Local file header: 4 sig + 2 version + 2 flags + 2 method + 2 time + 2 date
          #   + 4 crc + 4 compressed + 4 uncompressed + 2 name_len + 2 extra_len = 30 bytes
          break if pos + 30 > data.bytesize

          gp_flag        = data[pos + 6, 2].unpack1("v")
          method         = data[pos + 8, 2].unpack1("v")
          data[pos + 14, 4].unpack1("V")
          compressed_size = data[pos + 18, 4].unpack1("V")
          data[pos + 22, 4].unpack1("V")
          name_len       = data[pos + 26, 2].unpack1("v")
          extra_len      = data[pos + 28, 2].unpack1("v")

          entry_name = data[pos + 30, name_len].force_encoding("UTF-8")
          file_data_offset = pos + 30 + name_len + extra_len

          # Data descriptor present if bit 3 of gp_flag is set
          has_data_descriptor = gp_flag.anybits?(0x08)

          if has_data_descriptor && compressed_size.zero?
            # Need to find the data descriptor to know sizes
            raw, comp_sz, = find_data_descriptor(data, file_data_offset, method)
            entry_data = decompress(raw, method)
            result[entry_name] = entry_data unless entry_name.end_with?("/")
            # Skip past data + descriptor (12 or 16 bytes)
            desc_offset = file_data_offset + comp_sz
            # Check for optional signature
            pos = if desc_offset + 4 <= data.bytesize && data[desc_offset, 4] == [0x50, 0x4B, 0x07, 0x08].pack("C4")
                    desc_offset + 16
                  else
                    desc_offset + 12
                  end
          else
            raw = data[file_data_offset, compressed_size]
            entry_data = decompress(raw, method)
            result[entry_name] = entry_data unless entry_name.end_with?("/")
            pos = file_data_offset + compressed_size
          end
        end

        result
      end

      def find_data_descriptor(data, offset, method)
        # For deflated data, we inflate to find the end
        if method == 8
          inflater = Zlib::Inflate.new(-Zlib::MAX_WBITS)
          result = +""
          consumed = 0
          chunk_size = 4096
          pos = offset
          begin
            while pos < data.bytesize
              chunk = data[pos, [chunk_size, data.bytesize - pos].min]
              break if chunk.nil? || chunk.empty?

              result << inflater.inflate(chunk)
              consumed += chunk.bytesize
              pos += chunk.bytesize
            end
          rescue Zlib::BufError, Zlib::DataError
            # Inflation ended — find actual consumed size
          ensure
            # Calculate actual compressed bytes consumed
            consumed -= inflater.avail_in
            inflater.close
          end
          [result, consumed, result.bytesize]
        else
          # Stored — scan for data descriptor signature or central directory
          # This is a fallback; stored + data descriptor is rare
          [data[offset, 0], 0, 0]
        end
      end

      def decompress(raw, method)
        return raw&.dup&.force_encoding("UTF-8") || "" if method.zero? # stored

        # Deflated
        Zlib::Inflate.inflate(-raw || "")
      rescue Zlib::DataError
        # Try with raw deflate (no header)
        inflater = Zlib::Inflate.new(-Zlib::MAX_WBITS)
        begin
          inflater.inflate(raw || "")
        ensure
          inflater.close
        end
      end
    end
  end
end
