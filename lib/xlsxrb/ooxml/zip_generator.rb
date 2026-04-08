# frozen_string_literal: true

require "zlib"

module Xlsxrb
  module Ooxml
    # Generates a simple ZIP file using only the standard library.
    class ZipGenerator
      def initialize(filepath)
        @filepath = filepath
        @entries = []
        @file = nil
      end

      # Adds an entry to the ZIP archive.
      def add_entry(filepath, content)
        compressed_data = compress(content)
        @entries << {
          path: filepath,
          content: content,
          compressed_data: compressed_data,
          crc32: crc32(content),
          uncompressed_size: content.bytesize,
          compressed_size: compressed_data.bytesize,
          mtime: Time.now
        }
      end

      # Generates the ZIP file.
      def generate
        File.open(@filepath, "wb") do |f|
          @file = f
          write_entries
          write_central_directory
        end
      end

      private

      def write_entries
        @entries.each do |entry|
          write_local_header(entry)
          write_file_data(entry)
        end
      end

      def write_local_header(entry)
        filename = entry[:path]
        filename_bytes = filename.bytes

        # Local file header
        local_header = []
        local_header += [0x50, 0x4b, 0x03, 0x04] # Signature
        local_header += [0x14, 0x00]  # Version needed to extract
        local_header += [0x00, 0x00]  # General purpose bit flag
        local_header += [0x08, 0x00]  # Compression method (8 = deflate)

        # Last mod file time/date
        dos_time = dos_datetime(entry[:mtime])
        local_header += dos_time

        local_header += le32(entry[:crc32]) # CRC-32
        local_header += le32(entry[:compressed_size]) # Compressed size
        local_header += le32(entry[:uncompressed_size]) # Uncompressed size
        local_header += le16(filename_bytes.length) # Filename length
        local_header += le16(0) # Extra field length

        @file.write(local_header.pack("C*"))
        @file.write(filename)
      end

      def write_file_data(entry)
        @file.write(entry[:compressed_data])
      end

      def write_central_directory
        offset = @file.pos
        cd_offset = 0
        @entries.each do |entry|
          filename = entry[:path]
          filename_bytes = filename.bytes

          cd_header = []
          cd_header += [0x50, 0x4b, 0x01, 0x02] # Central directory file header signature
          cd_header += [0x14, 0x03]  # Version made by
          cd_header += [0x14, 0x00]  # Version needed to extract
          cd_header += [0x00, 0x00]  # General purpose bit flag
          cd_header += [0x08, 0x00]  # Compression method

          dos_time = dos_datetime(entry[:mtime])
          cd_header += dos_time

          cd_header += le32(entry[:crc32]) # CRC-32
          cd_header += le32(entry[:compressed_size]) # Compressed size
          cd_header += le32(entry[:uncompressed_size]) # Uncompressed size
          cd_header += le16(filename_bytes.length) # Filename length
          cd_header += le16(0)  # Extra field length
          cd_header += le16(0)  # File comment length
          cd_header += le16(0)  # Disk number start
          cd_header += le16(0)  # Internal file attributes
          cd_header += le32(0)  # External file attributes
          cd_header += le32(cd_offset) # Local header offset

          @file.write(cd_header.pack("C*"))
          @file.write(filename)

          # Advance offset past the local header and compressed data for this entry.
          cd_offset += 30 + filename_bytes.length + entry[:compressed_size]
        end

        # End of central directory record
        end_cd = []
        end_cd += [0x50, 0x4b, 0x05, 0x06] # End of central directory signature
        end_cd += le16(0)  # Disk number
        end_cd += le16(0)  # Disk with central directory
        end_cd += le16(@entries.length)  # Number of central directory records on this disk
        end_cd += le16(@entries.length)  # Total number of central directory records
        end_cd += le32(calculate_central_dir_size) # Size of central directory
        end_cd += le32(offset) # Offset of central directory
        end_cd += le16(0) # Comment length

        @file.write(end_cd.pack("C*"))
      end

      def compress(content)
        deflater = Zlib::Deflate.new(Zlib::DEFAULT_COMPRESSION, -Zlib::MAX_WBITS)
        begin
          deflater.deflate(content, Zlib::FINISH)
        ensure
          deflater.close
        end
      end

      def crc32(content)
        Zlib.crc32(content) & 0xFFFFFFFF
      end

      def dos_datetime(time)
        dos_date = ((time.year - 1980) << 9) | (time.month << 5) | time.day
        dos_time = (time.hour << 11) | (time.min << 5) | (time.sec / 2)
        le16(dos_time) + le16(dos_date)
      end

      def le16(value)
        [(value & 0xFF), ((value >> 8) & 0xFF)]
      end

      def le32(value)
        [(value & 0xFF), ((value >> 8) & 0xFF), ((value >> 16) & 0xFF), ((value >> 24) & 0xFF)]
      end

      def calculate_central_dir_size
        size = 0
        @entries.each do |entry|
          filename = entry[:path]
          size += 46 + filename.bytesize
        end
        size
      end
    end
  end
end
