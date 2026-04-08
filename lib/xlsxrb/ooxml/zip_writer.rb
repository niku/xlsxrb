# frozen_string_literal: true

require "zlib"

module Xlsxrb
  module Ooxml
    # Writes ZIP archives using only stdlib (zlib).
    # Supports streaming: entries are written sequentially, central directory at close.
    class ZipWriter
      def self.open(target, &block)
        io = target.is_a?(String) ? File.open(target, "wb") : target
        writer = new(io)
        if block
          begin
            yield writer
          ensure
            writer.close
            io.close if target.is_a?(String)
          end
        else
          writer
        end
      end

      def initialize(io)
        @io = io
        @entries = []
        @closed = false
      end

      # Add a file entry with string content.
      def add_entry(path, content)
        raise "ZipWriter is closed" if @closed

        content_bytes = content.encode("UTF-8").b
        crc = Zlib.crc32(content_bytes) & 0xFFFFFFFF
        compressed = deflate(content_bytes)

        offset = @io.is_a?(StringIO) ? @io.pos : @io.tell

        write_local_header(path, crc, compressed.bytesize, content_bytes.bytesize)
        @io.write(compressed)

        @entries << {
          path: path,
          crc32: crc,
          compressed_size: compressed.bytesize,
          uncompressed_size: content_bytes.bytesize,
          offset: offset
        }
      end

      # Write a string directly into the current ZIP entry stream.
      # Use start_entry / write_data / finish_entry for true streaming.
      def start_entry(path)
        raise "ZipWriter is closed" if @closed

        @current_entry = {
          path: path,
          offset: @io.is_a?(StringIO) ? @io.pos : @io.tell,
          deflater: Zlib::Deflate.new(Zlib::DEFAULT_COMPRESSION, -Zlib::MAX_WBITS),
          crc: Zlib.crc32,
          uncompressed_size: 0,
          compressed_size: 0
        }

        # Write a placeholder local header (will be patched later if IO supports seek)
        write_local_header(path, 0, 0, 0)
      end

      def write_data(str)
        raise "No entry started" unless @current_entry

        bytes = str.b
        @current_entry[:crc] = Zlib.crc32(bytes, @current_entry[:crc])
        @current_entry[:uncompressed_size] += bytes.bytesize
        compressed = @current_entry[:deflater].deflate(bytes, Zlib::SYNC_FLUSH)
        @current_entry[:compressed_size] += compressed.bytesize
        @io.write(compressed)
      end

      def finish_entry
        raise "No entry started" unless @current_entry

        entry = @current_entry
        @current_entry = nil

        # Flush remaining deflate data
        remaining = entry[:deflater].finish
        entry[:compressed_size] += remaining.bytesize
        @io.write(remaining)
        entry[:deflater].close

        final_crc = entry[:crc] & 0xFFFFFFFF

        # Patch the local header if seekable
        current_pos = @io.is_a?(StringIO) ? @io.pos : @io.tell
        if @io.respond_to?(:seek)
          @io.seek(entry[:offset])
          write_local_header(entry[:path], final_crc, entry[:compressed_size], entry[:uncompressed_size])
          @io.seek(current_pos)
        end

        @entries << {
          path: entry[:path],
          crc32: final_crc,
          compressed_size: entry[:compressed_size],
          uncompressed_size: entry[:uncompressed_size],
          offset: entry[:offset]
        }
      end

      def close
        return if @closed

        finish_entry if @current_entry

        @closed = true
        cd_offset = @io.is_a?(StringIO) ? @io.pos : @io.tell
        cd_size = write_central_directory
        write_end_of_central_directory(cd_offset, cd_size)
      end

      private

      def deflate(content)
        deflater = Zlib::Deflate.new(Zlib::DEFAULT_COMPRESSION, -Zlib::MAX_WBITS)
        begin
          deflater.deflate(content, Zlib::FINISH)
        ensure
          deflater.close
        end
      end

      def write_local_header(path, crc, compressed_size, uncompressed_size)
        name_bytes = path.encode("UTF-8").b
        header = [
          0x04034B50,        # local file header signature
          20,                # version needed (2.0)
          0,                 # general purpose bit flag
          8,                 # compression method (deflate)
          0,                 # last mod file time
          0,                 # last mod file date
          crc,
          compressed_size,
          uncompressed_size,
          name_bytes.bytesize,
          0                  # extra field length
        ].pack("VvvvvvVVVvv")
        @io.write(header)
        @io.write(name_bytes)
      end

      def write_central_directory
        size = 0
        @entries.each do |entry|
          name_bytes = entry[:path].encode("UTF-8").b
          header = [
            0x02014B50,        # central directory file header signature
            20,                # version made by
            20,                # version needed
            0,                 # general purpose bit flag
            8,                 # compression method
            0,                 # last mod file time
            0,                 # last mod file date
            entry[:crc32],
            entry[:compressed_size],
            entry[:uncompressed_size],
            name_bytes.bytesize,
            0,                 # extra field length
            0,                 # file comment length
            0,                 # disk number start
            0,                 # internal file attributes
            0,                 # external file attributes
            entry[:offset]     # relative offset of local header
          ].pack("VvvvvvvVVVvvvvvVV")
          @io.write(header)
          @io.write(name_bytes)
          size += header.bytesize + name_bytes.bytesize
        end
        size
      end

      def write_end_of_central_directory(cd_offset, cd_size)
        eocd = [
          0x06054B50,          # end of central dir signature
          0,                   # disk number
          0,                   # disk with central directory
          @entries.size,       # entries on this disk
          @entries.size,       # total entries
          cd_size,             # size of central directory
          cd_offset,           # offset of central directory
          0                    # comment length
        ].pack("VvvvvVVv")
        @io.write(eocd)
      end
    end
  end
end
