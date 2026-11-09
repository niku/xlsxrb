# frozen_string_literal: true

# rbs_inline: enabled

require "zlib"

module Xlsxrb
  module Ooxml
    # Writes ZIP archives using only stdlib (zlib).
    # Supports streaming: entries are written sequentially, central directory at close.
    #
    # Sizes, offsets, and entry counts that do not fit the 32-bit (16-bit for counts)
    # fields switch to ZIP64 records for that value only, so archives below those
    # limits are written exactly as plain ZIP.
    class ZipWriter
      MAX_UINT32 = 0xFFFF_FFFF
      MAX_UINT16 = 0xFFFF
      VERSION_DEFAULT = 20
      VERSION_ZIP64 = 45
      DATA_DESCRIPTOR_FLAG = 0x0008

      def self.open(target, compression_level: Zlib::DEFAULT_COMPRESSION, &block)
        io = target.is_a?(String) ? File.open(target, "wb") : target
        writer = new(io, compression_level: compression_level)
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

      # @return [Integer] Deflate compression level (0..9).
      attr_reader :compression_level

      # @param io [IO, StringIO] Destination IO stream.
      # @param compression_level [Integer, nil] ZIP deflate compression level (0..9, or Zlib::DEFAULT_COMPRESSION).
      #: (untyped io, ?compression_level: Integer | nil) -> void
      def initialize(io, compression_level: Zlib::DEFAULT_COMPRESSION)
        level = compression_level.nil? ? Zlib::DEFAULT_COMPRESSION : compression_level
        raise ArgumentError, "Invalid compression_level: #{compression_level.inspect}. Must be an Integer between 0 and 9 (or Zlib::DEFAULT_COMPRESSION)." unless level.is_a?(Integer) && (level == Zlib::DEFAULT_COMPRESSION || (0..9).cover?(level))

        @io = io
        @compression_level = level
        @entries = []
        @closed = false
        @bytes_written = 0
        @seekable = determine_seekable
        # Values at or above this limit are written as ZIP64. 0xFFFFFFFF itself is the
        # ZIP64 marker, so it cannot be stored in a 32-bit field.
        @zip64_limit = MAX_UINT32
      end

      # Add a file entry with string content.
      def add_entry(path, content)
        raise "ZipWriter is closed" if @closed

        content_bytes = content.encode("UTF-8").b
        write_raw_entry(path, content_bytes)
      end

      # Add a file entry with binary content (no encoding conversion).
      def add_binary_entry(path, content)
        raise "ZipWriter is closed" if @closed

        content_bytes = content.b
        write_raw_entry(path, content_bytes)
      end

      # Write a string directly into the current ZIP entry stream.
      # Use start_entry / write_data / finish_entry for true streaming.
      def start_entry(path)
        raise "ZipWriter is closed" if @closed

        @entry_path = path
        @entry_offset = @bytes_written
        @entry_gp_flag = @seekable ? 0 : DATA_DESCRIPTOR_FLAG
        @entry_deflater = Zlib::Deflate.new(@compression_level, -Zlib::MAX_WBITS)
        @entry_crc = Zlib.crc32
        @entry_uncompressed_size = 0
        @entry_compressed_size = 0
        @entry_seekable = @seekable
        @has_current_entry = true

        # Write local header (if unseekable, sizes/crc are 0 and bit 3 is set). Unseekable
        # output cannot patch the header once the size is known, and Excel reads the data
        # descriptor as 8-byte ZIP64 only when the local header states version 4.5, so such
        # entries are declared ZIP64 up front.
        write_local_header(path, 0, 0, 0, gp_flag: @entry_gp_flag, version: @seekable ? VERSION_DEFAULT : VERSION_ZIP64)
      end

      def write_data(str)
        raise "No entry started" unless @has_current_entry
        return if str.empty?

        @entry_crc = Zlib.crc32(str, @entry_crc)
        @entry_uncompressed_size += str.bytesize
        compressed = @entry_deflater.deflate(str)
        @entry_compressed_size += compressed.bytesize
        write_bytes(compressed)
      end

      def finish_entry
        raise "No entry started" unless @has_current_entry

        @has_current_entry = false

        # Flush remaining deflate data
        remaining = @entry_deflater.finish
        @entry_compressed_size += remaining.bytesize
        write_bytes(remaining)
        @entry_deflater.close

        final_crc = @entry_crc & 0xFFFFFFFF
        zip64 = !@entry_seekable || zip64?(@entry_compressed_size) || zip64?(@entry_uncompressed_size)

        if @entry_seekable && !zip64
          # Patch the local header if seekable
          patch_local_header { write_local_header(@entry_path, final_crc, @entry_compressed_size, @entry_uncompressed_size, gp_flag: 0, count_bytes: false) }
        else
          # Unseekable output always ends a streamed entry with an 8-byte data descriptor.
          # On seekable output, an entry that outgrew 32-bit sizes falls back to one too,
          # because the local header has no room for a ZIP64 extra field.
          write_data_descriptor(final_crc, @entry_compressed_size, @entry_uncompressed_size, zip64: zip64)
          if @entry_seekable
            @entry_gp_flag = DATA_DESCRIPTOR_FLAG
            patch_local_header { write_local_header(@entry_path, 0, 0, 0, gp_flag: @entry_gp_flag, count_bytes: false, version: VERSION_ZIP64) }
          end
        end

        @entries << {
          path: @entry_path,
          crc32: final_crc,
          compressed_size: @entry_compressed_size,
          uncompressed_size: @entry_uncompressed_size,
          offset: @entry_offset,
          gp_flag: @entry_gp_flag,
          method: 8,
          zip64_sizes: !@entry_seekable
        }
      end

      # Copies an entry from an existing archive IO without decompressing or recompressing.
      #: (String path, untyped source_io, data_offset: Integer, compressed_size: Integer, uncompressed_size: Integer, crc32: Integer, ?method: Integer) -> void
      def copy_raw_entry(path, source_io, data_offset:, compressed_size:, uncompressed_size:, crc32:, method: 8)
        raise "ZipWriter is closed" if @closed

        offset = @bytes_written
        write_local_header(path, crc32, compressed_size, uncompressed_size, gp_flag: 0, method: method)

        if compressed_size.positive?
          source_io.seek(data_offset, IO::SEEK_SET)
          IO.copy_stream(source_io, @io, compressed_size)
          @bytes_written += compressed_size
        end

        @entries << {
          path: path,
          crc32: crc32,
          compressed_size: compressed_size,
          uncompressed_size: uncompressed_size,
          offset: offset,
          gp_flag: 0,
          method: method
        }
      end

      def close
        return if @closed

        finish_entry if @current_entry

        @closed = true
        cd_offset = @bytes_written
        cd_size = write_central_directory
        write_end_of_central_directory(cd_offset, cd_size)
      end

      private

      def determine_seekable
        return true if @io.is_a?(StringIO)
        return false unless @io.respond_to?(:seek) && @io.respond_to?(:tell)

        begin
          @io.tell
          true
        rescue SystemCallError, IOError
          false
        end
      end

      #: (Integer value) -> bool
      def zip64?(value)
        value >= @zip64_limit
      end

      #: () { () -> void } -> void
      def patch_local_header
        current_pos = @io.is_a?(StringIO) ? @io.pos : @io.tell
        @io.seek(@entry_offset)
        yield
        @io.seek(current_pos)
      end

      # Data descriptor: signature + crc32 + compressed size + uncompressed size.
      # ZIP64 entries store both sizes as 8-byte values (APPNOTE 4.3.9.1).
      def write_data_descriptor(crc, compressed_size, uncompressed_size, zip64:)
        write_bytes([0x08074B50, crc, compressed_size, uncompressed_size].pack(zip64 ? "VVQ<Q<" : "VVVV"))
      end

      def write_bytes(data)
        return if data.nil? || data.empty?

        @io.write(data)
        @bytes_written += data.bytesize
      end

      def write_raw_entry(path, content_bytes)
        crc = Zlib.crc32(content_bytes) & 0xFFFFFFFF
        compressed = deflate(content_bytes)
        offset = @bytes_written

        write_local_header(path, crc, compressed.bytesize, content_bytes.bytesize, gp_flag: 0)
        write_bytes(compressed)

        @entries << {
          path: path,
          crc32: crc,
          compressed_size: compressed.bytesize,
          uncompressed_size: content_bytes.bytesize,
          offset: offset,
          gp_flag: 0,
          method: 8
        }
      end

      def deflate(content)
        deflater = Zlib::Deflate.new(@compression_level, -Zlib::MAX_WBITS)
        begin
          deflater.deflate(content, Zlib::FINISH)
        ensure
          deflater.close
        end
      end

      def write_local_header(path, crc, compressed_size, uncompressed_size, gp_flag: 0, count_bytes: true, method: 8, version: VERSION_DEFAULT)
        name_bytes = path.encode("UTF-8").b
        extra = "".b
        if zip64?(compressed_size) || zip64?(uncompressed_size)
          # A ZIP64 extra field in a local header must carry both sizes (APPNOTE 4.5.3).
          extra = [0x0001, 16, uncompressed_size, compressed_size].pack("vvQ<Q<")
          compressed_size = uncompressed_size = MAX_UINT32
          version = VERSION_ZIP64
        end
        header = [
          0x04034B50,        # local file header signature
          version,           # version needed (2.0, or 4.5 for ZIP64)
          gp_flag,           # general purpose bit flag (0x0008 if data descriptor follows)
          method,            # compression method (0 = store, 8 = deflate)
          0,                 # last mod file time
          33,                # last mod file date
          crc,
          compressed_size,
          uncompressed_size,
          name_bytes.bytesize,
          extra.bytesize     # extra field length
        ].pack("VvvvvvVVVvv")

        if count_bytes
          write_bytes(header)
          write_bytes(name_bytes)
          write_bytes(extra)
        else
          @io.write(header)
          @io.write(name_bytes)
          @io.write(extra)
        end
      end

      def write_central_directory
        size = 0
        @entries.each do |entry|
          name_bytes = entry[:path].encode("UTF-8").b
          uncompressed_size = entry[:uncompressed_size]
          compressed_size = entry[:compressed_size]
          offset = entry[:offset]
          # The ZIP64 extra field lists only the values whose 32-bit field holds
          # 0xFFFFFFFF, in this fixed order (APPNOTE 4.5.3).
          # Entries declared ZIP64 up front (streamed to unseekable output) list both sizes.
          zip64_sizes = entry[:zip64_sizes]
          usize64 = zip64_sizes || zip64?(uncompressed_size)
          csize64 = zip64_sizes || zip64?(compressed_size)
          zip64_values = []
          zip64_values << uncompressed_size if usize64
          zip64_values << compressed_size if csize64
          zip64_values << offset if zip64?(offset)
          extra = zip64_values.empty? ? "".b : [0x0001, zip64_values.size * 8, *zip64_values].pack("vvQ<*")
          version = zip64_values.empty? ? VERSION_DEFAULT : VERSION_ZIP64
          header = [
            0x02014B50,        # central directory file header signature
            version,           # version made by
            version,           # version needed
            entry[:gp_flag] || 0, # general purpose bit flag
            entry[:method] || 8, # compression method
            0,                 # last mod file time
            33,                # last mod file date
            entry[:crc32],
            csize64 ? MAX_UINT32 : compressed_size,
            usize64 ? MAX_UINT32 : uncompressed_size,
            name_bytes.bytesize,
            extra.bytesize,    # extra field length
            0,                 # file comment length
            0,                 # disk number start
            0,                 # internal file attributes
            0,                 # external file attributes
            zip64?(offset) ? MAX_UINT32 : offset # relative offset of local header
          ].pack("VvvvvvvVVVvvvvvVV")

          write_bytes(header)
          write_bytes(name_bytes)
          write_bytes(extra)
          size += header.bytesize + name_bytes.bytesize + extra.bytesize
        end
        size
      end

      def write_end_of_central_directory(cd_offset, cd_size)
        count = @entries.size
        zip64_count = count >= MAX_UINT16
        if zip64_count || zip64?(cd_size) || zip64?(cd_offset)
          zip64_eocd_offset = @bytes_written
          write_bytes([
            0x06064B50,        # ZIP64 end of central dir signature
            44,                # size of the remaining record
            VERSION_ZIP64,     # version made by
            VERSION_ZIP64,     # version needed
            0,                 # disk number
            0,                 # disk with central directory
            count,             # entries on this disk
            count,             # total entries
            cd_size,           # size of central directory
            cd_offset          # offset of central directory
          ].pack("VQ<vvVVQ<Q<Q<Q<"))
          write_bytes([
            0x07064B50,        # ZIP64 end of central dir locator signature
            0,                 # disk with ZIP64 end of central dir
            zip64_eocd_offset, # offset of ZIP64 end of central dir
            1                  # total number of disks
          ].pack("VVQ<V"))
        end

        eocd = [
          0x06054B50,          # end of central dir signature
          0,                   # disk number
          0,                   # disk with central directory
          zip64_count ? MAX_UINT16 : count, # entries on this disk
          zip64_count ? MAX_UINT16 : count, # total entries
          zip64?(cd_size) ? MAX_UINT32 : cd_size, # size of central directory
          zip64?(cd_offset) ? MAX_UINT32 : cd_offset, # offset of central directory
          0 # comment length
        ].pack("VvvvvVVv")
        write_bytes(eocd)
      end
    end
  end
end
