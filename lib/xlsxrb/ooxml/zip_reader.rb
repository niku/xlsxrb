# frozen_string_literal: true

# rbs_inline: enabled

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
        
        io = if @io.is_a?(StringIO)
               @io
             elsif @io.respond_to?(:read) && @io.respond_to?(:seek) && @io.respond_to?(:pos)
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

        while true
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
            raw, comp_sz = find_data_descriptor_stream(io, method)
            entry_data = decompress(raw, method)
            result[entry_name] = entry_data unless entry_name.end_with?("/")
          else
            raw = io.read(compressed_size)
            entry_data = decompress(raw, method)
            result[entry_name] = entry_data unless entry_name.end_with?("/")
          end
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

              result << inflater.inflate(chunk)
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
          if desc_sig == [0x50, 0x4B, 0x07, 0x08].pack("C4")
            io.read(12)
          else
            io.seek(-4, IO::SEEK_CUR)
            io.read(12)
          end

          [result, consumed]
        else
          [io.read(0), 0]
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
