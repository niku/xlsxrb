# frozen_string_literal: true

require "zlib"
require "rexml/parsers/sax2parser"
require "rexml/sax2listener"

module Xlsxrb
  # Reads cells from an XLSX file.
  class Reader
    def initialize(filepath)
      @filepath = filepath
    end

    def cells
      worksheet_xml = extract_zip_entry("xl/worksheets/sheet1.xml")
      return {} if worksheet_xml.nil? || worksheet_xml.empty?

      shared_strings = load_shared_strings
      parse_worksheet_cells(worksheet_xml, shared_strings)
    end

    private

    def load_shared_strings
      sst_xml = extract_zip_entry("xl/sharedStrings.xml")
      return [] if sst_xml.nil? || sst_xml.empty?

      parser = REXML::Parsers::SAX2Parser.new(sst_xml)
      listener = SharedStringsListener.new
      parser.listen(listener)
      parser.parse
      listener.strings
    end

    def extract_zip_entry(entry_name)
      File.open(@filepath, "rb") do |file|
        loop do
          signature = file.read(4)
          break if signature.nil? || signature.bytesize < 4

          signature_value = signature.unpack1("V")
          break if [0x02014b50, 0x06054b50].include?(signature_value)

          raise Error, "invalid ZIP local header signature" unless signature_value == 0x04034b50

          header = file.read(26)
          raise Error, "truncated ZIP local header" if header.nil? || header.bytesize < 26

          _version, flags, compression_method, _mtime, _mdate, _crc32, compressed_size,
            _uncompressed_size, file_name_length, extra_field_length = header.unpack("v v v v v V V V v v")

          raise Error, "ZIP data descriptor is not supported" if flags.anybits?(0x0008)

          file_name = file.read(file_name_length)
          raise Error, "truncated ZIP file name" if file_name.nil? || file_name.bytesize < file_name_length

          file.read(extra_field_length)

          compressed_data = file.read(compressed_size)
          raise Error, "truncated ZIP entry data" if compressed_data.nil? || compressed_data.bytesize < compressed_size

          next unless file_name == entry_name

          case compression_method
          when 0
            return compressed_data
          when 8
            inflater = Zlib::Inflate.new(-Zlib::MAX_WBITS)
            begin
              return inflater.inflate(compressed_data)
            ensure
              inflater.close
            end
          else
            raise Error, "unsupported ZIP compression method: #{compression_method}"
          end
        end
      end

      nil
    end

    def parse_worksheet_cells(xml, shared_strings)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = WorksheetListener.new(shared_strings)
      parser.listen(listener)
      parser.parse
      listener.cells
    end

    # SAX2 listener for parsing shared string table (xl/sharedStrings.xml).
    class SharedStringsListener
      include REXML::SAX2Listener

      attr_reader :strings

      def initialize
        @strings = []
        @inside_si = false
        @inside_t = false
        @text_buffer = +""
      end

      def start_element(_uri, local_name, qname, _attributes)
        name = element_name(local_name, qname)

        case name
        when "si"
          @inside_si = true
          @text_buffer = +""
        when "t"
          @inside_t = @inside_si
        end
      end

      def characters(text)
        @text_buffer << text if @inside_t
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)

        case name
        when "t"
          @inside_t = false
        when "si"
          @strings << @text_buffer.dup
          @inside_si = false
          @text_buffer = +""
        end
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing worksheet cells.
    class WorksheetListener
      include REXML::SAX2Listener

      attr_reader :cells

      def initialize(shared_strings = [])
        @shared_strings = shared_strings
        @cells = {}
        @current_cell_ref = nil
        @current_cell_type = nil
        @inside_value = false
        @inside_inline_text = false
        @value_buffer = +""
        @inline_text_buffer = +""
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)

        case name
        when "c"
          @current_cell_ref = attributes["r"]
          @current_cell_type = attributes["t"]
          @value_buffer = +""
          @inline_text_buffer = +""
        when "v"
          @inside_value = true
        when "t"
          @inside_inline_text = @current_cell_type == "inlineStr" && !@current_cell_ref.nil?
        end
      end

      def characters(text)
        @value_buffer << text if @inside_value
        @inline_text_buffer << text if @inside_inline_text
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)

        case name
        when "v"
          @inside_value = false
        when "t"
          @inside_inline_text = false
        when "c"
          store_cell_value
          @current_cell_ref = nil
          @current_cell_type = nil
          @value_buffer = +""
          @inline_text_buffer = +""
        end
      end

      private

      def store_cell_value
        return if @current_cell_ref.nil?

        case @current_cell_type
        when "inlineStr"
          @cells[@current_cell_ref] = @inline_text_buffer.dup
        when "s"
          index = @value_buffer.to_i
          @cells[@current_cell_ref] = @shared_strings[index] || ""
        when "e"
          @cells[@current_cell_ref] = @value_buffer.dup
        when "b"
          @cells[@current_cell_ref] = @value_buffer.strip == "1"
        when nil, "", "n"
          return if @value_buffer.empty?

          raw = @value_buffer.dup
          @cells[@current_cell_ref] = numeric_value(raw)
        end
      end

      def numeric_value(raw)
        if raw.include?(".")
          raw.to_f
        else
          raw.to_i
        end
      end

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end
  end
end
