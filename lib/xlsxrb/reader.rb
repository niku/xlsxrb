# frozen_string_literal: true

require "zlib"
require "rexml/parsers/sax2parser"
require "rexml/sax2listener"

module Xlsxrb
  # Reads inline string cells from an XLSX file.
  class Reader
    def initialize(filepath)
      @filepath = filepath
    end

    def cells
      worksheet_xml = extract_zip_entry("xl/worksheets/sheet1.xml")
      return {} if worksheet_xml.nil? || worksheet_xml.empty?

      parse_inline_string_cells(worksheet_xml)
    end

    private

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

    def parse_inline_string_cells(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = InlineStringWorksheetListener.new
      parser.listen(listener)
      parser.parse
      listener.cells
    end

    # SAX2 listener for parsing inline string cells from worksheet XML.
    class InlineStringWorksheetListener
      include REXML::SAX2Listener

      attr_reader :cells

      def initialize
        @cells = {}
        @current_cell_ref = nil
        @current_cell_type = nil
        @inside_cell_text = false
        @text_buffer = +""
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)

        case name
        when "c"
          @current_cell_ref = attributes["r"]
          @current_cell_type = attributes["t"]
          @text_buffer = +""
        when "t"
          @inside_cell_text = @current_cell_type == "inlineStr" && !@current_cell_ref.nil?
        end
      end

      def characters(text)
        @text_buffer << text if @inside_cell_text
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)

        case name
        when "t"
          @inside_cell_text = false
        when "c"
          @cells[@current_cell_ref] = @text_buffer.dup if @current_cell_type == "inlineStr" && !@current_cell_ref.nil?

          @current_cell_ref = nil
          @current_cell_type = nil
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
  end
end
