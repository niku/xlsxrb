# frozen_string_literal: true

require_relative "xml_parser"

module Xlsxrb
  module Ooxml
    # SAX-based parser for xl/sharedStrings.xml.
    # Returns an Array of strings (index = SST index).
    class SharedStringsParser
      def self.parse(xml_string)
        return [] if xml_string.nil? || xml_string.empty?

        listener = Listener.new
        XmlParser.parse(xml_string, listener)
        listener.strings
      end

      # SAX listener for shared string table.
      class Listener
        include REXML::SAX2Listener

        attr_reader :strings

        def initialize
          @strings = []
          @in_si = false
          @in_t = false
          @current_text = +""
        end

        def start_element(_uri, localname, _qname, _attrs)
          case localname
          when "si"
            @in_si = true
            @current_text = +""
          when "t"
            @in_t = true
          end
        end

        def end_element(_uri, localname, _qname)
          case localname
          when "si"
            @in_si = false
            @strings << @current_text.freeze
          when "t"
            @in_t = false
          end
        end

        def characters(text)
          @current_text << text if @in_si && @in_t
        end
      end
    end
  end
end
