# frozen_string_literal: true

# rbs_inline: enabled

require_relative "xml_parser"

module Xlsxrb
  module Ooxml
    # SAX-based parser for xl/sharedStrings.xml.
    # Returns an Array of strings (index = SST index).
    class SharedStringsParser
      # Parses all shared strings and returns an Array of strings.
      def self.parse(xml_string, part_name: "xl/sharedStrings.xml")
        return [] if xml_string.nil? || xml_string.empty?

        strings = []
        each_event(xml_string, part_name: part_name) do |event|
          strings << event.args[0] if event.type == :sst_item
        end
        strings
      end

      # Yields Event objects for each shared string.
      def self.each_event(xml_string, part_name: "xl/sharedStrings.xml", &block)
        return enum_for(:each_event, xml_string, part_name: part_name) unless block
        return if xml_string.nil? || xml_string.empty?

        listener = EventListener.new(part_name, &block)
        XmlParser.parse(xml_string, listener)
      end

      # SAX listener for generating events from shared string table.
      class EventListener
        include REXML::SAX2Listener

        def initialize(part_name, &block)
          @part_name = part_name
          @block = block
          @in_si = false
          @in_t = false
          @current_text = +""
          @index = 0
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
            frozen_str = @current_text.freeze
            @block.call(Event.new(
                          type: :sst_item,
                          args: [frozen_str],
                          source: { part: @part_name, index: @index }
                        ))
            @index += 1
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
