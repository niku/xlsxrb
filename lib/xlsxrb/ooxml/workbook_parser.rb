# frozen_string_literal: true

# rbs_inline: enabled

require_relative "xml_parser"

module Xlsxrb
  module Ooxml
    # SAX-based parser for xl/workbook.xml.
    # Returns sheet list: [{ name:, sheet_id:, r_id: }, ...].
    class WorkbookParser
      def self.parse(xml_string)
        parse_with_properties(xml_string)[:sheets]
      end

      def self.parse_with_properties(xml_string)
        return { sheets: [], date1904: false } if xml_string.nil? || xml_string.empty?

        listener = Listener.new
        XmlParser.parse(xml_string, listener)
        { sheets: listener.sheets, date1904: listener.date1904 }
      end

      # SAX listener for workbook.xml sheets and properties.
      class Listener
        include REXML::SAX2Listener

        attr_reader :sheets, :date1904

        def initialize
          @sheets = []
          @date1904 = false
        end

        def start_element(_uri, localname, _qname, attrs)
          case localname
          when "sheet"
            name_attr = attrs["name"]
            name_attr = REXML::Text.unnormalize(name_attr) if name_attr

            @sheets << {
              name: name_attr,
              sheet_id: attrs["sheetId"]&.to_i,
              r_id: attrs["r:id"] || attrs["id"] || attrs.find { |k, _| k.end_with?(":id") }&.last
            }
          when "workbookPr"
            d1904 = attrs["date1904"]
            @date1904 = %w[1 true].include?(d1904) unless d1904.nil?
          end
        end

        def end_element(_uri, _localname, _qname); end

        def characters(_text); end
      end
    end

    # Parses .rels files to build rId -> target mapping.
    class RelationshipsParser
      def self.parse(xml_string)
        return {} if xml_string.nil? || xml_string.empty?

        listener = Listener.new
        XmlParser.parse(xml_string, listener)
        listener.relationships
      end

      # SAX listener for .rels relationship files.
      class Listener
        include REXML::SAX2Listener

        attr_reader :relationships

        def initialize
          @relationships = {}
        end

        def start_element(_uri, localname, _qname, attrs)
          return unless localname == "Relationship"

          rid = attrs["Id"]
          target = attrs["Target"]
          @relationships[rid] = target if rid && target
        end

        def end_element(_uri, _localname, _qname); end

        def characters(_text); end
      end
    end
  end
end
