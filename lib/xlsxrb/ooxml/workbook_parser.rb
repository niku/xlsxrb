# frozen_string_literal: true

require_relative "xml_parser"

module Xlsxrb
  module Ooxml
    # SAX-based parser for xl/workbook.xml.
    # Returns sheet list: [{ name:, sheet_id:, r_id: }, ...].
    class WorkbookParser
      def self.parse(xml_string)
        return [] if xml_string.nil? || xml_string.empty?

        listener = Listener.new
        XmlParser.parse(xml_string, listener)
        listener.sheets
      end

      # SAX listener for workbook.xml sheets.
      class Listener
        include REXML::SAX2Listener

        attr_reader :sheets

        def initialize
          @sheets = []
        end

        def start_element(_uri, localname, _qname, attrs)
          return unless localname == "sheet"

          @sheets << {
            name: attrs["name"],
            sheet_id: attrs["sheetId"]&.to_i,
            r_id: attrs["r:id"] || attrs["id"] || attrs.find { |k, _| k.end_with?(":id") }&.last
          }
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
