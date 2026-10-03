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
        return { sheets: [], date1904: false, defined_names: [] } if xml_string.nil? || xml_string.empty?

        listener = Listener.new
        XmlParser.parse(xml_string, listener)
        { sheets: listener.sheets, date1904: listener.date1904, defined_names: listener.defined_names }
      end

      # SAX listener for workbook.xml sheets, properties, and defined names.
      class Listener
        include REXML::SAX2Listener

        attr_reader :sheets, :date1904, :defined_names

        def initialize
          @sheets = []
          @date1904 = false
          @defined_names = []
          @in_defined_name = false
          @current_dn = nil
          @dn_text = +""
        end

        def start_element(_uri, localname, _qname, attrs)
          case localname
          when "sheet"
            name_attr = attrs["name"]
            name_attr = REXML::Text.unnormalize(name_attr) if name_attr
            raw_state = attrs["state"]
            state_sym = case raw_state
                        when "hidden" then :hidden
                        when "veryHidden" then :very_hidden
                        else :visible
                        end

            @sheets << {
              name: name_attr,
              sheet_id: attrs["sheetId"]&.to_i,
              r_id: attrs["r:id"] || attrs["id"] || attrs.find { |k, _| k.end_with?(":id") }&.last,
              state: state_sym
            }
          when "workbookPr"
            d1904 = attrs["date1904"]
            @date1904 = %w[1 true].include?(d1904) unless d1904.nil?
          when "definedName"
            @in_defined_name = true
            @current_dn = {
              name: attrs["name"],
              hidden: %w[1 true].include?(attrs["hidden"])
            }
            @current_dn[:local_sheet_id] = attrs["localSheetId"].to_i if attrs["localSheetId"]
            @current_dn[:comment] = attrs["comment"] if attrs["comment"]
            @dn_text = +""
          end
        end

        def characters(text)
          @dn_text << text if @in_defined_name
        end

        def end_element(_uri, localname, _qname)
          return unless localname == "definedName" && @in_defined_name

          @current_dn[:value] = @dn_text.strip
          @defined_names << @current_dn
          @in_defined_name = false
          @current_dn = nil
        end
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
