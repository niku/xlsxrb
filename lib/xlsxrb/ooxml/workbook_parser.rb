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
        return { sheets: [], date1904: false, defined_names: [], active_tab: nil, full_calc_on_load: nil } if xml_string.nil? || xml_string.empty?

        listener = Listener.new
        XmlParser.parse(xml_string, listener)
        {
          sheets: listener.sheets,
          date1904: listener.date1904,
          defined_names: listener.defined_names,
          active_tab: listener.active_tab,
          full_calc_on_load: listener.full_calc_on_load
        }
      end

      # SAX listener for workbook.xml sheets, properties, and defined names.
      class Listener
        include REXML::SAX2Listener

        attr_reader :sheets, :date1904, :defined_names, :active_tab, :full_calc_on_load

        def initialize
          @sheets = []
          @date1904 = false
          @defined_names = []
          @active_tab = nil
          @full_calc_on_load = nil
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
          when "workbookView"
            at = attrs["activeTab"]
            @active_tab = at.to_i if at
          when "calcPr"
            fcol = attrs["fullCalcOnLoad"]
            @full_calc_on_load = %w[1 true].include?(fcol) unless fcol.nil?
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

        sanitized = xml_string.lstrip
        if sanitized.start_with?("<?xml")
          first_decl_end = sanitized.index("?>")
          if first_decl_end
            decl = sanitized[0..(first_decl_end + 1)]
            rest = sanitized[(first_decl_end + 2)..].gsub(/<\?xml\b[^>]*\?>/i, "")
            sanitized = "#{decl}#{rest}".lstrip
          end
        end

        listener = Listener.new
        begin
          XmlParser.parse(sanitized, listener)
        rescue StandardError
          # Fallback to regex scanning if SAX parser fails on malformed XML
          sanitized.scan(%r{<(?:[a-zA-Z0-9_]+:)?Relationship\b([^>]*)/?>}i) do |match|
            attrs_str = match[0]
            id = attrs_str[/\bId=["']([^"']+)["']/i, 1]
            target = attrs_str[/\bTarget=["']([^"']+)["']/i, 1]
            listener.relationships[id] = target if id && target
          end
        end
        listener.relationships
      end

      # Parses .rels files to return array of relationship hashes with id, type, and target.
      def self.parse_entries(xml_string)
        return [] if xml_string.nil? || xml_string.empty?

        sanitized = xml_string.lstrip
        if sanitized.start_with?("<?xml")
          first_decl_end = sanitized.index("?>")
          if first_decl_end
            decl = sanitized[0..(first_decl_end + 1)]
            rest = sanitized[(first_decl_end + 2)..].gsub(/<\?xml\b[^>]*\?>/i, "")
            sanitized = "#{decl}#{rest}".lstrip
          end
        end

        entries = []
        sanitized.scan(%r{<(?:[a-zA-Z0-9_]+:)?Relationship\b([^>]*)/?>}i) do |match|
          attrs_str = match[0]
          id = attrs_str[/\bId=["']([^"']+)["']/i, 1]
          type = attrs_str[/\bType=["']([^"']+)["']/i, 1]
          target = attrs_str[/\bTarget=["']([^"']+)["']/i, 1]
          entries << { id: id, type: type, target: target } if id && target
        end
        entries
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
