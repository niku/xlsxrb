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
        if xml_string.nil? || xml_string.empty?
          return {
            sheets: [], date1904: false, defined_names: [], active_tab: nil,
            full_calc_on_load: nil, first_sheet: nil, show_sheet_tabs: nil,
            tab_ratio: nil, code_name: nil, workbook_protection: nil
          }
        end

        listener = Listener.new
        XmlParser.parse(xml_string, listener)
        {
          sheets: listener.sheets,
          date1904: listener.date1904,
          defined_names: listener.defined_names,
          active_tab: listener.active_tab,
          full_calc_on_load: listener.full_calc_on_load,
          first_sheet: listener.first_sheet,
          show_sheet_tabs: listener.show_sheet_tabs,
          tab_ratio: listener.tab_ratio,
          code_name: listener.code_name,
          workbook_protection: listener.workbook_protection
        }
      end

      # SAX listener for workbook.xml sheets, properties, and defined names.
      class Listener
        include REXML::SAX2Listener

        WORKBOOK_PROTECTION_ATTRIBUTES = {
          "workbookPassword" => %i[workbook_password string],
          "lockStructure" => %i[lock_structure bool],
          "lockWindows" => %i[lock_windows bool],
          "lockRevision" => %i[lock_revision bool],
          "revisionsPassword" => %i[revisions_password string],
          "workbookAlgorithmName" => %i[algorithm_name string],
          "workbookHashValue" => %i[hash_value string],
          "workbookSaltValue" => %i[salt_value string],
          "workbookSpinCount" => %i[spin_count int],
          "revisionsAlgorithmName" => %i[revisions_algorithm_name string],
          "revisionsHashValue" => %i[revisions_hash_value string],
          "revisionsSaltValue" => %i[revisions_salt_value string],
          "revisionsSpinCount" => %i[revisions_spin_count int]
        }.transform_values(&:freeze).freeze
        private_constant :WORKBOOK_PROTECTION_ATTRIBUTES

        attr_reader :sheets, :date1904, :defined_names, :active_tab, :full_calc_on_load, :first_sheet, :show_sheet_tabs, :tab_ratio, :code_name, :workbook_protection

        def initialize
          @sheets = []
          @date1904 = false
          @defined_names = []
          @active_tab = nil
          @full_calc_on_load = nil
          @first_sheet = nil
          @show_sheet_tabs = nil
          @tab_ratio = nil
          @code_name = nil
          @workbook_protection = nil
          @in_defined_name = false
          @current_dn = nil
          @dn_text = +""
        end

        def start_element(_uri, localname, _qname, attrs)
          case localname
          when "sheet"
            name_attr = attrs["name"]
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
            @code_name = attrs["codeName"] if attrs["codeName"]
          when "workbookView"
            at = attrs["activeTab"]
            @active_tab = at.to_i if at
            fs = attrs["firstSheet"]
            @first_sheet = fs.to_i if fs
            sst = attrs["showSheetTabs"]
            @show_sheet_tabs = %w[1 true].include?(sst) unless sst.nil?
            tr = attrs["tabRatio"]
            @tab_ratio = tr.to_i if tr
          when "calcPr"
            fcol = attrs["fullCalcOnLoad"]
            @full_calc_on_load = %w[1 true].include?(fcol) unless fcol.nil?
          when "workbookProtection"
            prot = {}
            attrs.each do |attr, value|
              key, kind = WORKBOOK_PROTECTION_ATTRIBUTES[attr]
              next unless key

              prot[key] = case kind
                          when :bool then %w[1 true].include?(value)
                          when :int then value.to_i
                          else value
                          end
            end
            @workbook_protection = prot unless prot.empty?
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

    # Parses .rels files to build rId -> target, type, and entry mappings.
    class RelationshipsParser
      EMPTY_PARSED_RELS = { entries: [].freeze, targets: {}.freeze, types: {}.freeze }.freeze
      private_constant :EMPTY_PARSED_RELS

      # Parses a .rels XML part in one pass, returning targets, types, and entry hashes.
      #: (String? xml_string) -> Hash[Symbol, untyped]
      def self.parse_all(xml_string)
        return EMPTY_PARSED_RELS if xml_string.nil? || xml_string.empty?

        sanitized = sanitize_xml_declarations(xml_string)
        listener = Listener.new
        begin
          XmlParser.parse(sanitized, listener)
        rescue StandardError
          # Fallback to regex scanning if SAX parser fails on malformed XML
          listener = Listener.new
          safe = sanitized.encoding == Encoding::UTF_8 ? sanitized : sanitized.dup.force_encoding(Encoding::UTF_8)
          safe = safe.scrub unless safe.valid_encoding?
          safe.scan(%r{<(?:[a-zA-Z0-9_]+:)?Relationship\b([^>]*)/?>}i) do |match|
            attrs_str = match[0]
            id = attrs_str[/\bId=["']([^"']+)["']/i, 1]
            type = attrs_str[/\bType=["']([^"']+)["']/i, 1]
            target = attrs_str[/\bTarget=["']([^"']+)["']/i, 1]
            listener.record(id, target, type)
          end
        end
        { entries: listener.entries, targets: listener.relationships, types: listener.types }
      end

      def self.parse(xml_string)
        parse_all(xml_string)[:targets]
      end

      # Parses .rels files to return array of relationship hashes with id, type, and target.
      def self.parse_entries(xml_string)
        parse_all(xml_string)[:entries]
      end

      #: (String xml_string) -> String
      def self.sanitize_xml_declarations(xml_string)
        sanitized = xml_string.lstrip
        return sanitized unless sanitized.start_with?("<?xml")

        first_decl_end = sanitized.index("?>")
        return sanitized unless first_decl_end

        decl = sanitized[0..(first_decl_end + 1)]
        rest = sanitized[(first_decl_end + 2)..].gsub(/<\?xml\b[^>]*\?>/i, "")
        "#{decl}#{rest}".lstrip
      rescue ArgumentError
        # Keep raw bytes unchanged if invalid byte sequences prevent declaration stripping.
        xml_string
      end
      private_class_method :sanitize_xml_declarations

      # SAX listener for .rels relationship files.
      class Listener
        include REXML::SAX2Listener

        attr_reader :relationships, :types, :entries

        def initialize
          @relationships = {}
          @types = {}
          @entries = []
        end

        def start_element(_uri, localname, _qname, attrs)
          return unless localname == "Relationship"

          record(attrs["Id"], attrs["Target"], attrs["Type"])
        end

        #: (String? rid, String? target, String? rel_type) -> void
        def record(rid, target, rel_type)
          return unless rid && target

          @relationships[rid] = target
          @types[rid] = rel_type if rel_type
          @entries << { id: rid, type: rel_type, target: target }
        end

        def end_element(_uri, _localname, _qname); end

        def characters(_text); end
      end
    end
  end
end
