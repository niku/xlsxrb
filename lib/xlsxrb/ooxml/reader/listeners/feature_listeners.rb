# frozen_string_literal: true

# rbs_inline: enabled

require "rexml/parsers/sax2parser"
require "rexml/sax2listener"

module Xlsxrb
  module Ooxml
    class Reader
      # SAX2 listener for parsing <autoFilter> element.
      class AutoFilterListener
        include REXML::SAX2Listener

        attr_reader :ref, :filter_columns

        def initialize
          @ref = nil
          @filter_columns = {}
          @current_col_id = nil
          @current_filter = nil
          @inside_custom_filters = false
          @custom_filters_list = []
          @custom_filters_and = false
          @filter_values = []
          @filter_blank = false
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          case name
          when "autoFilter"
            @ref = attributes["ref"]
          when "filterColumn"
            @current_col_id = attributes["colId"]&.to_i
            @fc_hidden_button = attributes["hiddenButton"] == "1"
            @fc_show_button = attributes["showButton"] == "0" ? false : nil
          when "filters"
            @filter_blank = attributes["blank"] == "1"
            @filter_calendar_type = attributes["calendarType"]
            @filter_values = []
            @date_group_items = []
          when "filter"
            val = attributes["val"]
            @filter_values << val if val
          when "dateGroupItem"
            dg = { date_time_grouping: attributes["dateTimeGrouping"] }
            dg[:year] = attributes["year"].to_i if attributes["year"]
            dg[:month] = attributes["month"].to_i if attributes["month"]
            dg[:day] = attributes["day"].to_i if attributes["day"]
            dg[:hour] = attributes["hour"].to_i if attributes["hour"]
            dg[:minute] = attributes["minute"].to_i if attributes["minute"]
            dg[:second] = attributes["second"].to_i if attributes["second"]
            @date_group_items << dg
          when "customFilters"
            @inside_custom_filters = true
            @custom_filters_and = attributes["and"] == "1"
            @custom_filters_list = []
          when "customFilter"
            @custom_filters_list << { operator: attributes["operator"], val: attributes["val"] } if @inside_custom_filters
          when "dynamicFilter"
            df = { type: :dynamic, dynamic_type: attributes["type"] }
            df[:val] = attributes["val"].to_f if attributes["val"]
            df[:val_iso] = attributes["valIso"] if attributes["valIso"]
            df[:max_val] = attributes["maxVal"].to_f if attributes["maxVal"]
            df[:max_val_iso] = attributes["maxValIso"] if attributes["maxValIso"]
            @current_filter = df
          when "top10"
            t10 = {
              type: :top10,
              top: attributes["top"] == "1",
              percent: attributes["percent"] == "1",
              val: attributes["val"]&.to_f&.to_i
            }
            t10[:filter_val] = attributes["filterVal"].to_f if attributes["filterVal"]
            @current_filter = t10
          when "colorFilter"
            cf = { type: :color_filter, dxf_id: attributes["dxfId"]&.to_i }
            cf[:cell_color] = false if attributes["cellColor"] == "0"
            @current_filter = cf
          when "iconFilter"
            icf = { type: :icon_filter, icon_set: attributes["iconSet"] }
            icf[:icon_id] = attributes["iconId"].to_i if attributes["iconId"]
            @current_filter = icf
          end
        end

        def end_element(_uri, local_name, qname)
          name = element_name(local_name, qname)
          case name
          when "filters"
            f = { type: :filters }
            f[:blank] = true if @filter_blank
            f[:calendar_type] = @filter_calendar_type if @filter_calendar_type
            f[:values] = @filter_values unless @filter_values.empty?
            f[:date_group_items] = @date_group_items unless @date_group_items.empty?
            @current_filter = f
          when "customFilters"
            if @custom_filters_list.size == 1
              cf = @custom_filters_list.first
              @current_filter = { type: :custom, operator: cf[:operator], val: cf[:val] }
            else
              @current_filter = { type: :custom, filters: @custom_filters_list, and: @custom_filters_and }
            end
            @inside_custom_filters = false
          when "filterColumn"
            if @current_col_id && @current_filter
              @current_filter[:hidden_button] = true if @fc_hidden_button
              @current_filter[:show_button] = false if @fc_show_button == false
              @filter_columns[@current_col_id] = @current_filter
            end
            @current_col_id = nil
            @current_filter = nil
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

      # SAX2 listener for parsing <sortState> element.
      class SortStateListener
        include REXML::SAX2Listener

        attr_reader :sort_state

        def initialize
          @sort_state = nil
          @inside_sort_state = false
          @sort_conditions = []
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          case name
          when "sortState"
            @inside_sort_state = true
            ss = { ref: attributes["ref"], sort_conditions: [] }
            ss[:column_sort] = true if attributes["columnSort"] == "1"
            ss[:case_sensitive] = true if attributes["caseSensitive"] == "1"
            ss[:sort_method] = attributes["sortMethod"] if attributes["sortMethod"]
            @sort_state = ss
          when "sortCondition"
            return unless @inside_sort_state

            sc = { ref: attributes["ref"] }
            sc[:descending] = true if attributes["descending"] == "1"
            sc[:sort_by] = attributes["sortBy"] if attributes["sortBy"]
            sc[:custom_list] = attributes["customList"] if attributes["customList"]
            dxf = attributes["dxfId"]
            sc[:dxf_id] = dxf.to_i if dxf
            sc[:icon_set] = attributes["iconSet"] if attributes["iconSet"]
            iid = attributes["iconId"]
            sc[:icon_id] = iid.to_i if iid
            @sort_conditions << sc
          end
        end

        def end_element(_uri, local_name, qname)
          name = element_name(local_name, qname)
          return unless name == "sortState" && @inside_sort_state

          @sort_state[:sort_conditions] = @sort_conditions
          @inside_sort_state = false
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

      # SAX2 listener for parsing docProps/app.xml.
      class AppPropertiesListener
        include REXML::SAX2Listener

        attr_reader :properties

        def initialize
          @properties = {}
          @current_field = nil
          @text_buffer = +""
          @inside_vector = false
          @vector_items = []
          @heading_pairs = []
          @titles_of_parts = []
          @inside_heading_pairs = false
          @inside_titles_of_parts = false
          @inside_variant = false
        end

        def start_element(_uri, local_name, qname, _attributes)
          name = element_name(local_name, qname)
          case name
          when "Application", "AppVersion"
            @current_field = name
            @text_buffer = +""
          when "HeadingPairs"
            @inside_heading_pairs = true
          when "TitlesOfParts"
            @inside_titles_of_parts = true
          when "variant"
            @inside_variant = true
            @text_buffer = +""
          when "lpstr", "i4"
            @text_buffer = +""
          end
        end

        def characters(text)
          @text_buffer << text
        end

        def end_element(_uri, local_name, qname)
          name = element_name(local_name, qname)
          case name
          when "Application"
            @properties[:application] = @text_buffer.dup
            @current_field = nil
          when "AppVersion"
            @properties[:app_version] = @text_buffer.dup
            @current_field = nil
          when "lpstr"
            if @inside_titles_of_parts
              @titles_of_parts << @text_buffer.dup
            elsif @inside_heading_pairs && @inside_variant
              @vector_items << @text_buffer.dup
            end
          when "i4"
            @vector_items << @text_buffer.to_i if @inside_heading_pairs && @inside_variant
          when "variant"
            @inside_variant = false
          when "HeadingPairs"
            # Convert flat array to pairs: [label, count, label, count, ...]
            @heading_pairs = @vector_items.each_slice(2).to_a
            @vector_items = []
            @inside_heading_pairs = false
            @properties[:heading_pairs] = @heading_pairs
          when "TitlesOfParts"
            @inside_titles_of_parts = false
            @properties[:titles_of_parts] = @titles_of_parts
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

      # SAX2 listener for parsing docProps/core.xml.
      class CorePropertiesListener
        include REXML::SAX2Listener

        attr_reader :properties

        FIELD_MAP = {
          "title" => :title,
          "subject" => :subject,
          "creator" => :creator,
          "keywords" => :keywords,
          "description" => :description,
          "lastModifiedBy" => :last_modified_by,
          "revision" => :revision,
          "created" => :created,
          "modified" => :modified,
          "category" => :category,
          "contentStatus" => :content_status,
          "language" => :language
        }.freeze

        def initialize
          @properties = {}
          @current_field = nil
          @text_buffer = +""
        end

        def start_element(_uri, local_name, qname, _attributes)
          name = element_name(local_name, qname)
          return unless FIELD_MAP.key?(name)

          @current_field = FIELD_MAP[name]
          @text_buffer = +""
        end

        def characters(text)
          @text_buffer << text if @current_field
        end

        def end_element(_uri, local_name, qname)
          name = element_name(local_name, qname)
          return unless @current_field && FIELD_MAP.key?(name)

          @properties[@current_field] = @text_buffer.dup unless @text_buffer.empty?
          @current_field = nil
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

      # SAX2 listener for parsing custom properties (docProps/custom.xml).
      class CustomPropertiesListener
        include REXML::SAX2Listener

        attr_reader :properties

        def initialize
          @properties = []
          @current_name = nil
          @current_type = nil
          @text_buffer = +""
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          case name
          when "property"
            @current_name = attributes["name"]
          when "lpwstr"
            @current_type = :string
            @text_buffer = +""
          when "i4"
            @current_type = :number
            @text_buffer = +""
          when "r8"
            @current_type = :float
            @text_buffer = +""
          when "bool"
            @current_type = :bool
            @text_buffer = +""
          when "filetime"
            @current_type = :date
            @text_buffer = +""
          end
        end

        def characters(text)
          @text_buffer << text if @current_type
        end

        def end_element(_uri, local_name, qname)
          name = element_name(local_name, qname)
          case name
          when "property"
            @current_name = nil
          when "lpwstr", "i4", "r8", "bool", "filetime"
            if @current_name
              value = case @current_type
                      when :number then @text_buffer.to_i
                      when :float then @text_buffer.to_f
                      when :bool then @text_buffer == "true"
                      else @text_buffer.dup
                      end
              @properties << { name: @current_name, value: value, type: @current_type }
            end
            @current_type = nil
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

      # SAX2 listener for parsing <sheetPr> element (tabColor, outlinePr).
      class SheetPropertiesListener
        include REXML::SAX2Listener

        attr_reader :properties

        def initialize
          @properties = {}
          @inside_sheet_pr = false
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          case name
          when "sheetPr"
            @inside_sheet_pr = true
            sh = attributes["syncHorizontal"]
            @properties[:sync_horizontal] = %w[1 true].include?(sh) unless sh.nil?
            sv = attributes["syncVertical"]
            @properties[:sync_vertical] = %w[1 true].include?(sv) unless sv.nil?
            @properties[:sync_ref] = attributes["syncRef"] if attributes["syncRef"]
            te = attributes["transitionEvaluation"]
            @properties[:transition_evaluation] = %w[1 true].include?(te) unless te.nil?
            tent = attributes["transitionEntry"]
            @properties[:transition_entry] = %w[1 true].include?(tent) unless tent.nil?
            @properties[:code_name] = attributes["codeName"] if attributes["codeName"]
            fm = attributes["filterMode"]
            @properties[:filter_mode] = %w[1 true].include?(fm) unless fm.nil?
            pub = attributes["published"]
            @properties[:published] = %w[1 true].include?(pub) unless pub.nil?
            efcc = attributes["enableFormatConditionsCalculation"]
            @properties[:enable_format_conditions_calculation] = %w[1 true].include?(efcc) unless efcc.nil?
          when "tabColor"
            if @inside_sheet_pr
              @properties[:tab_color] = attributes["rgb"] if attributes["rgb"]
              @properties[:tab_color_theme] = attributes["theme"].to_i if attributes["theme"]
              @properties[:tab_color_tint] = attributes["tint"].to_f if attributes["tint"]
              @properties[:tab_color_indexed] = attributes["indexed"].to_i if attributes["indexed"]
              @properties[:tab_color_auto] = %w[1 true].include?(attributes["auto"]) if attributes["auto"]
            end
          when "outlinePr"
            if @inside_sheet_pr
              apply_s = attributes["applyStyles"]
              @properties[:apply_styles] = %w[1 true].include?(apply_s) unless apply_s.nil?
              sb = attributes["summaryBelow"]
              @properties[:summary_below] = %w[1 true].include?(sb) unless sb.nil?
              sr = attributes["summaryRight"]
              @properties[:summary_right] = %w[1 true].include?(sr) unless sr.nil?
              sos = attributes["showOutlineSymbols"]
              @properties[:show_outline_symbols] = %w[1 true].include?(sos) unless sos.nil?
            end
          when "pageSetUpPr"
            if @inside_sheet_pr
              ftp = attributes["fitToPage"]
              @properties[:fit_to_page] = %w[1 true].include?(ftp) unless ftp.nil?
              apb = attributes["autoPageBreaks"]
              @properties[:auto_page_breaks] = %w[1 true].include?(apb) unless apb.nil?
            end
          when "sheetCalcPr"
            fcol = attributes["fullCalcOnLoad"]
            @properties[:full_calc_on_load] = %w[1 true].include?(fcol) unless fcol.nil?
          end
        end

        def end_element(_uri, local_name, qname)
          name = element_name(local_name, qname)
          @inside_sheet_pr = false if name == "sheetPr"
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

      # SAX2 listener for parsing <sheetProtection> element.
      class SheetProtectionListener
        include REXML::SAX2Listener

        attr_reader :protection

        def initialize
          @protection = nil
        end

        BOOL_ATTRS = %i[sheet objects scenarios select_locked_cells select_unlocked_cells].freeze
        FALSE_ATTRS = %i[format_cells format_columns format_rows insert_columns insert_rows
                         insert_hyperlinks delete_columns delete_rows sort auto_filter pivot_tables].freeze
        ATTR_MAP = {
          "sheet" => :sheet, "objects" => :objects, "scenarios" => :scenarios,
          "formatCells" => :format_cells, "formatColumns" => :format_columns,
          "formatRows" => :format_rows, "insertColumns" => :insert_columns,
          "insertRows" => :insert_rows, "insertHyperlinks" => :insert_hyperlinks,
          "deleteColumns" => :delete_columns, "deleteRows" => :delete_rows,
          "selectLockedCells" => :select_locked_cells, "sort" => :sort,
          "autoFilter" => :auto_filter, "pivotTables" => :pivot_tables,
          "selectUnlockedCells" => :select_unlocked_cells,
          "password" => :password, "algorithmName" => :algorithm_name,
          "hashValue" => :hash_value, "saltValue" => :salt_value, "spinCount" => :spin_count
        }.freeze

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          return unless name == "sheetProtection"

          prot = {}
          ATTR_MAP.each do |xml_attr, sym|
            val = attributes[xml_attr]
            next if val.nil?

            prot[sym] = if sym == :spin_count
                          val.to_i
                        elsif %i[password algorithm_name hash_value salt_value].include?(sym)
                          val
                        else
                          %w[1 true].include?(val)
                        end
          end
          @protection = prot unless prot.empty?
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

      # SAX2 listener for parsing <protectedRanges> element.
      class ProtectedRangesListener
        include REXML::SAX2Listener

        attr_reader :ranges

        def initialize
          @ranges = []
          @current = nil
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          case name
          when "protectedRange"
            pr = {}
            pr[:sqref] = attributes["sqref"] if attributes["sqref"]
            pr[:name] = attributes["name"] if attributes["name"]
            pr[:algorithm_name] = attributes["algorithmName"] if attributes["algorithmName"]
            pr[:hash_value] = attributes["hashValue"] if attributes["hashValue"]
            pr[:salt_value] = attributes["saltValue"] if attributes["saltValue"]
            pr[:spin_count] = attributes["spinCount"].to_i if attributes["spinCount"]
            @current = pr
          when "securityDescriptor"
            @in_sd = true
            @sd_text = +""
          end
        end

        def characters(text)
          @sd_text << text if @in_sd
        end

        def end_element(_uri, local_name, qname)
          name = element_name(local_name, qname)
          case name
          when "securityDescriptor"
            (@current[:security_descriptors] ||= []) << @sd_text if @current
            @in_sd = false
          when "protectedRange"
            @ranges << @current if @current
            @current = nil
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

      # SAX2 listener for parsing <cellWatches> element.
      class CellWatchesListener
        include REXML::SAX2Listener

        attr_reader :watches

        def initialize
          @watches = []
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          @watches << attributes["r"] if name == "cellWatch" && attributes["r"]
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

      # SAX2 listener for parsing <ignoredErrors> element.
      class IgnoredErrorsListener
        include REXML::SAX2Listener

        attr_reader :errors

        def initialize
          @errors = []
        end

        IGNORED_ERROR_BOOL_ATTRS = {
          "evalError" => :eval_error, "twoDigitTextYear" => :two_digit_text_year,
          "numberStoredAsText" => :number_stored_as_text, "formula" => :formula,
          "formulaRange" => :formula_range, "unlockedFormula" => :unlocked_formula,
          "emptyCellReference" => :empty_cell_reference, "listDataValidation" => :list_data_validation,
          "calculatedColumn" => :calculated_column
        }.freeze

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          return unless name == "ignoredError" && attributes["sqref"]

          ie = { sqref: attributes["sqref"] }
          IGNORED_ERROR_BOOL_ATTRS.each do |xml_attr, sym|
            ie[sym] = true if attributes[xml_attr] == "1"
          end
          @errors << ie
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

      # SAX2 listener for parsing <dataConsolidate> element.
      class DataConsolidateListener
        include REXML::SAX2Listener

        attr_reader :result

        def initialize
          @result = nil
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          case name
          when "dataConsolidate"
            @result = {}
            @result[:function] = attributes["function"] if attributes["function"]
            @result[:start_labels] = true if %w[1 true].include?(attributes["startLabels"])
            @result[:left_labels] = true if %w[1 true].include?(attributes["leftLabels"])
            @result[:top_labels] = true if %w[1 true].include?(attributes["topLabels"])
            @result[:link] = true if %w[1 true].include?(attributes["link"])
            @result[:data_refs] = []
          when "dataRef"
            if @result
              ref = {}
              ref[:ref] = attributes["ref"] if attributes["ref"]
              ref[:name] = attributes["name"] if attributes["name"]
              ref[:sheet] = attributes["sheet"] if attributes["sheet"]
              @result[:data_refs] << ref
            end
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

      # SAX2 listener for parsing <scenarios> element.
      class ScenariosListener
        include REXML::SAX2Listener

        attr_reader :result

        def initialize
          @result = nil
          @current_scenario = nil
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          case name
          when "scenarios"
            @result = { scenarios: [] }
            @result[:current] = attributes["current"].to_i if attributes["current"]
            @result[:show] = attributes["show"].to_i if attributes["show"]
            @result[:sqref] = attributes["sqref"] if attributes["sqref"]
          when "scenario"
            if @result
              sc = { name: attributes["name"], input_cells: [] }
              sc[:locked] = true if %w[1 true].include?(attributes["locked"])
              sc[:hidden] = true if %w[1 true].include?(attributes["hidden"])
              sc[:user] = attributes["user"] if attributes["user"]
              sc[:comment] = attributes["comment"] if attributes["comment"]
              @current_scenario = sc
            end
          when "inputCells"
            if @current_scenario
              ic = { r: attributes["r"], val: attributes["val"] }
              ic[:deleted] = true if %w[1 true].include?(attributes["deleted"])
              ic[:undone] = true if %w[1 true].include?(attributes["undone"])
              ic[:num_fmt_id] = attributes["numFmtId"].to_i if attributes["numFmtId"]
              @current_scenario[:input_cells] << ic
            end
          end
        end

        def end_element(_uri, local_name, qname)
          name = element_name(local_name, qname)
          return unless name == "scenario" && @current_scenario && @result

          @result[:scenarios] << @current_scenario
          @current_scenario = nil
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

      # SAX2 listener for parsing <sheetFormatPr> element.
      class SheetFormatListener
        include REXML::SAX2Listener

        attr_reader :properties

        def initialize
          @properties = {}
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          return unless name == "sheetFormatPr"

          drh = attributes["defaultRowHeight"]
          @properties[:default_row_height] = drh.to_f if drh
          dcw = attributes["defaultColWidth"]
          @properties[:default_col_width] = dcw.to_f if dcw
          bcw = attributes["baseColWidth"]
          @properties[:base_col_width] = bcw.to_i if bcw
          olr = attributes["outlineLevelRow"]
          @properties[:outline_level_row] = olr.to_i if olr
          olc = attributes["outlineLevelCol"]
          @properties[:outline_level_col] = olc.to_i if olc
          ch = attributes["customHeight"]
          @properties[:custom_height] = true if %w[1 true].include?(ch)
          zh = attributes["zeroHeight"]
          @properties[:zero_height] = true if %w[1 true].include?(zh)
          tt = attributes["thickTop"]
          @properties[:thick_top] = true if %w[1 true].include?(tt)
          tb = attributes["thickBottom"]
          @properties[:thick_bottom] = true if %w[1 true].include?(tb)
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

      # SAX2 listener for parsing <sheetViews><sheetView>, <pane>, and <selection>.
      class SheetViewListener
        include REXML::SAX2Listener

        attr_reader :view, :pane, :selection

        def initialize
          @view = {}
          @pane = nil
          @selection = nil
          @inside_sheet_views = false
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          case name
          when "sheetViews"
            @inside_sheet_views = true
          when "sheetView"
            return unless @inside_sheet_views

            wp = attributes["windowProtection"]
            @view[:window_protection] = %w[1 true].include?(wp) unless wp.nil?
            sf = attributes["showFormulas"]
            @view[:show_formulas] = %w[1 true].include?(sf) unless sf.nil?
            sgl = attributes["showGridLines"]
            @view[:show_grid_lines] = %w[1 true].include?(sgl) unless sgl.nil?
            srch = attributes["showRowColHeaders"]
            @view[:show_row_col_headers] = %w[1 true].include?(srch) unless srch.nil?
            szv = attributes["showZeros"]
            @view[:show_zeros] = %w[1 true].include?(szv) unless szv.nil?
            rtl = attributes["rightToLeft"]
            @view[:right_to_left] = %w[1 true].include?(rtl) unless rtl.nil?
            ts = attributes["tabSelected"]
            @view[:tab_selected] = true if ts == "1"
            srr = attributes["showRuler"]
            @view[:show_ruler] = %w[1 true].include?(srr) unless srr.nil?
            soss = attributes["showOutlineSymbols"]
            @view[:show_outline_symbols] = %w[1 true].include?(soss) unless soss.nil?
            dgc = attributes["defaultGridColor"]
            @view[:default_grid_color] = %w[1 true].include?(dgc) unless dgc.nil?
            sws = attributes["showWhiteSpace"]
            @view[:show_white_space] = %w[1 true].include?(sws) unless sws.nil?
            vm = attributes["view"]
            @view[:view] = vm if vm
            tlc = attributes["topLeftCell"]
            @view[:top_left_cell] = tlc if tlc
            cid = attributes["colorId"]
            @view[:color_id] = cid.to_i if cid
            zs = attributes["zoomScale"]
            @view[:zoom_scale] = zs.to_i if zs
            zsn = attributes["zoomScaleNormal"]
            @view[:zoom_scale_normal] = zsn.to_i if zsn
            zssl = attributes["zoomScaleSheetLayoutView"]
            @view[:zoom_scale_sheet_layout_view] = zssl.to_i if zssl
            zspl = attributes["zoomScalePageLayoutView"]
            @view[:zoom_scale_page_layout_view] = zspl.to_i if zspl
          when "pane"
            return unless @inside_sheet_views

            ys = attributes["ySplit"]
            xs = attributes["xSplit"]
            frozen = attributes["state"] == "frozen"
            tlc = attributes["topLeftCell"]
            ap = attributes["activePane"]
            p = if frozen
                  {
                    row: ys ? ys.to_i : 0,
                    col: xs ? xs.to_i : 0,
                    state: :frozen
                  }
                else
                  {
                    row: ys ? ys.to_i : 0,
                    col: xs ? xs.to_i : 0,
                    x_split: xs ? xs.to_i : 0,
                    y_split: ys ? ys.to_i : 0,
                    top_left_cell: tlc,
                    state: :split
                  }
                end
            p[:active_pane] = ap if ap
            @pane = p
          when "selection"
            return unless @inside_sheet_views

            ac = attributes["activeCell"]
            sq = attributes["sqref"]
            sel = { active_cell: ac, sqref: sq }
            pn = attributes["pane"]
            sel[:pane] = pn if pn
            acid = attributes["activeCellId"]
            sel[:active_cell_id] = acid.to_i if acid
            @selection = sel if ac || sq || pn || acid
          end
        end

        def end_element(_uri, local_name, qname)
          name = element_name(local_name, qname)
          @inside_sheet_views = false if name == "sheetViews"
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

      # SAX2 listener for parsing <dataValidations> elements.
      class DataValidationsListener
        include REXML::SAX2Listener

        attr_reader :validations, :container_options

        def initialize
          @validations = []
          @container_options = {}
          @current_dv = nil
          @inside_formula1 = false
          @inside_formula2 = false
          @text_buffer = +""
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          case name
          when "dataValidations"
            dp = attributes["disablePrompts"]
            @container_options[:disable_prompts] = true if %w[1 true].include?(dp)
            xw = attributes["xWindow"]
            @container_options[:x_window] = xw.to_i if xw
            yw = attributes["yWindow"]
            @container_options[:y_window] = yw.to_i if yw
          when "dataValidation"
            @current_dv = { sqref: attributes["sqref"] }
            @current_dv[:type] = attributes["type"] if attributes["type"]
            @current_dv[:operator] = attributes["operator"] if attributes["operator"]
            @current_dv[:error_style] = attributes["errorStyle"] if attributes["errorStyle"]
            @current_dv[:allow_blank] = true if attributes["allowBlank"] == "1"
            @current_dv[:show_input_message] = true if attributes["showInputMessage"] == "1"
            @current_dv[:show_error_message] = true if attributes["showErrorMessage"] == "1"
            @current_dv[:error_title] = xml_unescape(attributes["errorTitle"]) if attributes["errorTitle"]
            @current_dv[:error] = xml_unescape(attributes["error"]) if attributes["error"]
            @current_dv[:prompt_title] = xml_unescape(attributes["promptTitle"]) if attributes["promptTitle"]
            @current_dv[:prompt] = xml_unescape(attributes["prompt"]) if attributes["prompt"]
            @current_dv[:show_drop_down] = true if attributes["showDropDown"] == "1"
            @current_dv[:ime_mode] = attributes["imeMode"] if attributes["imeMode"]
          when "formula1"
            @inside_formula1 = true
            @text_buffer = +""
          when "formula2"
            @inside_formula2 = true
            @text_buffer = +""
          end
        end

        def characters(text)
          @text_buffer << text if @inside_formula1 || @inside_formula2
        end

        def end_element(_uri, local_name, qname)
          name = element_name(local_name, qname)
          case name
          when "formula1"
            @current_dv[:formula1] = @text_buffer.dup if @current_dv
            @inside_formula1 = false
          when "formula2"
            @current_dv[:formula2] = @text_buffer.dup if @current_dv
            @inside_formula2 = false
          when "dataValidation"
            @validations << @current_dv if @current_dv
            @current_dv = nil
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

        def xml_unescape(str)
          str.gsub("&amp;", "&").gsub("&lt;", "<").gsub("&gt;", ">").gsub("&quot;", '"').gsub("&apos;", "'")
        end
      end

      # SAX2 listener for parsing conditionalFormatting elements.
      class ConditionalFormattingListener
        include REXML::SAX2Listener

        attr_reader :rules

        def initialize
          @rules = []
          @current_sqref = nil
          @current_pivot = false
          @current_rule = nil
          @inside_formula = false
          @text_buffer = +""
          @cfvo_target = nil
          @color_target = nil
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          case name
          when "conditionalFormatting"
            @current_sqref = attributes["sqref"]
            @current_pivot = attributes["pivot"] == "1"
          when "cfRule"
            @current_rule = { sqref: @current_sqref, type: attributes["type"] }
            @current_rule[:pivot] = true if @current_pivot
            @current_rule[:priority] = attributes["priority"].to_i if attributes["priority"]
            @current_rule[:operator] = attributes["operator"] if attributes["operator"]
            @current_rule[:format_id] = attributes["dxfId"].to_i if attributes["dxfId"]
            @current_rule[:stop_if_true] = true if attributes["stopIfTrue"] == "1"
            @current_rule[:above_average] = false if attributes["aboveAverage"] == "0"
            @current_rule[:equal_average] = true if attributes["equalAverage"] == "1"
            @current_rule[:rank] = attributes["rank"].to_i if attributes["rank"]
            @current_rule[:percent] = true if attributes["percent"] == "1"
            @current_rule[:bottom] = true if attributes["bottom"] == "1"
            @current_rule[:text] = attributes["text"] if attributes["text"]
            @current_rule[:time_period] = attributes["timePeriod"] if attributes["timePeriod"]
            sd = attributes["stdDev"]
            @current_rule[:std_dev] = sd.to_i if sd
          when "formula"
            @inside_formula = true
            @text_buffer = +""
          when "colorScale"
            @current_rule[:color_scale] = { cfvo: [], colors: [] } if @current_rule
            @cfvo_target = :color_scale
            @color_target = :color_scale
          when "dataBar"
            if @current_rule
              db = { cfvo: [] }
              db[:min_length] = attributes["minLength"].to_i if attributes["minLength"]
              db[:max_length] = attributes["maxLength"].to_i if attributes["maxLength"]
              sv = attributes["showValue"]
              db[:show_value] = %w[1 true].include?(sv) unless sv.nil?
              @current_rule[:data_bar] = db
            end
            @cfvo_target = :data_bar
            @color_target = :data_bar
          when "iconSet"
            if @current_rule
              is = { cfvo: [] }
              is[:icon_set] = attributes["iconSet"] if attributes["iconSet"]
              rv = attributes["reverse"]
              is[:reverse] = %w[1 true].include?(rv) unless rv.nil?
              pct = attributes["percent"]
              is[:percent] = %w[1 true].include?(pct) unless pct.nil?
              sv = attributes["showValue"]
              is[:show_value] = %w[1 true].include?(sv) unless sv.nil?
              @current_rule[:icon_set] = is
            end
            @cfvo_target = :icon_set
          when "cfvo"
            cfvo = { type: attributes["type"] }
            cfvo[:val] = attributes["val"] if attributes["val"]
            gte = attributes["gte"]
            cfvo[:gte] = %w[1 true].include?(gte) unless gte.nil?
            append_cfvo(cfvo)
          when "color"
            if attributes["rgb"]
              append_cf_color({ rgb: attributes["rgb"] })
            elsif attributes["theme"]
              c = { theme: attributes["theme"].to_i }
              c[:tint] = attributes["tint"].to_f if attributes["tint"]
              append_cf_color(c)
            elsif attributes["indexed"]
              append_cf_color({ indexed: attributes["indexed"].to_i })
            end
          end
        end

        def characters(text)
          @text_buffer << text if @inside_formula
        end

        def end_element(_uri, local_name, qname)
          name = element_name(local_name, qname)
          case name
          when "formula"
            (@current_rule[:formulas] ||= []) << @text_buffer.dup if @current_rule
            @inside_formula = false
          when "cfRule"
            @rules << @current_rule if @current_rule
            @current_rule = nil
          when "conditionalFormatting"
            @current_sqref = nil
          when "colorScale", "dataBar", "iconSet"
            @cfvo_target = nil
            @color_target = nil
          end
        end

        private

        def append_cfvo(cfvo)
          return unless @current_rule && @cfvo_target

          container = @current_rule[@cfvo_target]
          container[:cfvo] << cfvo if container
        end

        def append_cf_color(color)
          return unless @current_rule && @color_target

          container = @current_rule[@color_target]
          if container.is_a?(Hash) && container.key?(:colors)
            container[:colors] << color
          elsif container.is_a?(Hash)
            container[:color] = color
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

      # SAX2 listener for parsing print/page elements from worksheet XML.
      class PrintPageListener
        include REXML::SAX2Listener

        attr_reader :print_options, :page_margins, :page_setup, :header_footer, :row_breaks, :col_breaks

        def initialize
          @print_options = {}
          @page_margins = nil
          @page_setup = {}
          @header_footer = {}
          @row_breaks = []
          @col_breaks = []
          @inside_header_footer = false
          @inside_row_breaks = false
          @inside_col_breaks = false
          @current_hf_field = nil
          @text_buffer = +""
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          case name
          when "printOptions"
            @print_options[:grid_lines] = true if attributes["gridLines"] == "1"
            @print_options[:headings] = true if attributes["headings"] == "1"
            @print_options[:horizontal_centered] = true if attributes["horizontalCentered"] == "1"
            @print_options[:vertical_centered] = true if attributes["verticalCentered"] == "1"
            gls = attributes["gridLinesSet"]
            @print_options[:grid_lines_set] = gls != "0" if gls
          when "pageMargins"
            m = {}
            %w[left right top bottom header footer].each do |k|
              v = attributes[k]
              m[k.to_sym] = v.to_f if v
            end
            @page_margins = m unless m.empty?
          when "pageSetup"
            o = attributes["orientation"]
            @page_setup[:orientation] = o if o
            ps = attributes["paperSize"]
            @page_setup[:paper_size] = ps.to_i if ps
            sc = attributes["scale"]
            @page_setup[:scale] = sc.to_i if sc
            ftw = attributes["fitToWidth"]
            @page_setup[:fit_to_width] = ftw.to_i if ftw
            fth = attributes["fitToHeight"]
            @page_setup[:fit_to_height] = fth.to_i if fth
            po = attributes["pageOrder"]
            @page_setup[:page_order] = po if po
            baw = attributes["blackAndWhite"]
            @page_setup[:black_and_white] = true if %w[1 true].include?(baw)
            dr = attributes["draft"]
            @page_setup[:draft] = true if %w[1 true].include?(dr)
            cc = attributes["cellComments"]
            @page_setup[:cell_comments] = cc if cc
            fpn = attributes["firstPageNumber"]
            @page_setup[:first_page_number] = fpn.to_i if fpn
            ufpn = attributes["useFirstPageNumber"]
            @page_setup[:use_first_page_number] = true if %w[1 true].include?(ufpn)
            hdpi = attributes["horizontalDpi"]
            @page_setup[:horizontal_dpi] = hdpi.to_i if hdpi
            vdpi = attributes["verticalDpi"]
            @page_setup[:vertical_dpi] = vdpi.to_i if vdpi
            cp = attributes["copies"]
            @page_setup[:copies] = cp.to_i if cp
            ph = attributes["paperHeight"]
            @page_setup[:paper_height] = ph if ph
            pw = attributes["paperWidth"]
            @page_setup[:paper_width] = pw if pw
            err = attributes["errors"]
            @page_setup[:errors] = err if err
            upd = attributes["usePrinterDefaults"]
            @page_setup[:use_printer_defaults] = %w[1 true].include?(upd) unless upd.nil?
          when "headerFooter"
            @inside_header_footer = true
            df = attributes["differentFirst"]
            @header_footer[:different_first] = true if %w[1 true].include?(df)
            doe = attributes["differentOddEven"]
            @header_footer[:different_odd_even] = true if %w[1 true].include?(doe)
            swd = attributes["scaleWithDoc"]
            @header_footer[:scale_with_doc] = swd != "0" if swd
            awm = attributes["alignWithMargins"]
            @header_footer[:align_with_margins] = awm != "0" if awm
          when "oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter"
            if @inside_header_footer
              @current_hf_field = name
              @text_buffer = +""
            end
          when "rowBreaks"
            @inside_row_breaks = true
          when "colBreaks"
            @inside_col_breaks = true
          when "brk"
            id = attributes["id"]&.to_i
            if id
              brk = { id: id }
              mn = attributes["min"]
              brk[:min] = mn.to_i if mn
              mx = attributes["max"]
              brk[:max] = mx.to_i if mx
              brk[:man] = true if %w[1 true].include?(attributes["man"])
              brk[:pt] = true if %w[1 true].include?(attributes["pt"])
              @row_breaks << brk if @inside_row_breaks
              @col_breaks << brk if @inside_col_breaks
            end
          end
        end

        def characters(text)
          @text_buffer << text if @current_hf_field
        end

        def end_element(_uri, local_name, qname)
          name = element_name(local_name, qname)
          case name
          when "headerFooter"
            @inside_header_footer = false
          when "oddHeader"
            @header_footer[:odd_header] = @text_buffer.dup if @current_hf_field == "oddHeader"
            @current_hf_field = nil
          when "oddFooter"
            @header_footer[:odd_footer] = @text_buffer.dup if @current_hf_field == "oddFooter"
            @current_hf_field = nil
          when "evenHeader"
            @header_footer[:even_header] = @text_buffer.dup if @current_hf_field == "evenHeader"
            @current_hf_field = nil
          when "evenFooter"
            @header_footer[:even_footer] = @text_buffer.dup if @current_hf_field == "evenFooter"
            @current_hf_field = nil
          when "firstHeader"
            @header_footer[:first_header] = @text_buffer.dup if @current_hf_field == "firstHeader"
            @current_hf_field = nil
          when "firstFooter"
            @header_footer[:first_footer] = @text_buffer.dup if @current_hf_field == "firstFooter"
            @current_hf_field = nil
          when "rowBreaks"
            @inside_row_breaks = false
          when "colBreaks"
            @inside_col_breaks = false
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

      # SAX2 listener for parsing table XML.
      class TableListener
        include REXML::SAX2Listener

        attr_reader :table

        def initialize
          @table = nil
          @columns = []
          @current_column = nil
          @inside_calc_formula = false
          @inside_totals_formula = false
          @text_buffer = +""
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          case name
          when "table"
            @table = {
              id: attributes["id"]&.to_i,
              name: attributes["name"],
              display_name: attributes["displayName"],
              ref: attributes["ref"]
            }
            trc = attributes["totalsRowCount"]
            @table[:totals_row_count] = trc.to_i if trc
            hrc = attributes["headerRowCount"]
            @table[:header_row_count] = hrc.to_i if hrc
            @table[:published] = true if attributes["published"] == "1"
            @table[:comment] = attributes["comment"] if attributes["comment"]
            @table[:insert_row] = true if attributes["insertRow"] == "1"
            @table[:insert_row_shift] = true if attributes["insertRowShift"] == "1"
            hrd = attributes["headerRowDxfId"]
            @table[:header_row_dxf_id] = hrd.to_i if hrd
            dd = attributes["dataDxfId"]
            @table[:data_dxf_id] = dd.to_i if dd
            trd = attributes["totalsRowDxfId"]
            @table[:totals_row_dxf_id] = trd.to_i if trd
            hrbd = attributes["headerRowBorderDxfId"]
            @table[:header_row_border_dxf_id] = hrbd.to_i if hrbd
            tbd = attributes["tableBorderDxfId"]
            @table[:table_border_dxf_id] = tbd.to_i if tbd
            trbd = attributes["totalsRowBorderDxfId"]
            @table[:totals_row_border_dxf_id] = trbd.to_i if trbd
            @table[:header_row_cell_style] = attributes["headerRowCellStyle"] if attributes["headerRowCellStyle"]
            @table[:totals_row_cell_style] = attributes["totalsRowCellStyle"] if attributes["totalsRowCellStyle"]
            cid = attributes["connectionId"]
            @table[:connection_id] = cid.to_i if cid
            @table[:table_type] = attributes["tableType"] if attributes["tableType"]
          when "tableColumn"
            col = { name: attributes["name"] }
            trf = attributes["totalsRowFunction"]
            col[:totals_row_function] = trf if trf
            trl = attributes["totalsRowLabel"]
            col[:totals_row_label] = trl if trl
            dd = attributes["dataDxfId"]
            col[:data_dxf_id] = dd.to_i if dd
            td = attributes["totalsRowDxfId"]
            col[:totals_row_dxf_id] = td.to_i if td
            hd = attributes["headerRowDxfId"]
            col[:header_row_dxf_id] = hd.to_i if hd
            dcs = attributes["dataCellStyle"]
            col[:data_cell_style] = dcs if dcs
            @current_column = col
          when "calculatedColumnFormula"
            @inside_calc_formula = true
            @text_buffer = +""
          when "totalsRowFormula"
            @inside_totals_formula = true
            @text_buffer = +""
          when "tableStyleInfo"
            if @table
              si = {}
              si[:name] = attributes["name"] if attributes["name"]
              sfc = attributes["showFirstColumn"]
              si[:show_first_column] = sfc == "1" unless sfc.nil?
              slc = attributes["showLastColumn"]
              si[:show_last_column] = slc == "1" unless slc.nil?
              srs = attributes["showRowStripes"]
              si[:show_row_stripes] = srs == "1" unless srs.nil?
              scs = attributes["showColumnStripes"]
              si[:show_column_stripes] = scs == "1" unless scs.nil?
              @table[:style] = si
            end
          end
        end

        def characters(text)
          @text_buffer << text if @inside_calc_formula || @inside_totals_formula
        end

        def end_element(_uri, local_name, qname)
          name = element_name(local_name, qname)
          case name
          when "calculatedColumnFormula"
            @current_column[:calculated_column_formula] = @text_buffer.dup if @current_column
            @inside_calc_formula = false
          when "totalsRowFormula"
            @current_column[:totals_row_formula] = @text_buffer.dup if @current_column
            @inside_totals_formula = false
          when "tableColumn"
            @columns << @current_column if @current_column
            @current_column = nil
          when "table"
            @table[:columns] = @columns if @table
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

      # SAX2 listener for parsing calcChain.xml.
      class CalcChainListener
        include REXML::SAX2Listener

        attr_reader :entries

        def initialize
          @entries = []
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          return unless name == "c"

          entry = {}
          entry[:ref] = attributes["r"] if attributes["r"]
          i = attributes["i"]
          entry[:sheet_id] = i.to_i if i
          @entries << entry
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

      # SAX2 listener for parsing pivotTable XML.
      class PivotTableListener
        include REXML::SAX2Listener

        attr_reader :pivot_table

        def initialize
          @pivot_table = nil
          @fields = []
          @row_fields = []
          @col_fields = []
          @data_fields = []
          @inside_row_fields = false
          @inside_col_fields = false
          @inside_data_fields = false
          @inside_pivot_field = false
          @inside_items = false
          @current_items = []
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          case name
          when "pivotTableDefinition"
            @pivot_table = {
              name: attributes["name"],
              cache_id: attributes["cacheId"]&.to_i
            }
            @pivot_table[:data_caption] = attributes["dataCaption"] if attributes["dataCaption"]
            @pivot_table[:data_on_rows] = attributes["dataOnRows"] == "1" if attributes["dataOnRows"]
            @pivot_table[:row_grand_totals] = attributes["rowGrandTotals"] != "0" if attributes["rowGrandTotals"]
            @pivot_table[:col_grand_totals] = attributes["colGrandTotals"] != "0" if attributes["colGrandTotals"]
            @pivot_table[:compact] = attributes["compact"] != "0" if attributes["compact"]
            @pivot_table[:outline] = attributes["outline"] != "0" if attributes["outline"]
            @pivot_table[:outline_data] = attributes["outlineData"] == "1" if attributes["outlineData"]
            @pivot_table[:compact_data] = attributes["compactData"] != "0" if attributes["compactData"]
            @pivot_table[:show_headers] = attributes["showHeaders"] != "0" if attributes["showHeaders"]
            @pivot_table[:show_multiple_label] = attributes["showMultipleLabel"] != "0" if attributes["showMultipleLabel"]
            @pivot_table[:show_data_drop_down] = attributes["showDataDropDown"] != "0" if attributes["showDataDropDown"]
            @pivot_table[:grand_total_caption] = attributes["grandTotalCaption"] if attributes["grandTotalCaption"]
            @pivot_table[:error_caption] = attributes["errorCaption"] if attributes["errorCaption"]
            @pivot_table[:show_error] = attributes["showError"] == "1" if attributes["showError"]
            @pivot_table[:missing_caption] = attributes["missingCaption"] if attributes["missingCaption"]
            @pivot_table[:show_missing] = attributes["showMissing"] != "0" if attributes["showMissing"]
            @pivot_table[:tag] = attributes["tag"] if attributes["tag"]
            @pivot_table[:indent] = attributes["indent"]&.to_i if attributes["indent"]
            @pivot_table[:published] = attributes["published"] == "1" if attributes["published"]
            @pivot_table[:edit_data] = attributes["editData"] == "1" if attributes["editData"]
            @pivot_table[:disable_field_list] = attributes["disableFieldList"] == "1" if attributes["disableFieldList"]
            @pivot_table[:visual_totals] = attributes["visualTotals"] != "0" if attributes["visualTotals"]
            @pivot_table[:print_drill] = attributes["printDrill"] == "1" if attributes["printDrill"]
            @pivot_table[:created_version] = attributes["createdVersion"]&.to_i if attributes["createdVersion"]
            @pivot_table[:updated_version] = attributes["updatedVersion"]&.to_i if attributes["updatedVersion"]
            @pivot_table[:min_refreshable_version] = attributes["minRefreshableVersion"]&.to_i if attributes["minRefreshableVersion"]
            %w[applyNumberFormats applyBorderFormats applyFontFormats
               applyPatternFormats applyAlignmentFormats applyWidthHeightFormats].each do |attr|
              next if attributes[attr].nil?

              key = attr.gsub(/[A-Z]/) { |m| "_#{m.downcase}" }.to_sym
              @pivot_table[key] = %w[1 true].include?(attributes[attr])
            end
            mff = attributes["multipleFieldFilters"]
            @pivot_table[:multiple_field_filters] = mff != "0" unless mff.nil?
            sdr = attributes["showDrill"]
            @pivot_table[:show_drill] = sdr != "0" unless sdr.nil?
            sdt = attributes["showDataTips"]
            @pivot_table[:show_data_tips] = sdt != "0" unless sdt.nil?
            edr = attributes["enableDrill"]
            @pivot_table[:enable_drill] = edr != "0" unless edr.nil?
            smpt = attributes["showMemberPropertyTips"]
            @pivot_table[:show_member_property_tips] = smpt != "0" unless smpt.nil?
            ipt = attributes["itemPrintTitles"]
            @pivot_table[:item_print_titles] = ipt == "1" unless ipt.nil?
            fpt = attributes["fieldPrintTitles"]
            @pivot_table[:field_print_titles] = fpt == "1" unless fpt.nil?
            pf = attributes["preserveFormatting"]
            @pivot_table[:preserve_formatting] = pf != "0" unless pf.nil?
            potd = attributes["pageOverThenDown"]
            @pivot_table[:page_over_then_down] = potd == "1" unless potd.nil?
            pw = attributes["pageWrap"]
            @pivot_table[:page_wrap] = pw.to_i if pw
          when "location"
            @pivot_table[:ref] = attributes["ref"] if @pivot_table
            @pivot_table[:row_page_count] = attributes["rowPageCount"]&.to_i if attributes["rowPageCount"]
            @pivot_table[:col_page_count] = attributes["colPageCount"]&.to_i if attributes["colPageCount"]
          when "pivotField"
            @inside_pivot_field = true
            @current_field = {}
            @current_field[:axis] = attributes["axis"] if attributes["axis"]
            @current_field[:data_field] = true if attributes["dataField"] == "1"
            @current_field[:name] = attributes["name"] if attributes["name"]
            @current_field[:show_all] = attributes["showAll"] != "0" if attributes["showAll"]
            @current_field[:compact] = attributes["compact"] != "0" if attributes["compact"]
            @current_field[:outline] = attributes["outline"] != "0" if attributes["outline"]
            @current_field[:subtotal_top] = attributes["subtotalTop"] != "0" if attributes["subtotalTop"]
            @current_field[:num_fmt_id] = attributes["numFmtId"]&.to_i if attributes["numFmtId"]
            @current_field[:sort_type] = attributes["sortType"] if attributes["sortType"]
            ds = attributes["defaultSubtotal"]
            @current_field[:default_subtotal] = ds != "0" unless ds.nil?
            @current_field[:insert_blank_row] = true if attributes["insertBlankRow"] == "1"
            @current_field[:insert_page_break] = true if attributes["insertPageBreak"] == "1"
            @current_field[:include_new_items_in_filter] = true if attributes["includeNewItemsInFilter"] == "1"
            @current_items = []
          when "items"
            @inside_items = true if @inside_pivot_field
          when "item"
            if @inside_items
              item_type = attributes["t"]
              item_x = attributes["x"]&.to_i
              @current_items << { x: item_x, t: item_type } if item_type || item_x
            end
          when "rowFields"
            @inside_row_fields = true
          when "colFields"
            @inside_col_fields = true
          when "dataFields"
            @inside_data_fields = true
          when "field"
            idx = attributes["x"]&.to_i
            @row_fields << idx if @inside_row_fields && idx
            @col_fields << idx if @inside_col_fields && idx
          when "dataField"
            if @inside_data_fields
              df = {
                name: attributes["name"],
                fld: attributes["fld"]&.to_i,
                subtotal: attributes["subtotal"] || "sum"
              }
              df[:show_data_as] = attributes["showDataAs"] if attributes["showDataAs"]
              df[:base_field] = attributes["baseField"]&.to_i if attributes["baseField"]
              df[:base_item] = attributes["baseItem"]&.to_i if attributes["baseItem"]
              df[:num_fmt_id] = attributes["numFmtId"]&.to_i if attributes["numFmtId"]
              @data_fields << df
            end
          when "pivotTableStyleInfo"
            if @pivot_table
              psi = {}
              psi[:name] = attributes["name"] if attributes["name"]
              srh = attributes["showRowHeaders"]
              psi[:show_row_headers] = srh == "1" unless srh.nil?
              sch = attributes["showColHeaders"]
              psi[:show_col_headers] = sch == "1" unless sch.nil?
              srs = attributes["showRowStripes"]
              psi[:show_row_stripes] = srs == "1" unless srs.nil?
              scs = attributes["showColStripes"]
              psi[:show_col_stripes] = scs == "1" unless scs.nil?
              slc = attributes["showLastColumn"]
              psi[:show_last_column] = slc == "1" unless slc.nil?
              @pivot_table[:pivot_table_style] = psi
            end
          end
        end

        def end_element(_uri, local_name, qname)
          name = element_name(local_name, qname)
          case name
          when "pivotField"
            @current_field[:items] = @current_items unless @current_items.empty?
            @fields << @current_field
            @inside_pivot_field = false
          when "items"
            @inside_items = false
          when "pivotTableDefinition"
            if @pivot_table
              @pivot_table[:fields] = @fields
              @pivot_table[:row_fields] = @row_fields
              @pivot_table[:col_fields] = @col_fields
              @pivot_table[:data_fields] = @data_fields
            end
          when "rowFields"
            @inside_row_fields = false
          when "colFields"
            @inside_col_fields = false
          when "dataFields"
            @inside_data_fields = false
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

      # SAX2 listener for parsing pivotCacheDefinition XML.
      class PivotCacheDefinitionListener
        include REXML::SAX2Listener

        attr_reader :cache_definition

        def initialize
          @cache_definition = {}
          @fields = []
          @current_field = nil
          @current_shared_items = nil
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          case name
          when "pivotCacheDefinition"
            sd = attributes["saveData"]
            @cache_definition[:save_data] = sd != "0" unless sd.nil?
            er = attributes["enableRefresh"]
            @cache_definition[:enable_refresh] = er != "0" unless er.nil?
            @cache_definition[:refreshed_by] = attributes["refreshedBy"] if attributes["refreshedBy"]
            @cache_definition[:refreshed_version] = attributes["refreshedVersion"]&.to_i if attributes["refreshedVersion"]
            @cache_definition[:created_version] = attributes["createdVersion"]&.to_i if attributes["createdVersion"]
            @cache_definition[:record_count] = attributes["recordCount"]&.to_i if attributes["recordCount"]
            om = attributes["optimizeMemory"]
            @cache_definition[:optimize_memory] = om == "1" unless om.nil?
          when "cacheSource"
            @cache_definition[:source_type] = attributes["type"] if attributes["type"]
          when "worksheetSource"
            @cache_definition[:source_ref] = attributes["ref"] if attributes["ref"]
            @cache_definition[:source_sheet] = attributes["sheet"] if attributes["sheet"]
            @cache_definition[:source_name] = attributes["name"] if attributes["name"]
          when "cacheField"
            @current_field = {}
            @current_field[:name] = attributes["name"] if attributes["name"]
            @current_field[:num_fmt_id] = attributes["numFmtId"]&.to_i if attributes["numFmtId"]
            @current_field[:caption] = attributes["caption"] if attributes["caption"]
            @current_field[:formula] = xml_unescape(attributes["formula"]) if attributes["formula"]
          when "sharedItems"
            @current_shared_items = [] if @current_field
          when "s", "d", "e"
            @current_shared_items << attributes["v"] if @current_shared_items && attributes["v"]
          when "n"
            @current_shared_items << attributes["v"]&.to_f if @current_shared_items && attributes["v"]
          when "b"
            @current_shared_items << (attributes["v"] == "1") if @current_shared_items
          when "m"
            @current_shared_items << nil if @current_shared_items
          end
        end

        def end_element(_uri, local_name, qname)
          name = element_name(local_name, qname)
          case name
          when "sharedItems"
            @current_field[:shared_items] = @current_shared_items if @current_field && @current_shared_items && !@current_shared_items.empty?
            @current_shared_items = nil
          when "cacheField"
            if @current_field
              @fields << @current_field
              @current_field = nil
            end
          end
        end

        def characters(_text); end

        def end_document
          @cache_definition[:fields] = @fields unless @fields.empty?
        end

        private

        def element_name(local_name, qname)
          if local_name.nil? || local_name.empty?
            qname.to_s.split(":").last
          else
            local_name
          end
        end

        def xml_unescape(str)
          str.gsub("&amp;", "&").gsub("&lt;", "<").gsub("&gt;", ">").gsub("&quot;", '"').gsub("&apos;", "'")
        end
      end

      # SAX2 listener for parsing pivotCacheRecords XML.
      class PivotCacheRecordsListener
        include REXML::SAX2Listener

        attr_reader :records

        def initialize
          @records = []
          @current_record = nil
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          case name
          when "r"
            @current_record = []
          when "x"
            @current_record << { x: attributes["v"]&.to_i } if @current_record
          when "s", "d", "e"
            @current_record << attributes["v"] if @current_record && attributes["v"]
          when "n"
            @current_record << attributes["v"]&.to_f if @current_record && attributes["v"]
          when "b"
            @current_record << (attributes["v"] == "1") if @current_record
          when "m"
            @current_record << nil if @current_record
          end
        end

        def end_element(_uri, local_name, qname)
          name = element_name(local_name, qname)
          return unless name == "r" && @current_record

          @records << @current_record
          @current_record = nil
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

      # SAX2 listener for parsing externalLink XML.
      class ExternalLinkListener
        include REXML::SAX2Listener

        attr_reader :sheet_names

        def initialize
          @sheet_names = []
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          return unless name == "sheetName"

          @sheet_names << attributes["val"] if attributes["val"]
        end

        def end_element(_uri, _local_name, _qname); end

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
end
