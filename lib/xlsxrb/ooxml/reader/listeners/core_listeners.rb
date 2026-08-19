# frozen_string_literal: true

# rbs_inline: enabled

require "rexml/parsers/sax2parser"
require "rexml/sax2listener"

module Xlsxrb
  module Ooxml
    class Reader
      # SAX2 listener for parsing shared string table (xl/sharedStrings.xml).
      class SharedStringsListener
        include REXML::SAX2Listener

        attr_reader :strings

        def initialize
          @strings = []
          @inside_si = false
          @inside_r = false
          @inside_rpr = false
          @inside_t = false
          @text_buffer = +""
          @runs = []
          @current_font = {}
          @has_runs = false
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)

          case name
          when "si"
            @inside_si = true
            @text_buffer = +""
            @runs = []
            @has_runs = false
          when "r"
            @inside_r = true
            @has_runs = true
            @current_font = {}
          when "rPr"
            @inside_rpr = true if @inside_r
          when "b"
            @current_font[:bold] = true if @inside_rpr
          when "i"
            @current_font[:italic] = true if @inside_rpr
          when "strike"
            @current_font[:strike] = true if @inside_rpr
          when "u"
            if @inside_rpr
              val = attributes["val"]
              @current_font[:underline] = val || true
            end
          when "vertAlign"
            @current_font[:vert_align] = attributes["val"] if @inside_rpr && attributes["val"]
          when "sz"
            @current_font[:sz] = attributes["val"]&.to_f if @inside_rpr
          when "color"
            if @inside_rpr
              if attributes["rgb"]
                @current_font[:color] = attributes["rgb"]
              elsif attributes["theme"]
                @current_font[:theme] = attributes["theme"].to_i
                @current_font[:tint] = attributes["tint"].to_f if attributes["tint"]
              elsif attributes["indexed"]
                @current_font[:indexed] = attributes["indexed"].to_i
              end
            end
          when "rFont"
            @current_font[:name] = attributes["val"] if @inside_rpr
          when "family"
            @current_font[:family] = attributes["val"]&.to_i if @inside_rpr
          when "scheme"
            @current_font[:scheme] = attributes["val"] if @inside_rpr
          when "t"
            @inside_t = @inside_si
            @text_buffer = +"" if @inside_r
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
          when "rPr"
            @inside_rpr = false
          when "r"
            run = { text: @text_buffer.dup }
            run[:font] = @current_font.dup unless @current_font.empty?
            @runs << run
            @inside_r = false
          when "si"
            @strings << if @has_runs
                          Xlsxrb::Elements::RichText.new(runs: @runs)
                        else
                          @text_buffer.dup
                        end
            @inside_si = false
            @text_buffer = +""
            @runs = []
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

        attr_reader :cells, :row_attributes, :cell_phonetic

        def initialize(shared_strings = [])
          @shared_strings = shared_strings
          @cells = {}
          @row_attributes = {}
          @cell_phonetic = {}
          @current_cell_ref = nil
          @current_cell_type = nil
          @inside_value = false
          @inside_inline_text = false
          @inside_formula = false
          @inside_is = false
          @inside_is_r = false
          @inside_is_rpr = false
          @value_buffer = +""
          @inline_text_buffer = +""
          @formula_buffer = +""
          @is_runs = []
          @is_has_runs = false
          @is_current_font = {}
          @is_run_text = +""
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)

          case name
          when "row"
            parse_row_attributes(attributes)
          when "c"
            @current_cell_ref = attributes["r"]
            @current_cell_type = attributes["t"]
            @current_cell_ph = %w[1 true].include?(attributes["ph"])
            @value_buffer = +""
            @inline_text_buffer = +""
            @formula_buffer = +""
            @formula_type = nil
            @formula_ref = nil
            @formula_si = nil
            @formula_ca = nil
            @formula_aca = nil
            @formula_bx = nil
            @formula_dt2d = nil
            @formula_dtr = nil
            @formula_r1 = nil
            @formula_r2 = nil
            @is_runs = []
            @is_has_runs = false
          when "v"
            @inside_value = true
          when "f"
            @inside_formula = true
            @formula_type = attributes["t"]
            @formula_ref = attributes["ref"]
            si = attributes["si"]
            @formula_si = si&.to_i
            @formula_ca = true if %w[1 true].include?(attributes["ca"])
            @formula_aca = true if %w[1 true].include?(attributes["aca"])
            @formula_bx = true if %w[1 true].include?(attributes["bx"])
            @formula_dt2d = true if %w[1 true].include?(attributes["dt2D"])
            @formula_dtr = true if %w[1 true].include?(attributes["dtr"])
            @formula_r1 = attributes["r1"] if attributes["r1"]
            @formula_r2 = attributes["r2"] if attributes["r2"]
          when "is"
            @inside_is = true if @current_cell_type == "inlineStr"
          when "r"
            if @inside_is
              @inside_is_r = true
              @is_has_runs = true
              @is_current_font = {}
              @is_run_text = +""
            end
          when "rPr"
            @inside_is_rpr = true if @inside_is_r
          when "b"
            @is_current_font[:bold] = true if @inside_is_rpr
          when "i"
            @is_current_font[:italic] = true if @inside_is_rpr
          when "strike"
            @is_current_font[:strike] = true if @inside_is_rpr
          when "u"
            if @inside_is_rpr
              val = attributes["val"]
              @is_current_font[:underline] = val || true
            end
          when "vertAlign"
            @is_current_font[:vert_align] = attributes["val"] if @inside_is_rpr && attributes["val"]
          when "sz"
            @is_current_font[:sz] = attributes["val"]&.to_f if @inside_is_rpr
          when "color"
            if @inside_is_rpr
              if attributes["rgb"]
                @is_current_font[:color] = attributes["rgb"]
              elsif attributes["theme"]
                @is_current_font[:theme] = attributes["theme"].to_i
                @is_current_font[:tint] = attributes["tint"].to_f if attributes["tint"]
              elsif attributes["indexed"]
                @is_current_font[:indexed] = attributes["indexed"].to_i
              end
            end
          when "rFont"
            @is_current_font[:name] = attributes["val"] if @inside_is_rpr
          when "family"
            @is_current_font[:family] = attributes["val"]&.to_i if @inside_is_rpr
          when "scheme"
            @is_current_font[:scheme] = attributes["val"] if @inside_is_rpr
          when "t"
            if @inside_is_r
              @is_run_text = +""
              @inside_inline_text = true
            elsif @inside_is || (@current_cell_type == "inlineStr" && !@current_cell_ref.nil?)
              @inside_inline_text = true
            end
          end
        end

        def characters(text)
          @value_buffer << text if @inside_value
          if @inside_inline_text
            if @inside_is_r
              @is_run_text << text
            else
              @inline_text_buffer << text
            end
          end
          @formula_buffer << text if @inside_formula
        end

        def end_element(_uri, local_name, qname)
          name = element_name(local_name, qname)

          case name
          when "v"
            @inside_value = false
          when "f"
            @inside_formula = false
          when "t"
            @inside_inline_text = false
          when "rPr"
            @inside_is_rpr = false
          when "r"
            if @inside_is_r
              run = { text: @is_run_text.dup }
              run[:font] = @is_current_font.dup unless @is_current_font.empty?
              @is_runs << run
              @inside_is_r = false
            end
          when "is"
            @inside_is = false
          when "c"
            store_cell_value
            @cell_phonetic[@current_cell_ref] = true if @current_cell_ph
            @current_cell_ref = nil
            @current_cell_type = nil
            @value_buffer = +""
            @inline_text_buffer = +""
            @formula_buffer = +""
          end
        end

        private

        def parse_row_attributes(attributes)
          row_num = attributes["r"]&.to_i
          return unless row_num

          attrs = {}
          ht = attributes["ht"]
          attrs[:height] = ht.to_f if ht && attributes["customHeight"] == "1"
          attrs[:hidden] = true if attributes["hidden"] == "1"
          ol = attributes["outlineLevel"]
          attrs[:outline_level] = ol.to_i if ol && ol != "0"
          attrs[:collapsed] = true if attributes["collapsed"] == "1"
          s = attributes["s"]
          attrs[:style] = s.to_i if s
          attrs[:thick_top] = true if attributes["thickTop"] == "1"
          attrs[:thick_bot] = true if attributes["thickBot"] == "1"
          attrs[:ph] = true if attributes["ph"] == "1"
          @row_attributes[row_num] = attrs unless attrs.empty?
        end

        def store_cell_value
          return if @current_cell_ref.nil?

          unless @formula_buffer.empty? && @formula_type.nil?
            cached = @value_buffer.empty? ? nil : @value_buffer.dup
            f_type = { "shared" => :shared, "array" => :array, "dataTable" => :data_table }[@formula_type]
            @cells[@current_cell_ref] = Xlsxrb::Elements::Formula.new(
              expression: @formula_buffer.dup,
              cached_value: cached,
              type: f_type,
              ref: @formula_ref,
              shared_index: @formula_si,
              calculate_always: @formula_ca,
              aca: @formula_aca,
              bx: @formula_bx,
              dt2d: @formula_dt2d,
              dtr: @formula_dtr,
              r1: @formula_r1,
              r2: @formula_r2
            )
            return
          end

          case @current_cell_type
          when "inlineStr"
            @cells[@current_cell_ref] = if @is_has_runs
                                          Xlsxrb::Elements::RichText.new(runs: @is_runs.map(&:dup))
                                        else
                                          @inline_text_buffer.dup
                                        end
          when "s"
            index = @value_buffer.to_i
            @cells[@current_cell_ref] = @shared_strings[index] || ""
          when "e"
            code = @value_buffer.dup
            @cells[@current_cell_ref] = if Xlsxrb::Elements::VALID_ERROR_CODES.include?(code)
                                          Xlsxrb::Elements::CellError.new(code:)
                                        else
                                          code
                                        end
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

      # SAX2 listener for parsing workbook.xml to discover sheet names, rIds, and workbook-level properties.
      class WorkbookListener
        include REXML::SAX2Listener

        attr_reader :sheets, :workbook_properties, :workbook_views, :calc_properties, :defined_names,
                    :workbook_protection, :file_version, :file_sharing, :conformance, :file_recovery_properties

        def initialize
          @sheets = []
          @workbook_properties = {}
          @workbook_views = {}
          @calc_properties = {}
          @file_recovery_properties = {}
          @defined_names = []
          @workbook_protection = nil
          @file_version = {}
          @file_sharing = {}
          @conformance = nil
          @inside_defined_name = false
          @current_dn_attrs = nil
          @dn_text_buffer = +""
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          case name
          when "workbook"
            @conformance = attributes["conformance"] if attributes["conformance"]
          when "sheet"
            @sheets << { name: attributes["name"], rid: attributes["r:id"], state: attributes["state"] }
          when "fileVersion"
            an = attributes["appName"]
            @file_version[:app_name] = an if an
            le = attributes["lastEdited"]
            @file_version[:last_edited] = le if le
            loe = attributes["lowestEdited"]
            @file_version[:lowest_edited] = loe if loe
            rb = attributes["rupBuild"]
            @file_version[:rup_build] = rb if rb
            cn = attributes["codeName"]
            @file_version[:code_name] = cn if cn
          when "fileSharing"
            @file_sharing[:read_only_recommended] = true if %w[1 true].include?(attributes["readOnlyRecommended"])
            un = attributes["userName"]
            @file_sharing[:user_name] = un if un
            an = attributes["algorithmName"]
            @file_sharing[:algorithm_name] = an if an
            hv = attributes["hashValue"]
            @file_sharing[:hash_value] = hv if hv
            sv = attributes["saltValue"]
            @file_sharing[:salt_value] = sv if sv
            sc = attributes["spinCount"]
            @file_sharing[:spin_count] = sc.to_i if sc
          when "workbookPr"
            d1904 = attributes["date1904"]
            @workbook_properties[:date1904] = %w[1 true].include?(d1904) unless d1904.nil?
            dtv = attributes["defaultThemeVersion"]
            @workbook_properties[:default_theme_version] = dtv.to_i if dtv
            cn = attributes["codeName"]
            @workbook_properties[:code_name] = cn if cn
            fp = attributes["filterPrivacy"]
            @workbook_properties[:filter_privacy] = %w[1 true].include?(fp) unless fp.nil?
            acp = attributes["autoCompressPictures"]
            @workbook_properties[:auto_compress_pictures] = %w[1 true].include?(acp) unless acp.nil?
            bf = attributes["backupFile"]
            @workbook_properties[:backup_file] = %w[1 true].include?(bf) unless bf.nil?
            so = attributes["showObjects"]
            @workbook_properties[:show_objects] = so if so
            ul = attributes["updateLinks"]
            @workbook_properties[:update_links] = ul if ul
            rac = attributes["refreshAllConnections"]
            @workbook_properties[:refresh_all_connections] = %w[1 true].include?(rac) unless rac.nil?
            cc = attributes["checkCompatibility"]
            @workbook_properties[:check_compatibility] = %w[1 true].include?(cc) unless cc.nil?
            hpfl = attributes["hidePivotFieldList"]
            @workbook_properties[:hide_pivot_field_list] = %w[1 true].include?(hpfl) unless hpfl.nil?
            sbut = attributes["showBorderUnselectedTables"]
            @workbook_properties[:show_border_unselected_tables] = %w[1 true].include?(sbut) unless sbut.nil?
            ps = attributes["promptedSolutions"]
            @workbook_properties[:prompted_solutions] = %w[1 true].include?(ps) unless ps.nil?
            sia = attributes["showInkAnnotation"]
            @workbook_properties[:show_ink_annotation] = %w[1 true].include?(sia) unless sia.nil?
            selv = attributes["saveExternalLinkValues"]
            @workbook_properties[:save_external_link_values] = %w[1 true].include?(selv) unless selv.nil?
            spcf = attributes["showPivotChartFilter"]
            @workbook_properties[:show_pivot_chart_filter] = %w[1 true].include?(spcf) unless spcf.nil?
            arq = attributes["allowRefreshQuery"]
            @workbook_properties[:allow_refresh_query] = %w[1 true].include?(arq) unless arq.nil?
            pi = attributes["publishItems"]
            @workbook_properties[:publish_items] = %w[1 true].include?(pi) unless pi.nil?
            dcompat = attributes["dateCompatibility"]
            @workbook_properties[:date_compatibility] = %w[1 true].include?(dcompat) unless dcompat.nil?
          when "workbookView"
            at = attributes["activeTab"]
            @workbook_views[:active_tab] = at.to_i if at
            fs = attributes["firstSheet"]
            @workbook_views[:first_sheet] = fs.to_i if fs
            vis = attributes["visibility"]
            @workbook_views[:visibility] = vis if vis
            min = attributes["minimized"]
            @workbook_views[:minimized] = %w[1 true].include?(min) unless min.nil?
            shs = attributes["showHorizontalScroll"]
            @workbook_views[:show_horizontal_scroll] = %w[1 true].include?(shs) unless shs.nil?
            svs = attributes["showVerticalScroll"]
            @workbook_views[:show_vertical_scroll] = %w[1 true].include?(svs) unless svs.nil?
            sst = attributes["showSheetTabs"]
            @workbook_views[:show_sheet_tabs] = %w[1 true].include?(sst) unless sst.nil?
            xw = attributes["xWindow"]
            @workbook_views[:x_window] = xw.to_i if xw
            yw = attributes["yWindow"]
            @workbook_views[:y_window] = yw.to_i if yw
            ww = attributes["windowWidth"]
            @workbook_views[:window_width] = ww.to_i if ww
            wh = attributes["windowHeight"]
            @workbook_views[:window_height] = wh.to_i if wh
            tr = attributes["tabRatio"]
            @workbook_views[:tab_ratio] = tr.to_i if tr
            afdg = attributes["autoFilterDateGrouping"]
            @workbook_views[:auto_filter_date_grouping] = %w[1 true].include?(afdg) unless afdg.nil?
          when "calcPr"
            ci = attributes["calcId"]
            @calc_properties[:calc_id] = ci.to_i if ci
            cm = attributes["calcMode"]
            @calc_properties[:calc_mode] = cm if cm
            fcol = attributes["fullCalcOnLoad"]
            @calc_properties[:full_calc_on_load] = %w[1 true].include?(fcol) unless fcol.nil?
            iter = attributes["iterate"]
            @calc_properties[:iterate] = %w[1 true].include?(iter) unless iter.nil?
            ic = attributes["iterateCount"]
            @calc_properties[:iterate_count] = ic.to_i if ic
            id = attributes["iterateDelta"]
            @calc_properties[:iterate_delta] = id.to_f if id
            rm = attributes["refMode"]
            @calc_properties[:ref_mode] = rm if rm
            cc = attributes["calcCompleted"]
            @calc_properties[:calc_completed] = %w[1 true].include?(cc) unless cc.nil?
            cos = attributes["calcOnSave"]
            @calc_properties[:calc_on_save] = %w[1 true].include?(cos) unless cos.nil?
            fprec = attributes["fullPrecision"]
            @calc_properties[:full_precision] = %w[1 true].include?(fprec) unless fprec.nil?
            conc = attributes["concurrentCalc"]
            @calc_properties[:concurrent_calc] = %w[1 true].include?(conc) unless conc.nil?
            cmc = attributes["concurrentManualCount"]
            @calc_properties[:concurrent_manual_count] = cmc.to_i if cmc
            ffc = attributes["forceFullCalc"]
            @calc_properties[:force_full_calc] = %w[1 true].include?(ffc) unless ffc.nil?
          when "fileRecoveryPr"
            ar = attributes["autoRecover"]
            @file_recovery_properties[:auto_recover] = %w[1 true].include?(ar) unless ar.nil?
            cs = attributes["crashSave"]
            @file_recovery_properties[:crash_save] = %w[1 true].include?(cs) unless cs.nil?
            del = attributes["dataExtractLoad"]
            @file_recovery_properties[:data_extract_load] = %w[1 true].include?(del) unless del.nil?
            rl = attributes["repairLoad"]
            @file_recovery_properties[:repair_load] = %w[1 true].include?(rl) unless rl.nil?
          when "workbookProtection"
            prot = {}
            ls = attributes["lockStructure"]
            prot[:lock_structure] = %w[1 true].include?(ls) unless ls.nil?
            lw = attributes["lockWindows"]
            prot[:lock_windows] = %w[1 true].include?(lw) unless lw.nil?
            lr = attributes["lockRevision"]
            prot[:lock_revision] = %w[1 true].include?(lr) unless lr.nil?
            wp = attributes["workbookPassword"]
            prot[:password] = wp if wp
            an = attributes["workbookAlgorithmName"]
            if an
              prot[:algorithm_name] = an
              hv = attributes["workbookHashValue"]
              prot[:hash_value] = hv if hv
              sv = attributes["workbookSaltValue"]
              prot[:salt_value] = sv if sv
              sc = attributes["workbookSpinCount"]
              prot[:spin_count] = sc.to_i if sc
            end
            ran = attributes["revisionsAlgorithmName"]
            if ran
              prot[:revisions_algorithm_name] = ran
              rhv = attributes["revisionsHashValue"]
              prot[:revisions_hash_value] = rhv if rhv
              rsv = attributes["revisionsSaltValue"]
              prot[:revisions_salt_value] = rsv if rsv
              rsc = attributes["revisionsSpinCount"]
              prot[:revisions_spin_count] = rsc.to_i if rsc
            end
            rp = attributes["revisionsPassword"]
            prot[:revisions_password] = rp if rp
            @workbook_protection = prot unless prot.empty?
          when "definedName"
            @inside_defined_name = true
            @current_dn_attrs = {
              name: attributes["name"],
              hidden: %w[1 true].include?(attributes["hidden"])
            }
            lsi = attributes["localSheetId"]
            @current_dn_attrs[:local_sheet_id] = lsi.to_i if lsi
            @current_dn_attrs[:comment] = attributes["comment"] if attributes["comment"]
            @current_dn_attrs[:description] = attributes["description"] if attributes["description"]
            @current_dn_attrs[:function] = true if %w[1 true].include?(attributes["function"])
            @current_dn_attrs[:vb_procedure] = true if %w[1 true].include?(attributes["vbProcedure"])
            @current_dn_attrs[:xlm] = true if %w[1 true].include?(attributes["xlm"])
            @current_dn_attrs[:shortcut_key] = attributes["shortcutKey"] if attributes["shortcutKey"]
            @current_dn_attrs[:publish_to_server] = true if %w[1 true].include?(attributes["publishToServer"])
            @current_dn_attrs[:workbook_parameter] = true if %w[1 true].include?(attributes["workbookParameter"])
            fgi = attributes["functionGroupId"]
            @current_dn_attrs[:function_group_id] = fgi.to_i if fgi
            @current_dn_attrs[:custom_menu] = attributes["customMenu"] if attributes["customMenu"]
            @current_dn_attrs[:help] = attributes["help"] if attributes["help"]
            @current_dn_attrs[:status_bar] = attributes["statusBar"] if attributes["statusBar"]
            @dn_text_buffer = +""
          end
        end

        def characters(text)
          @dn_text_buffer << text if @inside_defined_name
        end

        def end_element(_uri, local_name, qname)
          name = element_name(local_name, qname)
          return unless name == "definedName" && @inside_defined_name

          @current_dn_attrs[:value] = @dn_text_buffer.dup
          @defined_names << @current_dn_attrs
          @inside_defined_name = false
          @current_dn_attrs = nil
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

      # SAX2 listener for parsing .rels files to map rId to Target.
      class RelsListener
        include REXML::SAX2Listener

        attr_reader :relationships

        def initialize
          @relationships = []
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          return unless name == "Relationship"

          @relationships << { id: attributes["Id"], target: attributes["Target"], type: attributes["Type"] }
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

      # SAX2 listener for parsing <cols><col> elements from worksheet XML.
      class ColumnsListener
        include REXML::SAX2Listener

        # Returns { column_index => width } hash (1-based indices).
        attr_reader :raw_columns, :raw_column_attrs

        def initialize
          @raw_columns = {}
          @raw_column_attrs = {}
        end

        def columns
          @raw_columns
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          return unless name == "col"

          min_val = attributes["min"]&.to_i
          max_val = attributes["max"]&.to_i
          width = attributes["width"]&.to_f
          return unless min_val && max_val

          (min_val..max_val).each do |i|
            @raw_columns[i] = width if width
            attrs = {}
            attrs[:hidden] = true if attributes["hidden"] == "1"
            attrs[:best_fit] = true if attributes["bestFit"] == "1"
            ol = attributes["outlineLevel"]
            attrs[:outline_level] = ol.to_i if ol && ol != "0"
            attrs[:collapsed] = true if attributes["collapsed"] == "1"
            s = attributes["style"]
            attrs[:style] = s.to_i if s && s != "0"
            attrs[:phonetic] = true if attributes["phonetic"] == "1"
            @raw_column_attrs[i] = attrs unless attrs.empty?
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

      # SAX2 listener for parsing <mergeCells><mergeCell> elements.
      class MergeCellsListener
        include REXML::SAX2Listener

        attr_reader :ranges

        def initialize
          @ranges = []
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          return unless name == "mergeCell"

          ref = attributes["ref"]
          @ranges << ref if ref
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

      # SAX2 listener for parsing <hyperlinks><hyperlink> elements.
      class HyperlinksListener
        include REXML::SAX2Listener

        attr_reader :links

        def initialize
          @links = []
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          return unless name == "hyperlink"

          ref = attributes["ref"]
          return unless ref

          link = { ref: ref }
          link[:rid] = attributes["r:id"] if attributes["r:id"]
          link[:display] = attributes["display"] if attributes["display"]
          link[:tooltip] = attributes["tooltip"] if attributes["tooltip"]
          link[:location] = attributes["location"] if attributes["location"]
          @links << link
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

      # SAX2 listener for parsing <dimension> element.
      class DimensionListener
        include REXML::SAX2Listener

        attr_reader :ref

        def initialize
          @ref = nil
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          @ref = attributes["ref"] if name == "dimension"
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
end
