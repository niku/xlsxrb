# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Ooxml
    class Writer
      # Mixin containing Table, PivotTable, Document Properties, and External Link XML generation logic.
      module FeaturesXml
        # : (untyped pivot_table, untyped cache_id) -> untyped
        def generate_pivot_table_xml(pivot_table, cache_id)
          data_caption = pivot_table[:data_caption] || (pivot_table[:data_fields].first ? pivot_table[:data_fields].first[:name] : "Values")
          anf = pivot_table.fetch(:apply_number_formats, false)
          abf = pivot_table.fetch(:apply_border_formats, false)
          aff = pivot_table.fetch(:apply_font_formats, false)
          apf = pivot_table.fetch(:apply_pattern_formats, false)
          aaf = pivot_table.fetch(:apply_alignment_formats, false)
          awf = pivot_table.fetch(:apply_width_height_formats, true)

          pt_attrs = {
            xmlns: SSML_NS,
            name: pivot_table[:name] || "",
            cacheId: cache_id || "",
            dataCaption: data_caption || "",
            grandTotalCaption: pivot_table[:grand_total_caption],
            errorCaption: pivot_table[:error_caption],
            showError: pivot_table[:show_error] ? 1 : nil,
            missingCaption: pivot_table[:missing_caption],
            showMissing: pivot_table[:show_missing] == false ? 0 : nil,
            tag: pivot_table[:tag],
            dataOnRows: pivot_table[:data_on_rows] ? 1 : 0,
            rowGrandTotals: pivot_table[:row_grand_totals] == false ? 0 : nil,
            colGrandTotals: pivot_table[:col_grand_totals] == false ? 0 : nil,
            compact: pivot_table[:compact] == false ? 0 : nil,
            outline: pivot_table[:outline] == false ? 0 : nil,
            outlineData: pivot_table[:outline_data] ? 1 : nil,
            compactData: pivot_table[:compact_data] == false ? 0 : nil,
            showHeaders: pivot_table[:show_headers] == false ? 0 : nil,
            showMultipleLabel: pivot_table[:show_multiple_label] == false ? 0 : nil,
            showDataDropDown: pivot_table[:show_data_drop_down] == false ? 0 : nil,
            indent: pivot_table[:indent],
            published: pivot_table[:published] ? 1 : nil,
            editData: pivot_table[:edit_data] ? 1 : nil,
            disableFieldList: pivot_table[:disable_field_list] ? 1 : nil,
            visualTotals: pivot_table[:visual_totals] == false ? 0 : nil,
            printDrill: pivot_table[:print_drill] ? 1 : nil,
            createdVersion: pivot_table[:created_version],
            updatedVersion: pivot_table[:updated_version],
            minRefreshableVersion: pivot_table[:min_refreshable_version],
            applyNumberFormats: anf ? 1 : 0,
            applyBorderFormats: abf ? 1 : 0,
            applyFontFormats: aff ? 1 : 0,
            applyPatternFormats: apf ? 1 : 0,
            applyAlignmentFormats: aaf ? 1 : 0,
            applyWidthHeightFormats: awf ? 1 : 0,
            multipleFieldFilters: pivot_table[:multiple_field_filters] == false ? 0 : nil,
            showDrill: pivot_table[:show_drill] == false ? 0 : nil,
            showDataTips: pivot_table[:show_data_tips] == false ? 0 : nil,
            enableDrill: pivot_table[:enable_drill] == false ? 0 : nil,
            showMemberPropertyTips: pivot_table[:show_member_property_tips] == false ? 0 : nil,
            itemPrintTitles: pivot_table[:item_print_titles] ? 1 : nil,
            fieldPrintTitles: pivot_table[:field_print_titles] ? 1 : nil,
            preserveFormatting: pivot_table[:preserve_formatting] == false ? 0 : nil,
            pageOverThenDown: pivot_table[:page_over_then_down] ? 1 : nil,
            pageWrap: pivot_table[:page_wrap]
          }

          # Compute field count from source range or explicit field_names.
          field_count = if pivot_table[:field_names]
                          pivot_table[:field_names].size
                        else
                          (pivot_table[:row_fields].size + pivot_table[:col_fields].size + pivot_table[:data_fields].size).clamp(1, 100)
                        end

          XmlBuilder.build(declaration: :plain) do |xml|
            xml.tag("pivotTableDefinition", pt_attrs) do
              xml.empty_tag("location", {
                              ref: pivot_table[:dest_ref] || "",
                              firstHeaderRow: pivot_table[:first_header_row] || 1,
                              firstDataRow: pivot_table[:first_data_row] || 1,
                              firstDataCol: pivot_table[:first_data_col] || 1,
                              rowPageCount: pivot_table[:row_page_count],
                              colPageCount: pivot_table[:col_page_count]
                            })
              xml.tag("pivotFields", { count: field_count }) do
                field_count.times do |fi|
                  axis_val = if pivot_table[:row_fields].include?(fi)
                               "axisRow"
                             elsif pivot_table[:col_fields].include?(fi)
                               "axisCol"
                             end
                  fa = pivot_table[:field_attrs] && pivot_table[:field_attrs][fi]
                  pf_attrs = {
                    axis: axis_val,
                    dataField: pivot_table[:data_fields].any? { |df| df[:fld] == fi } ? 1 : nil,
                    compact: xml_bool_attr(fa&.[](:compact)),
                    outline: xml_bool_attr(fa&.[](:outline)),
                    subtotalTop: xml_bool_attr(fa&.[](:subtotal_top)),
                    showAll: fa && fa[:show_all] == true ? "1" : "0",
                    numFmtId: fa && fa[:num_fmt_id] ? fa[:num_fmt_id] : nil,
                    sortType: fa && fa[:sort_type] ? fa[:sort_type] : nil,
                    defaultSubtotal: fa && fa[:default_subtotal] == false ? 0 : nil,
                    insertBlankRow: fa && fa[:insert_blank_row] ? 1 : nil,
                    insertPageBreak: fa && fa[:insert_page_break] ? 1 : nil,
                    includeNewItemsInFilter: fa && fa[:include_new_items_in_filter] ? 1 : nil
                  }
                  field_items = pivot_table[:items] && pivot_table[:items][fi]
                  axis_field = pivot_table[:row_fields].include?(fi) || pivot_table[:col_fields].include?(fi)
                  # The default subtotal item goes with the default subtotal; a field with
                  # defaultSubtotal="0" must not list it.
                  default_subtotal = !(fa && fa[:default_subtotal] == false)
                  if field_items || (axis_field && default_subtotal)
                    # A row or column field needs an <items> list; without known items it holds only the
                    # default subtotal, which Excel fills in when it refreshes the cache on open.
                    field_items ||= []
                    xml.tag("pivotField", pf_attrs) do
                      xml.tag("items", { count: field_items.size + (default_subtotal ? 1 : 0) }) do
                        field_items.size.times { |ix| xml.empty_tag("item", { x: ix }) }
                        xml.empty_tag("item", { t: "default" }) if default_subtotal
                      end
                    end
                  else
                    xml.empty_tag("pivotField", pf_attrs)
                  end
                end
              end

              unless pivot_table[:row_fields].empty?
                xml.tag("rowFields", { count: pivot_table[:row_fields].size }) do
                  pivot_table[:row_fields].each { |f| xml.empty_tag("field", { x: f || "" }) }
                end
              end

              unless pivot_table[:col_fields].empty?
                xml.tag("colFields", { count: pivot_table[:col_fields].size }) do
                  pivot_table[:col_fields].each { |f| xml.empty_tag("field", { x: f || "" }) }
                end
              end

              unless pivot_table[:data_fields].empty?
                xml.tag("dataFields", { count: pivot_table[:data_fields].size }) do
                  pivot_table[:data_fields].each do |df|
                    xml.empty_tag("dataField", {
                                    name: df[:name] || "",
                                    fld: df[:fld] || "",
                                    subtotal: df[:subtotal] || "sum",
                                    showDataAs: df[:show_data_as],
                                    baseField: df[:base_field],
                                    baseItem: df[:base_item],
                                    numFmtId: df[:num_fmt_id]
                                  })
                  end
                end
              end

              if (psi = pivot_table[:pivot_table_style])
                xml.empty_tag("pivotTableStyleInfo", {
                                name: psi[:name],
                                showRowHeaders: xml_bool_attr(psi[:show_row_headers]),
                                showColHeaders: xml_bool_attr(psi[:show_col_headers]),
                                showRowStripes: xml_bool_attr(psi[:show_row_stripes]),
                                showColStripes: xml_bool_attr(psi[:show_col_stripes]),
                                showLastColumn: xml_bool_attr(psi[:show_last_column])
                              })
              end
            end
          end
        end

        # : (untyped pivot_table, untyped _cache_id) -> untyped
        def generate_pivot_cache_definition_xml(pivot_table, _cache_id)
          pcd_attrs = {
            xmlns: SSML_NS,
            "xmlns:r": DOC_REL_NS,
            "r:id": "rId1",
            refreshOnLoad: 1,
            saveData: pivot_table[:cache_save_data] == false ? 0 : nil,
            enableRefresh: pivot_table[:cache_enable_refresh] == false ? 0 : nil,
            refreshedBy: pivot_table[:cache_refreshed_by],
            refreshedVersion: pivot_table[:cache_refreshed_version],
            createdVersion: pivot_table[:cache_created_version],
            recordCount: pivot_table[:cache_record_count],
            optimizeMemory: pivot_table[:cache_optimize_memory] ? 1 : nil
          }

          XmlBuilder.build(declaration: :plain) do |xml|
            xml.tag("pivotCacheDefinition", pcd_attrs) do
              # Parse source ref: "Sheet1!A1:C4" => sheet name + range.
              source = pivot_table[:source_ref]
              ws_attrs = {}
              if source&.include?("!")
                sname, srange = source.split("!", 2)
                ws_attrs[:ref] = srange
                ws_attrs[:sheet] = sname.delete("'")
              elsif !source.nil?
                ws_attrs[:ref] = source
              end
              ws_attrs[:name] = pivot_table[:source_name] if pivot_table[:source_name]
              xml.tag("cacheSource", { type: "worksheet" }) do
                xml.empty_tag("worksheetSource", ws_attrs)
              end

              field_count = pivot_table[:field_names] ? pivot_table[:field_names].size : (pivot_table[:row_fields].size + pivot_table[:col_fields].size + pivot_table[:data_fields].size)
              xml.tag("cacheFields", { count: field_count }) do
                field_count.times do |fi|
                  fname = if pivot_table[:field_names] && pivot_table[:field_names][fi]
                            pivot_table[:field_names][fi]
                          else
                            df = pivot_table[:data_fields].find { |d| d[:fld] == fi }
                            df ? df[:name] : "Field#{fi + 1}"
                          end
                  fa = pivot_table[:field_attrs] && pivot_table[:field_attrs][fi]
                  cf_attrs = {
                    name: fname || "",
                    numFmtId: (fa && fa[:cache_num_fmt_id]) || 0,
                    caption: fa && fa[:cache_caption] ? fa[:cache_caption] : nil,
                    formula: fa && fa[:cache_formula] ? fa[:cache_formula] : nil
                  }
                  field_items = pivot_table[:items] && pivot_table[:items][fi]
                  xml.tag("cacheField", cf_attrs) do
                    if field_items
                      write_shared_items_xml(xml, field_items)
                    else
                      xml.empty_tag("sharedItems")
                    end
                  end
                end
              end
            end
          end
        end

        # Writes <sharedItems> and its typed children for a cache field.
        # : (XmlBuilder xml, Array[untyped] values) -> void
        def write_shared_items_xml(xml, values)
          numbers = values.grep(Numeric)
          strings = values.any? { |v| !v.nil? && !v.is_a?(Numeric) && v != true && v != false }
          attrs = {
            containsSemiMixedTypes: strings || values.include?(nil) ? nil : 0,
            containsString: strings ? nil : 0,
            containsBlank: values.include?(nil) ? 1 : nil,
            containsNumber: numbers.empty? ? nil : 1,
            containsInteger: !numbers.empty? && numbers.all? { |n| n.is_a?(Integer) || n == n.to_i } ? 1 : nil,
            minValue: numbers.empty? ? nil : numbers.min,
            maxValue: numbers.empty? ? nil : numbers.max,
            count: values.size
          }
          xml.tag("sharedItems", attrs) do
            values.each do |value|
              case value
              when nil then xml.empty_tag("m")
              when true, false then xml.empty_tag("b", { v: value ? 1 : 0 })
              when Numeric then xml.empty_tag("n", { v: value })
              else xml.empty_tag("s", { v: value.to_s })
              end
            end
          end
        end

        # : (untyped pivot_table) -> untyped
        def generate_pivot_cache_records_xml(pivot_table)
          items = pivot_table[:items]
          XmlBuilder.build(declaration: :plain) do |xml|
            if items&.values&.any? { |v| v && !v.empty? }
              max_len = items.values.map { |v| v ? v.size : 0 }.max
              field_count = pivot_table[:field_names] ? pivot_table[:field_names].size : (pivot_table[:row_fields].size + pivot_table[:col_fields].size + pivot_table[:data_fields].size)
              xml.tag("pivotCacheRecords", { xmlns: SSML_NS, count: max_len }) do
                max_len.times do |ri|
                  xml.tag("r") do
                    field_count.times do |fi|
                      field_items = items[fi]
                      if field_items
                        xml.empty_tag("x", { v: ri < field_items.size ? ri : 0 })
                      else
                        xml.empty_tag("n", { v: 0 })
                      end
                    end
                  end
                end
              end
            else
              xml.empty_tag("pivotCacheRecords", { xmlns: SSML_NS, count: 0 })
            end
          end
        end

        # : (untyped cache_id) -> untyped
        def generate_pivot_cache_rels(cache_id)
          XmlBuilder.build(declaration: :plain) do |xml|
            xml.tag("Relationships", { xmlns: REL_NS }) do
              xml.empty_tag("Relationship", {
                              Id: "rId1",
                              Type: "#{DOC_REL_NS}/pivotCacheRecords",
                              Target: "pivotCacheRecords#{cache_id}.xml"
                            })
            end
          end
        end

        # : (untyped cache_id) -> untyped
        def generate_pivot_table_rels(cache_id)
          XmlBuilder.build(declaration: :plain) do |xml|
            xml.tag("Relationships", { xmlns: REL_NS }) do
              xml.empty_tag("Relationship", {
                              Id: "rId1",
                              Type: "#{DOC_REL_NS}/pivotCacheDefinition",
                              Target: "../pivotCache/pivotCacheDefinition#{cache_id}.xml"
                            })
            end
          end
        end

        # : (untyped ext_link) -> untyped
        def generate_external_link_xml(ext_link)
          XmlBuilder.build(declaration: :plain) do |xml|
            xml.tag("externalLink", { xmlns: SSML_NS, "xmlns:r": DOC_REL_NS }) do
              xml.tag("externalBook", { "r:id": "rId1" }) do
                unless ext_link[:sheet_names].empty?
                  xml.tag("sheetNames") do
                    ext_link[:sheet_names].each { |sn| xml.empty_tag("sheetName", { val: sn || "" }) }
                  end
                end
              end
            end
          end
        end

        # : (untyped ext_link) -> untyped
        def generate_external_link_rels(ext_link)
          XmlBuilder.build(declaration: :plain) do |xml|
            xml.tag("Relationships", { xmlns: REL_NS }) do
              xml.empty_tag("Relationship", {
                              Id: "rId1",
                              Type: "#{DOC_REL_NS}/externalLinkPath",
                              Target: ext_link[:target] || "",
                              TargetMode: "External"
                            })
            end
          end
        end

        # : (untyped ct_xml) -> untyped
        def parse_extra_content_types(ct_xml)
          ct_xml.scan(/<Default\s+Extension="([^"]+)"\s+ContentType="([^"]+)"/).each do |ext, ct|
            @extra_ct_defaults[ext] ||= ct
          end
          ct_xml.scan(/<Override\s+PartName="([^"]+)"\s+ContentType="([^"]+)"/).each do |pn, ct|
            @extra_ct_overrides[pn] ||= ct
          end
        end

        # : (XmlBuilder xml, untyped element, untyped opts) -> void
        def write_line_end_xml(xml, element, opts)
          return unless opts

          xml.empty_tag("a:#{element}", {
                          type: opts[:type],
                          w: opts[:w],
                          len: opts[:len]
                        })
        end

        # : (XmlBuilder xml, untyped rich_text) -> void
        def write_rich_text_xml(xml, rich_text)
          rich_text.runs.each do |run|
            font = run.respond_to?(:font) ? run.font : run[:font]
            text = run.respond_to?(:text) ? run.text : run[:text]
            xml.tag("r") do
              if font && !font.empty?
                xml.tag("rPr") do
                  xml.empty_tag("b") if font[:bold]
                  xml.empty_tag("i") if font[:italic]
                  xml.empty_tag("strike") if font[:strike]
                  if font[:underline]
                    u_val = font[:underline]
                    if u_val == true || u_val.to_s == "single"
                      xml.empty_tag("u")
                    else
                      mapped_u = StyleBuilder::UNDERLINE_MAP[u_val.to_sym] || u_val
                      xml.empty_tag("u", { val: mapped_u })
                    end
                  end
                  va_val = font[:vert_align] || font[:vertical_align]
                  va_val ||= "superscript" if font[:superscript]
                  va_val ||= "subscript" if font[:subscript]
                  if va_val
                    mapped_va = StyleBuilder::VERT_ALIGN_MAP[va_val.to_sym] || va_val
                    xml.empty_tag("vertAlign", { val: mapped_va })
                  end
                  font_sz = font[:size] || font[:sz]
                  xml.empty_tag("sz", { val: font_sz }) if font_sz
                  write_color_xml(xml, font)
                  xml.empty_tag("rFont", { val: font[:name] }) if font[:name]
                  xml.empty_tag("family", { val: font[:family] }) if font[:family]
                  xml.empty_tag("scheme", { val: font[:scheme] }) if font[:scheme]
                end
              end
              xml.tag("t") { xml.xstring_text(text) }
            end
          end
        end

        # : (XmlBuilder xml, untyped cell_ref, untyped value, untyped style_idx, ?untyped? sst, ?ph: untyped?) -> void
        def write_cell_xml(xml, cell_ref, value, style_idx, sst = nil, ph: nil) # rubocop:disable Naming/MethodParameterName
          ph_val = ph ? 1 : nil
          case value
          when Xlsxrb::Elements::CellError
            xml.tag("c", { r: cell_ref || "", t: "e", s: style_idx, ph: ph_val }) do
              xml.text_tag("v", value.code)
            end
          when Xlsxrb::Elements::Formula
            f_attrs = case value.type
                      when :shared
                        { t: "shared", si: value.shared_index || "", ref: value.ref }
                      when :array
                        value.ref ? { t: "array", ref: value.ref } : {}
                      when :data_table
                        {
                          t: "dataTable",
                          dt2D: value.dt2d ? 1 : nil,
                          dtr: value.dtr ? 1 : nil,
                          r1: value.r1,
                          r2: value.r2
                        }
                      else
                        {}
                      end
            f_attrs[:ca] = 1 if value.calculate_always?
            f_attrs[:aca] = 1 if value.aca
            f_attrs[:bx] = 1 if value.bx
            t_val = nil
            cached_val_str = value.cached_value.to_s
            case value.cached_value
            when String
              t_val = "str"
            when true
              t_val = "b"
              cached_val_str = "1"
            when false
              t_val = "b"
              cached_val_str = "0"
            end
            xml.tag("c", { r: cell_ref || "", t: t_val, s: style_idx, ph: ph_val }) do
              if value.expression && !value.expression.empty?
                xml.text_tag("f", value.expression, f_attrs)
              else
                xml.empty_tag("f", f_attrs)
              end
              xml.tag("v") { xml.xstring_text(cached_val_str) } unless value.cached_value.nil?
            end
          when Xlsxrb::Elements::RichText
            if sst
              rt_sst, = sst
              xml.tag("c", { r: cell_ref || "", t: "s", s: style_idx, ph: ph_val }) do
                xml.text_tag("v", rt_sst[value])
              end
            else
              xml.tag("c", { r: cell_ref || "", t: "inlineStr", s: style_idx, ph: ph_val }) do
                xml.tag("is") { write_rich_text_xml(xml, value) }
              end
            end
          when true, false
            xml.tag("c", { r: cell_ref || "", t: "b", s: style_idx, ph: ph_val }) do
              xml.text_tag("v", value ? 1 : 0)
            end
          when Time
            date1904 = @workbook_properties&.[](:date1904) ? true : false
            serial = Xlsxrb::Ooxml::Utils.datetime_to_serial(value, date1904: date1904)
            s_val = style_idx || resolve_style_index(datetime_num_fmt_id)
            xml.tag("c", { r: cell_ref || "", s: s_val, ph: ph_val }) { xml.text_tag("v", serial) }
          when Date
            date1904 = @workbook_properties&.[](:date1904) ? true : false
            serial = Xlsxrb::Ooxml::Utils.date_to_serial(value, date1904: date1904)
            s_val = style_idx || resolve_style_index(date_num_fmt_id)
            xml.tag("c", { r: cell_ref || "", s: s_val, ph: ph_val }) { xml.text_tag("v", serial) }
          when Numeric
            xml.tag("c", { r: cell_ref || "", s: style_idx, ph: ph_val }) { xml.text_tag("v", value) }
          else
            if sst
              _, str_sst = sst
              xml.tag("c", { r: cell_ref || "", t: "s", s: style_idx, ph: ph_val }) do
                xml.text_tag("v", str_sst[value.to_s])
              end
            else
              xml.tag("c", { r: cell_ref || "", t: "inlineStr", s: style_idx, ph: ph_val }) do
                xml.tag("is") { xml.tag("t") { xml.xstring_text(value) } }
              end
            end
          end
        end

        # Returns the numFmtId for dates, registering it on first use.
        # : () -> untyped
        def date_num_fmt_id
          @date_num_fmt_id ||= add_number_format(Xlsxrb::Ooxml::Utils::DEFAULT_DATE_FORMAT)
        end

        # Returns the numFmtId for datetime, registering it on first use.
        # : () -> untyped
        def datetime_num_fmt_id
          @datetime_num_fmt_id ||= add_number_format(Xlsxrb::Ooxml::Utils::DEFAULT_DATETIME_FORMAT)
        end

        # Maps a numFmtId to a cellXfs index. Index 0 is the default (no format).
        # : (untyped style_value) -> (nil | untyped)
        def resolve_style_index(style_value)
          return nil if style_value.nil?

          # New-style: { xf_index: N } from set_cell_style.
          return style_value[:xf_index] if style_value.is_a?(Hash) && style_value.key?(:xf_index)

          # Legacy: raw num_fmt_id from set_cell_format — find or create matching xf entry.
          num_fmt_id = style_value
          @xf_index_map ||= begin
            map = {}
            @num_fmts.each_with_index do |nf, _i|
              entry = { num_fmt_id: nf[:num_fmt_id], font_id: 0, fill_id: 0, border_id: 0 }
              idx = @xf_entries.index(entry)
              unless idx
                @xf_entries << entry
                idx = @xf_entries.size - 1
              end
              map[nf[:num_fmt_id]] = idx
            end
            map
          end
          @xf_index_map[num_fmt_id]
        end

        # : (XmlBuilder xml, untyped filter) -> void
        def write_filter_xml(xml, filter)
          case filter[:type]
          when :filters
            attrs = {
              blank: filter[:blank] ? 1 : nil,
              calendarType: filter[:calendar_type]
            }
            has_values = filter[:values]&.any?
            has_date_groups = filter[:date_group_items]&.any?
            if has_values || has_date_groups
              xml.tag("filters", attrs) do
                filter[:values]&.each { |v| xml.empty_tag("filter", { val: v || "" }) }
                filter[:date_group_items]&.each do |dg|
                  xml.empty_tag("dateGroupItem", {
                                  dateTimeGrouping: dg[:date_time_grouping] || "",
                                  year: dg[:year],
                                  month: dg[:month],
                                  day: dg[:day],
                                  hour: dg[:hour],
                                  minute: dg[:minute],
                                  second: dg[:second]
                                })
                end
              end
            else
              xml.empty_tag("filters", attrs)
            end
          when :custom
            if filter[:filters]
              xml.tag("customFilters", { and: filter[:and] ? 1 : nil }) do
                filter[:filters].each do |cf|
                  xml.empty_tag("customFilter", { operator: cf[:operator] || "", val: cf[:val] || "" })
                end
              end
            else
              xml.tag("customFilters") do
                xml.empty_tag("customFilter", { operator: filter[:operator] || "", val: filter[:val] || "" })
              end
            end
          when :dynamic
            xml.empty_tag("dynamicFilter", {
                            type: filter[:dynamic_type] || "",
                            val: filter[:val],
                            valIso: filter[:val_iso],
                            maxVal: filter[:max_val],
                            maxValIso: filter[:max_val_iso]
                          })
          when :top10
            xml.empty_tag("top10", {
                            top: filter[:top] ? 1 : nil,
                            percent: filter[:percent] ? 1 : nil,
                            val: filter[:val] || "",
                            filterVal: filter[:filter_val]
                          })
          when :color_filter
            xml.empty_tag("colorFilter", {
                            dxfId: filter[:dxf_id] || "",
                            cellColor: filter[:cell_color] == false ? 0 : nil
                          })
          when :icon_filter
            xml.empty_tag("iconFilter", {
                            iconSet: filter[:icon_set] || "",
                            iconId: filter[:icon_id]
                          })
          end
        end

        # : (XmlBuilder xml, untyped rule) -> void
        def write_cf_rule(xml, rule)
          type = rule[:type]
          if type.is_a?(String) || type.is_a?(Symbol)
            t_str = type.to_s
            snake = t_str.gsub(/([A-Z]+)([A-Z][a-z])/, '\1_\2')
                         .gsub(/([a-z\d])([A-Z])/, '\1_\2')
                         .downcase.to_sym
            type = snake if CF_TYPE_MAP.key?(snake)
          end
          rule_type = CF_TYPE_MAP[type] || type.to_s
          did = rule[:dxf_id] || rule[:format_id]
          tp_str = if rule[:time_period]
                     tp = rule[:time_period]
                     DslHelpers::CF_TIME_PERIODS[tp.to_sym] || tp.to_s
                   end
          rule_attrs = {
            type: rule_type,
            priority: rule[:priority],
            operator: rule[:operator],
            dxfId: did,
            stopIfTrue: rule[:stop_if_true] ? 1 : nil,
            aboveAverage: rule[:above_average] == false ? 0 : nil,
            equalAverage: rule[:equal_average] ? 1 : nil,
            rank: rule[:rank],
            percent: rule[:percent] ? 1 : nil,
            bottom: rule[:bottom] ? 1 : nil,
            text: rule[:text],
            timePeriod: tp_str,
            stdDev: rule[:std_dev]
          }

          case type
          when :cell_is, :expression, :above_average, :top10, :duplicate_values, :unique_values,
               :contains_text, :not_contains_text, :begins_with, :ends_with,
               :contains_blanks, :not_contains_blanks, :contains_errors, :not_contains_errors, :time_period
            formulas = rule[:formulas] || [rule[:formula]].compact
            if formulas.empty?
              xml.empty_tag("cfRule", rule_attrs)
            else
              xml.tag("cfRule", rule_attrs) do
                formulas.each { |f| xml.text_tag("formula", f) }
              end
            end
          when :color_scale
            cs = rule[:color_scale]
            xml.tag("cfRule", rule_attrs) do
              xml.tag("colorScale") do
                cs[:cfvo]&.each do |cfvo|
                  xml.empty_tag("cfvo", {
                                  type: cfvo[:type] || "",
                                  val: cfvo[:val],
                                  gte: cfvo[:gte] == false ? 0 : nil
                                })
                end
                cs[:colors]&.each { |c| write_cf_color_xml(xml, c) }
              end
            end
          when :data_bar
            db = rule[:data_bar]
            xml.tag("cfRule", rule_attrs) do
              xml.tag("dataBar", {
                        minLength: db[:min_length],
                        maxLength: db[:max_length],
                        showValue: xml_bool_attr(db[:show_value])
                      }) do
                db[:cfvo]&.each do |cfvo|
                  xml.empty_tag("cfvo", {
                                  type: cfvo[:type] || "",
                                  val: cfvo[:val],
                                  gte: cfvo[:gte] == false ? 0 : nil
                                })
                end
                write_cf_color_xml(xml, db[:color]) if db[:color]
              end
            end
          when :icon_set
            is = rule[:icon_set]
            xml.tag("cfRule", rule_attrs) do
              xml.tag("iconSet", {
                        iconSet: is[:icon_set],
                        reverse: xml_bool_attr(is[:reverse]),
                        percent: xml_bool_attr(is[:percent]),
                        showValue: xml_bool_attr(is[:show_value])
                      }) do
                is[:cfvo]&.each do |cfvo|
                  xml.empty_tag("cfvo", {
                                  type: cfvo[:type] || "",
                                  val: cfvo[:val],
                                  gte: cfvo[:gte] == false ? 0 : nil
                                })
                end
              end
            end
          else
            xml.empty_tag("cfRule", rule_attrs)
          end
        end

        # Emits a <color> element for CF rules, accepting either a plain RGB string or a hash.
        # : (XmlBuilder xml, untyped color) -> void
        def write_cf_color_xml(xml, color)
          if color.is_a?(Hash)
            write_color_xml(xml, color)
          else
            xml.empty_tag("color", { rgb: color || "" })
          end
        end

        # : (untyped sheet_cells) -> ("A1" | ::String)
        def compute_dimension(sheet_cells)
          return "A1" if sheet_cells.empty?

          min_col = Float::INFINITY
          max_col = 0
          min_row = Float::INFINITY
          max_row = 0
          sheet_cells.each_key do |addr|
            col_letter = extract_column_letter(addr)
            row_num = extract_row_number(addr)
            col_idx = column_letter_to_index(col_letter)
            min_col = col_idx if col_idx < min_col
            max_col = col_idx if col_idx > max_col
            min_row = row_num if row_num < min_row
            max_row = row_num if row_num > max_row
          end
          start_col = index_to_column_letter(min_col)
          end_col = index_to_column_letter(max_col)
          "#{start_col}#{min_row}:#{end_col}#{max_row}"
        end

        # : (untyped index) -> untyped
        def index_to_column_letter(index)
          result = +""
          while index.positive?
            index -= 1
            result.prepend(("A".ord + (index % 26)).chr)
            index /= 26
          end
          result
        end

        # : () -> untyped
        def generate_core_properties_xml
          XmlBuilder.build(declaration: :plain) do |xml|
            xml.tag("cp:coreProperties", {
                      "xmlns:cp": CP_NS,
                      "xmlns:dc": DC_NS,
                      "xmlns:dcterms": DCTERMS_NS,
                      "xmlns:xsi": XSI_NS
                    }) do
              xml.text_tag("dc:title", @core_properties[:title]) if @core_properties[:title]
              xml.text_tag("dc:subject", @core_properties[:subject]) if @core_properties[:subject]
              xml.text_tag("dc:creator", @core_properties[:creator]) if @core_properties[:creator]
              xml.text_tag("cp:keywords", @core_properties[:keywords]) if @core_properties[:keywords]
              xml.text_tag("dc:description", @core_properties[:description]) if @core_properties[:description]
              xml.text_tag("cp:lastModifiedBy", @core_properties[:last_modified_by]) if @core_properties[:last_modified_by]
              xml.text_tag("cp:revision", @core_properties[:revision]) if @core_properties[:revision]
              xml.text_tag("dcterms:created", @core_properties[:created], { "xsi:type": "dcterms:W3CDTF" }) if @core_properties[:created]
              xml.text_tag("dcterms:modified", @core_properties[:modified], { "xsi:type": "dcterms:W3CDTF" }) if @core_properties[:modified]
              xml.text_tag("cp:category", @core_properties[:category]) if @core_properties[:category]
              xml.text_tag("cp:contentStatus", @core_properties[:content_status]) if @core_properties[:content_status]
              xml.text_tag("dc:language", @core_properties[:language]) if @core_properties[:language]
            end
          end
        end

        # : () -> untyped
        def generate_app_properties_xml
          app_val = @app_properties[:app_name] || @app_properties[:application]
          XmlBuilder.build(declaration: :plain) do |xml|
            xml.tag("Properties", { xmlns: APP_NS, "xmlns:vt": VT_NS }) do
              xml.text_tag("Application", app_val) if app_val
              xml.text_tag("AppVersion", @app_properties[:app_version]) if @app_properties[:app_version]
              if @app_properties[:heading_pairs] && @app_properties[:titles_of_parts]
                hp = @app_properties[:heading_pairs]
                tp = @app_properties[:titles_of_parts]
                xml.tag("HeadingPairs") do
                  xml.tag("vt:vector", { size: hp.size * 2, baseType: "variant" }) do
                    hp.each do |label, count|
                      xml.tag("vt:variant") { xml.text_tag("vt:lpstr", label) }
                      xml.tag("vt:variant") { xml.text_tag("vt:i4", count) }
                    end
                  end
                end
                xml.tag("TitlesOfParts") do
                  xml.tag("vt:vector", { size: tp.size, baseType: "lpstr" }) do
                    tp.each { |t| xml.text_tag("vt:lpstr", t) }
                  end
                end
              end
            end
          end
        end

        # : () -> untyped
        def generate_custom_properties_xml
          custom_ns = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
          list = if @custom_properties.is_a?(Hash)
                   @custom_properties.map do |name, val|
                     if val.is_a?(Elements::CustomProperty) || (val.respond_to?(:name) && val.respond_to?(:value))
                       { name: val.name || name.to_s, value: val.value, type: val.type }
                     elsif val.is_a?(Hash)
                       { name: val[:name] || name.to_s, value: val[:value], type: val[:type] }
                     else
                       { name: name.to_s, value: val, type: nil }
                     end
                   end
                 else
                   @custom_properties.map do |prop|
                     if prop.is_a?(Elements::CustomProperty) || (prop.respond_to?(:name) && prop.respond_to?(:value))
                       { name: prop.name, value: prop.value, type: prop.type }
                     else
                       prop
                     end
                   end
                 end

          XmlBuilder.build(declaration: :plain) do |xml|
            xml.tag("Properties", { xmlns: custom_ns, "xmlns:vt": VT_NS }) do
              list.each_with_index do |prop, idx|
                fmtid = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
                xml.tag("property", { fmtid: fmtid, pid: idx + 2, name: prop[:name] || "" }) do
                  case prop[:type]
                  when :number, :int, :i4
                    xml.text_tag("vt:i4", prop[:value])
                  when :float, :r8
                    xml.text_tag("vt:r8", prop[:value])
                  when :bool, :boolean
                    xml.text_tag("vt:bool", prop[:value] ? "true" : "false")
                  when :date, :filetime
                    xml.text_tag("vt:filetime", prop[:value])
                  else
                    xml.text_tag("vt:lpwstr", prop[:value])
                  end
                end
              end
            end
          end
        end

        private

        # : (untyped val) -> (nil | ::Integer)
        def xml_bool_attr(val)
          return nil if val.nil?

          val ? 1 : 0
        end
      end
    end
  end
end
