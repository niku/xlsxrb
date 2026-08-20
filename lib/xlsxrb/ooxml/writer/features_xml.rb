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
          pt_attrs = %( xmlns="#{SSML_NS}" name="#{xml_escape(pivot_table[:name])}" cacheId="#{cache_id}" dataCaption="#{xml_escape(data_caption)}")
          pt_attrs << %( grandTotalCaption="#{xml_escape(pivot_table[:grand_total_caption])}") if pivot_table[:grand_total_caption]
          pt_attrs << %( errorCaption="#{xml_escape(pivot_table[:error_caption])}") if pivot_table[:error_caption]
          pt_attrs << ' showError="1"' if pivot_table[:show_error]
          pt_attrs << %( missingCaption="#{xml_escape(pivot_table[:missing_caption])}") if pivot_table[:missing_caption]
          pt_attrs << ' showMissing="0"' if pivot_table[:show_missing] == false
          pt_attrs << %( tag="#{xml_escape(pivot_table[:tag])}") if pivot_table[:tag]
          pt_attrs << %( dataOnRows="1") if pivot_table[:data_on_rows]
          pt_attrs << %( dataOnRows="0") unless pivot_table[:data_on_rows]
          pt_attrs << %( rowGrandTotals="0") if pivot_table[:row_grand_totals] == false
          pt_attrs << %( colGrandTotals="0") if pivot_table[:col_grand_totals] == false
          pt_attrs << %( compact="0") if pivot_table[:compact] == false
          pt_attrs << %( outline="0") if pivot_table[:outline] == false
          pt_attrs << %( outlineData="1") if pivot_table[:outline_data]
          pt_attrs << %( compactData="0") if pivot_table[:compact_data] == false
          pt_attrs << %( showHeaders="0") if pivot_table[:show_headers] == false
          pt_attrs << %( showMultipleLabel="0") if pivot_table[:show_multiple_label] == false
          pt_attrs << %( showDataDropDown="0") if pivot_table[:show_data_drop_down] == false
          pt_attrs << %( indent="#{pivot_table[:indent]}") if pivot_table[:indent]
          pt_attrs << ' published="1"' if pivot_table[:published]
          pt_attrs << ' editData="1"' if pivot_table[:edit_data]
          pt_attrs << ' disableFieldList="1"' if pivot_table[:disable_field_list]
          pt_attrs << ' visualTotals="0"' if pivot_table[:visual_totals] == false
          pt_attrs << ' printDrill="1"' if pivot_table[:print_drill]
          pt_attrs << %( createdVersion="#{pivot_table[:created_version]}") if pivot_table[:created_version]
          pt_attrs << %( updatedVersion="#{pivot_table[:updated_version]}") if pivot_table[:updated_version]
          pt_attrs << %( minRefreshableVersion="#{pivot_table[:min_refreshable_version]}") if pivot_table[:min_refreshable_version]
          anf = pivot_table.fetch(:apply_number_formats, false)
          abf = pivot_table.fetch(:apply_border_formats, false)
          aff = pivot_table.fetch(:apply_font_formats, false)
          apf = pivot_table.fetch(:apply_pattern_formats, false)
          aaf = pivot_table.fetch(:apply_alignment_formats, false)
          awf = pivot_table.fetch(:apply_width_height_formats, true)
          pt_attrs << %( applyNumberFormats="#{anf ? 1 : 0}")
          pt_attrs << %( applyBorderFormats="#{abf ? 1 : 0}")
          pt_attrs << %( applyFontFormats="#{aff ? 1 : 0}")
          pt_attrs << %( applyPatternFormats="#{apf ? 1 : 0}")
          pt_attrs << %( applyAlignmentFormats="#{aaf ? 1 : 0}")
          pt_attrs << %( applyWidthHeightFormats="#{awf ? 1 : 0}")
          pt_attrs << ' multipleFieldFilters="0"' if pivot_table[:multiple_field_filters] == false
          pt_attrs << ' showDrill="0"' if pivot_table[:show_drill] == false
          pt_attrs << ' showDataTips="0"' if pivot_table[:show_data_tips] == false
          pt_attrs << ' enableDrill="0"' if pivot_table[:enable_drill] == false
          pt_attrs << ' showMemberPropertyTips="0"' if pivot_table[:show_member_property_tips] == false
          pt_attrs << ' itemPrintTitles="1"' if pivot_table[:item_print_titles]
          pt_attrs << ' fieldPrintTitles="1"' if pivot_table[:field_print_titles]
          pt_attrs << ' preserveFormatting="0"' if pivot_table[:preserve_formatting] == false
          pt_attrs << ' pageOverThenDown="1"' if pivot_table[:page_over_then_down]
          pt_attrs << %( pageWrap="#{pivot_table[:page_wrap]}") if pivot_table[:page_wrap]
          parts = [
            XML_HEADER,
            "<pivotTableDefinition#{pt_attrs}>"
          ]

          # Compute field count from source range or explicit field_names.
          field_count = if pivot_table[:field_names]
                          pivot_table[:field_names].size
                        else
                          (pivot_table[:row_fields].size + pivot_table[:col_fields].size + pivot_table[:data_fields].size).clamp(1, 100)
                        end
          loc_attrs = %(<location ref="#{pivot_table[:dest_ref]}" firstHeaderRow="1" firstDataRow="1" firstDataCol="1")
          loc_attrs << %( rowPageCount="#{pivot_table[:row_page_count]}") if pivot_table[:row_page_count]
          loc_attrs << %( colPageCount="#{pivot_table[:col_page_count]}") if pivot_table[:col_page_count]
          loc_attrs << "/>"
          parts << loc_attrs
          parts << %(<pivotFields count="#{field_count}">)
          field_count.times do |fi|
            attrs = +""
            if pivot_table[:row_fields].include?(fi)
              attrs << ' axis="axisRow"'
            elsif pivot_table[:col_fields].include?(fi)
              attrs << ' axis="axisCol"'
            end
            attrs << ' dataField="1"' if pivot_table[:data_fields].any? { |df| df[:fld] == fi }
            fa = pivot_table[:field_attrs] && pivot_table[:field_attrs][fi]
            attrs << %( compact="#{fa[:compact] ? "1" : "0"}") if fa && !fa[:compact].nil?
            attrs << %( outline="#{fa[:outline] ? "1" : "0"}") if fa && !fa[:outline].nil?
            attrs << %( subtotalTop="#{fa[:subtotal_top] ? "1" : "0"}") if fa && !fa[:subtotal_top].nil?
            attrs << %( showAll="#{fa && fa[:show_all] == true ? "1" : "0"}")
            attrs << %( numFmtId="#{fa[:num_fmt_id]}") if fa && fa[:num_fmt_id]
            attrs << %( sortType="#{xml_escape(fa[:sort_type])}") if fa && fa[:sort_type]
            attrs << ' defaultSubtotal="0"' if fa && fa[:default_subtotal] == false
            attrs << ' insertBlankRow="1"' if fa && fa[:insert_blank_row]
            attrs << ' insertPageBreak="1"' if fa && fa[:insert_page_break]
            attrs << ' includeNewItemsInFilter="1"' if fa && fa[:include_new_items_in_filter]

            field_items = pivot_table[:items] && pivot_table[:items][fi]
            if field_items
              parts << "<pivotField#{attrs}>"
              parts << %(<items count="#{field_items.size + 1}">)
              field_items.size.times { |ix| parts << %(<item x="#{ix}"/>) }
              parts << '<item t="default"/>'
              parts << "</items>"
              parts << "</pivotField>"
            else
              parts << "<pivotField#{attrs}/>"
            end
          end
          parts << "</pivotFields>"

          unless pivot_table[:row_fields].empty?
            parts << %(<rowFields count="#{pivot_table[:row_fields].size}">)
            pivot_table[:row_fields].each { |f| parts << %(<field x="#{f}"/>) }
            parts << "</rowFields>"
          end

          unless pivot_table[:col_fields].empty?
            parts << %(<colFields count="#{pivot_table[:col_fields].size}">)
            pivot_table[:col_fields].each { |f| parts << %(<field x="#{f}"/>) }
            parts << "</colFields>"
          end

          unless pivot_table[:data_fields].empty?
            parts << %(<dataFields count="#{pivot_table[:data_fields].size}">)
            pivot_table[:data_fields].each do |df|
              df_attrs = %( name="#{xml_escape(df[:name])}" fld="#{df[:fld]}" subtotal="#{df[:subtotal] || "sum"}")
              df_attrs << %( showDataAs="#{xml_escape(df[:show_data_as])}") if df[:show_data_as]
              df_attrs << %( baseField="#{df[:base_field]}") if df[:base_field]
              df_attrs << %( baseItem="#{df[:base_item]}") if df[:base_item]
              df_attrs << %( numFmtId="#{df[:num_fmt_id]}") if df[:num_fmt_id]
              parts << "<dataField#{df_attrs}/>"
            end
            parts << "</dataFields>"
          end

          if pivot_table[:pivot_table_style]
            psi = pivot_table[:pivot_table_style]
            psi_attrs = +""
            psi_attrs << %( name="#{xml_escape(psi[:name])}") if psi[:name]
            psi_attrs << %( showRowHeaders="#{psi[:show_row_headers] ? "1" : "0"}") unless psi[:show_row_headers].nil?
            psi_attrs << %( showColHeaders="#{psi[:show_col_headers] ? "1" : "0"}") unless psi[:show_col_headers].nil?
            psi_attrs << %( showRowStripes="#{psi[:show_row_stripes] ? "1" : "0"}") unless psi[:show_row_stripes].nil?
            psi_attrs << %( showColStripes="#{psi[:show_col_stripes] ? "1" : "0"}") unless psi[:show_col_stripes].nil?
            psi_attrs << %( showLastColumn="#{psi[:show_last_column] ? "1" : "0"}") unless psi[:show_last_column].nil?
            parts << "<pivotTableStyleInfo#{psi_attrs}/>"
          end

          parts << "</pivotTableDefinition>"
          parts.join
        end

        # : (untyped pivot_table, untyped _cache_id) -> untyped
        def generate_pivot_cache_definition_xml(pivot_table, _cache_id)
          pcd_attrs = %( xmlns="#{SSML_NS}" xmlns:r="#{DOC_REL_NS}" r:id="rId1" refreshOnLoad="1")
          pcd_attrs << ' saveData="0"' if pivot_table[:cache_save_data] == false
          pcd_attrs << ' enableRefresh="0"' if pivot_table[:cache_enable_refresh] == false
          pcd_attrs << %( refreshedBy="#{xml_escape(pivot_table[:cache_refreshed_by])}") if pivot_table[:cache_refreshed_by]
          pcd_attrs << %( refreshedVersion="#{pivot_table[:cache_refreshed_version]}") if pivot_table[:cache_refreshed_version]
          pcd_attrs << %( createdVersion="#{pivot_table[:cache_created_version]}") if pivot_table[:cache_created_version]
          pcd_attrs << %( recordCount="#{pivot_table[:cache_record_count]}") if pivot_table[:cache_record_count]
          pcd_attrs << ' optimizeMemory="1"' if pivot_table[:cache_optimize_memory]
          parts = [
            XML_HEADER,
            "<pivotCacheDefinition#{pcd_attrs}>"
          ]

          # Parse source ref: "Sheet1!A1:C4" => sheet name + range.
          source = pivot_table[:source_ref]
          ws_name_attr = pivot_table[:source_name] ? %( name="#{xml_escape(pivot_table[:source_name])}") : ""
          if source.include?("!")
            sname, srange = source.split("!", 2)
            sname = sname.delete("'")
            parts << %(<cacheSource type="worksheet"><worksheetSource ref="#{srange}" sheet="#{xml_escape(sname)}"#{ws_name_attr}/></cacheSource>)
          else
            parts << %(<cacheSource type="worksheet"><worksheetSource ref="#{source}"#{ws_name_attr}/></cacheSource>)
          end

          field_count = pivot_table[:field_names] ? pivot_table[:field_names].size : (pivot_table[:row_fields].size + pivot_table[:col_fields].size + pivot_table[:data_fields].size)
          parts << %(<cacheFields count="#{field_count}">)
          field_count.times do |fi|
            fname = if pivot_table[:field_names] && pivot_table[:field_names][fi]
                      pivot_table[:field_names][fi]
                    else
                      df = pivot_table[:data_fields].find { |d| d[:fld] == fi }
                      df ? df[:name] : "Field#{fi + 1}"
                    end
            fa = pivot_table[:field_attrs] && pivot_table[:field_attrs][fi]
            cf_num_fmt = (fa && fa[:cache_num_fmt_id]) || 0
            cf_attrs = %( name="#{xml_escape(fname)}" numFmtId="#{cf_num_fmt}")
            cf_attrs << %( caption="#{xml_escape(fa[:cache_caption])}") if fa && fa[:cache_caption]
            cf_attrs << %( formula="#{xml_escape(fa[:cache_formula])}") if fa && fa[:cache_formula]
            field_items = pivot_table[:items] && pivot_table[:items][fi]
            if field_items
              parts << "<cacheField#{cf_attrs}>"
              parts << %(<sharedItems count="#{field_items.size}">)
              field_items.each { |v| parts << %(<s v="#{xml_escape(v.to_s)}"/>) }
              parts << "</sharedItems>"
              parts << "</cacheField>"
            else
              parts << "<cacheField#{cf_attrs}><sharedItems/></cacheField>"
            end
          end
          parts << "</cacheFields>"
          parts << "</pivotCacheDefinition>"
          parts.join
        end

        # : (untyped pivot_table) -> untyped
        def generate_pivot_cache_records_xml(pivot_table)
          items = pivot_table[:items]
          if items&.values&.any? { |v| v && !v.empty? }
            max_len = items.values.map { |v| v ? v.size : 0 }.max
            parts = [XML_HEADER, %(<pivotCacheRecords xmlns="#{SSML_NS}" count="#{max_len}">)]
            max_len.times do |ri|
              parts << "<r>"
              field_count = pivot_table[:field_names] ? pivot_table[:field_names].size : (pivot_table[:row_fields].size + pivot_table[:col_fields].size + pivot_table[:data_fields].size)
              field_count.times do |fi|
                field_items = items[fi]
                parts << if field_items
                           %(<x v="#{ri < field_items.size ? ri : 0}"/>)
                         else
                           %(<n v="0"/>)
                         end
              end
              parts << "</r>"
            end
            parts << "</pivotCacheRecords>"
            parts.join
          else
            [XML_HEADER, %(<pivotCacheRecords xmlns="#{SSML_NS}" count="0"/>)].join
          end
        end

        # : (untyped cache_id) -> untyped
        def generate_pivot_cache_rels(cache_id)
          [
            XML_HEADER,
            %(<Relationships xmlns="#{REL_NS}">),
            %(<Relationship Id="rId1" Type="#{DOC_REL_NS}/pivotCacheRecords" Target="pivotCacheRecords#{cache_id}.xml"/>),
            "</Relationships>"
          ].join
        end

        # : (untyped cache_id) -> untyped
        def generate_pivot_table_rels(cache_id)
          [
            XML_HEADER,
            %(<Relationships xmlns="#{REL_NS}">),
            %(<Relationship Id="rId1" Type="#{DOC_REL_NS}/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition#{cache_id}.xml"/>),
            "</Relationships>"
          ].join
        end

        # : (untyped ext_link) -> untyped
        def generate_external_link_xml(ext_link)
          parts = [
            XML_HEADER,
            %(<externalLink xmlns="#{SSML_NS}" xmlns:r="#{DOC_REL_NS}">),
            '<externalBook r:id="rId1">'
          ]
          unless ext_link[:sheet_names].empty?
            parts << "<sheetNames>"
            ext_link[:sheet_names].each { |sn| parts << %(<sheetName val="#{xml_escape(sn)}"/>) }
            parts << "</sheetNames>"
          end
          parts << "</externalBook>"
          parts << "</externalLink>"
          parts.join
        end

        # : (untyped ext_link) -> untyped
        def generate_external_link_rels(ext_link)
          [
            XML_HEADER,
            %(<Relationships xmlns="#{REL_NS}">),
            %(<Relationship Id="rId1" Type="#{DOC_REL_NS}/externalLinkPath" Target="#{xml_escape(ext_link[:target])}" TargetMode="External"/>),
            "</Relationships>"
          ].join
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

        # : (untyped element, untyped opts) -> ("" | ::String)
        def build_line_end_xml(element, opts)
          return "" unless opts

          attrs = +""
          attrs << %( type="#{xml_escape(opts[:type])}") if opts[:type]
          attrs << %( w="#{xml_escape(opts[:w])}") if opts[:w]
          attrs << %( len="#{xml_escape(opts[:len])}") if opts[:len]
          "<a:#{element}#{attrs}/>"
        end

        # : (untyped value) -> untyped
        def xml_escape(value)
          value.to_s
               .gsub(/[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD\u{10000}-\u{10FFFF}]/, "")
               .gsub("&", "&amp;")
               .gsub("<", "&lt;")
               .gsub(">", "&gt;")
               .gsub('"', "&quot;")
               .gsub("'", "&apos;")
        end

        # : (untyped rich_text) -> untyped
        def rich_text_xml(rich_text)
          rich_text.runs.map do |run|
            font = run[:font]
            if font && !font.empty?
              rpr = +""
              rpr << "<b/>" if font[:bold]
              rpr << "<i/>" if font[:italic]
              rpr << "<strike/>" if font[:strike]
              if font[:underline]
                rpr << if font[:underline] == true
                         "<u/>"
                       else
                         %(<u val="#{font[:underline]}"/>)
                       end
              end
              rpr << %(<vertAlign val="#{font[:vert_align]}"/>) if font[:vert_align]
              rpr << %(<sz val="#{font[:sz]}"/>) if font[:sz]
              rpr << emit_color_xml(font)
              rpr << %(<rFont val="#{xml_escape(font[:name])}"/>) if font[:name]
              rpr << %(<family val="#{font[:family]}"/>) if font[:family]
              rpr << %(<scheme val="#{font[:scheme]}"/>) if font[:scheme]
              "<r><rPr>#{rpr}</rPr><t>#{xml_escape(run[:text])}</t></r>"
            else
              "<r><t>#{xml_escape(run[:text])}</t></r>"
            end
          end.join
        end

        # : (untyped cell_ref, untyped value, untyped style_idx, ?untyped? sst, ?ph: untyped?) -> untyped
        def cell_xml(cell_ref, value, style_idx, sst = nil, ph: nil) # rubocop:disable Naming/MethodParameterName
          s_attr = style_idx ? %( s="#{style_idx}") : ""
          ph_attr = ph ? ' ph="1"' : ""
          case value
          when Xlsxrb::Elements::CellError
            %(<c r="#{cell_ref}" t="e"#{s_attr}#{ph_attr}><v>#{xml_escape(value.code)}</v></c>)
          when Xlsxrb::Elements::Formula
            f_attrs = +""
            case value.type
            when :shared
              f_attrs << %( t="shared" si="#{value.shared_index}")
              f_attrs << %( ref="#{value.ref}") if value.ref
            when :array
              f_attrs << %( t="array" ref="#{value.ref}") if value.ref
            when :data_table
              f_attrs << ' t="dataTable"'
              f_attrs << ' dt2D="1"' if value.dt2d
              f_attrs << ' dtr="1"' if value.dtr
              f_attrs << %( r1="#{value.r1}") if value.r1
              f_attrs << %( r2="#{value.r2}") if value.r2
            end
            f_attrs << ' ca="1"' if value.calculate_always
            f_attrs << ' aca="1"' if value.aca
            f_attrs << ' bx="1"' if value.bx
            t_attr = ""
            cached_val_str = value.cached_value.to_s
            case value.cached_value
            when String
              t_attr = ' t="str"'
            when true
              t_attr = ' t="b"'
              cached_val_str = "1"
            when false
              t_attr = ' t="b"'
              cached_val_str = "0"
            end
            parts = %(<c r="#{cell_ref}"#{t_attr}#{s_attr}#{ph_attr}><f#{f_attrs}>#{xml_escape(value.expression)}</f>)
            parts << "<v>#{xml_escape(cached_val_str)}</v>" unless value.cached_value.nil?
            parts << "</c>"
            parts
          when Xlsxrb::Elements::RichText
            if sst
              rt_sst, = sst
              idx = rt_sst[value]
              %(<c r="#{cell_ref}" t="s"#{s_attr}#{ph_attr}><v>#{idx}</v></c>)
            else
              %(<c r="#{cell_ref}" t="inlineStr"#{s_attr}#{ph_attr}><is>#{rich_text_xml(value)}</is></c>)
            end
          when true, false
            %(<c r="#{cell_ref}" t="b"#{s_attr}#{ph_attr}><v>#{value ? 1 : 0}</v></c>)
          when Time
            serial = Xlsxrb::Ooxml::Utils.datetime_to_serial(value)
            dt_attr = s_attr.empty? && (dt_style = resolve_style_index(datetime_num_fmt_id)) ? %( s="#{dt_style}") : s_attr
            %(<c r="#{cell_ref}"#{dt_attr}#{ph_attr}><v>#{serial}</v></c>)
          when Date
            serial = Xlsxrb::Ooxml::Utils.date_to_serial(value)
            ds_attr = s_attr.empty? && (date_style = resolve_style_index(date_num_fmt_id)) ? %( s="#{date_style}") : s_attr
            %(<c r="#{cell_ref}"#{ds_attr}#{ph_attr}><v>#{serial}</v></c>)
          when Numeric
            %(<c r="#{cell_ref}"#{s_attr}#{ph_attr}><v>#{value}</v></c>)
          else
            if sst
              _, str_sst = sst
              idx = str_sst[value.to_s]
              %(<c r="#{cell_ref}" t="s"#{s_attr}#{ph_attr}><v>#{idx}</v></c>)
            else
              %(<c r="#{cell_ref}" t="inlineStr"#{s_attr}#{ph_attr}><is><t>#{xml_escape(value)}</t></is></c>)
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

        # : (untyped filter) -> untyped
        def emit_filter_xml(filter)
          case filter[:type]
          when :filters
            attrs = filter[:blank] ? +' blank="1"' : +""
            attrs << %( calendarType="#{filter[:calendar_type]}") if filter[:calendar_type]
            has_values = filter[:values]&.any?
            has_date_groups = filter[:date_group_items]&.any?
            if has_values || has_date_groups
              parts = ["<filters#{attrs}>"]
              filter[:values]&.each { |v| parts << %(<filter val="#{xml_escape(v)}"/>) }
              filter[:date_group_items]&.each do |dg|
                dg_attrs = %(dateTimeGrouping="#{dg[:date_time_grouping]}")
                dg_attrs << %( year="#{dg[:year]}") if dg[:year]
                dg_attrs << %( month="#{dg[:month]}") if dg[:month]
                dg_attrs << %( day="#{dg[:day]}") if dg[:day]
                dg_attrs << %( hour="#{dg[:hour]}") if dg[:hour]
                dg_attrs << %( minute="#{dg[:minute]}") if dg[:minute]
                dg_attrs << %( second="#{dg[:second]}") if dg[:second]
                parts << "<dateGroupItem #{dg_attrs}/>"
              end
              parts << "</filters>"
              parts.join
            else
              "<filters#{attrs}/>"
            end
          when :custom
            if filter[:filters]
              and_attr = filter[:and] ? ' and="1"' : ""
              parts = ["<customFilters#{and_attr}>"]
              filter[:filters].each do |cf|
                parts << %(<customFilter operator="#{cf[:operator]}" val="#{xml_escape(cf[:val])}"/>)
              end
              parts << "</customFilters>"
              parts.join
            else
              %(<customFilters><customFilter operator="#{filter[:operator]}" val="#{xml_escape(filter[:val])}"/></customFilters>)
            end
          when :dynamic
            dyn_attrs = %( type="#{filter[:dynamic_type]}")
            dyn_attrs << %( val="#{filter[:val]}") if filter[:val]
            dyn_attrs << %( valIso="#{filter[:val_iso]}") if filter[:val_iso]
            dyn_attrs << %( maxVal="#{filter[:max_val]}") if filter[:max_val]
            dyn_attrs << %( maxValIso="#{filter[:max_val_iso]}") if filter[:max_val_iso]
            "<dynamicFilter#{dyn_attrs}/>"
          when :top10
            top_attr = filter[:top] ? ' top="1"' : ""
            pct_attr = filter[:percent] ? ' percent="1"' : ""
            fv_attr = filter[:filter_val] ? %( filterVal="#{filter[:filter_val]}") : ""
            %(<top10#{top_attr}#{pct_attr} val="#{filter[:val]}"#{fv_attr}/>)
          when :color_filter
            cf_attrs = %(dxfId="#{filter[:dxf_id]}")
            cf_attrs << ' cellColor="0"' if filter[:cell_color] == false
            %(<colorFilter #{cf_attrs}/>)
          when :icon_filter
            if_attrs = %(iconSet="#{filter[:icon_set]}")
            if_attrs << %( iconId="#{filter[:icon_id]}") if filter[:icon_id]
            %(<iconFilter #{if_attrs}/>)
          else
            ""
          end
        end

        # : (untyped parts, untyped rule) -> untyped
        def emit_cf_rule(parts, rule)
          type = rule[:type]
          if type.is_a?(String) || type.is_a?(Symbol)
            t_str = type.to_s
            snake = t_str.gsub(/([A-Z]+)([A-Z][a-z])/, '\1_\2')
                         .gsub(/([a-z\d])([A-Z])/, '\1_\2')
                         .downcase.to_sym
            type = snake if CF_TYPE_MAP.key?(snake)
          end
          rule_type = CF_TYPE_MAP[type] || type.to_s
          rule_attrs = %(type="#{rule_type}")
          rule_attrs << %( priority="#{rule[:priority]}") if rule[:priority]
          rule_attrs << %( operator="#{rule[:operator]}") if rule[:operator]
          rule_attrs << %( dxfId="#{rule[:format_id]}") if rule[:format_id]
          rule_attrs << %( stopIfTrue="1") if rule[:stop_if_true]
          rule_attrs << %( aboveAverage="0") if rule[:above_average] == false
          rule_attrs << %( equalAverage="1") if rule[:equal_average]
          rule_attrs << %( rank="#{rule[:rank]}") if rule[:rank]
          rule_attrs << %( percent="1") if rule[:percent]
          rule_attrs << %( bottom="1") if rule[:bottom]
          rule_attrs << %( text="#{xml_escape(rule[:text])}") if rule[:text]
          rule_attrs << %( timePeriod="#{rule[:time_period]}") if rule[:time_period]
          rule_attrs << %( stdDev="#{rule[:std_dev]}") if rule[:std_dev]

          case type
          when :cell_is, :expression, :above_average, :top10, :duplicate_values, :unique_values,
               :contains_text, :not_contains_text, :begins_with, :ends_with,
               :contains_blanks, :not_contains_blanks, :time_period
            formulas = rule[:formulas] || [rule[:formula]].compact
            if formulas.empty?
              parts << "<cfRule #{rule_attrs}/>"
            else
              parts << "<cfRule #{rule_attrs}>"
              formulas.each { |f| parts << "<formula>#{xml_escape(f)}</formula>" }
              parts << "</cfRule>"
            end
          when :color_scale
            cs = rule[:color_scale]
            parts << "<cfRule #{rule_attrs}>"
            parts << "<colorScale>"
            cs[:cfvo]&.each do |cfvo|
              cfvo_attrs = %(type="#{cfvo[:type]}")
              cfvo_attrs << %( val="#{cfvo[:val]}") if cfvo[:val]
              cfvo_attrs << ' gte="0"' if cfvo[:gte] == false
              parts << "<cfvo #{cfvo_attrs}/>"
            end
            cs[:colors]&.each { |c| parts << emit_cf_color_xml(c) }
            parts << "</colorScale>"
            parts << "</cfRule>"
          when :data_bar
            db = rule[:data_bar]
            parts << "<cfRule #{rule_attrs}>"
            db_attrs = +""
            db_attrs << %( minLength="#{db[:min_length]}") if db[:min_length]
            db_attrs << %( maxLength="#{db[:max_length]}") if db[:max_length]
            db_attrs << %( showValue="#{db[:show_value] ? 1 : 0}") unless db[:show_value].nil?
            parts << "<dataBar#{db_attrs}>"
            db[:cfvo]&.each do |cfvo|
              cfvo_attrs = %(type="#{cfvo[:type]}")
              cfvo_attrs << %( val="#{cfvo[:val]}") if cfvo[:val]
              cfvo_attrs << ' gte="0"' if cfvo[:gte] == false
              parts << "<cfvo #{cfvo_attrs}/>"
            end
            parts << emit_cf_color_xml(db[:color]) if db[:color]
            parts << "</dataBar>"
            parts << "</cfRule>"
          when :icon_set
            is = rule[:icon_set]
            parts << "<cfRule #{rule_attrs}>"
            is_attrs = +""
            is_attrs << %( iconSet="#{is[:icon_set]}") if is[:icon_set]
            is_attrs << %( reverse="#{is[:reverse] ? 1 : 0}") unless is[:reverse].nil?
            is_attrs << %( percent="#{is[:percent] ? 1 : 0}") unless is[:percent].nil?
            is_attrs << %( showValue="#{is[:show_value] ? 1 : 0}") unless is[:show_value].nil?
            parts << "<iconSet#{is_attrs}>"
            is[:cfvo]&.each do |cfvo|
              cfvo_attrs = %(type="#{cfvo[:type]}")
              cfvo_attrs << %( val="#{cfvo[:val]}") if cfvo[:val]
              cfvo_attrs << ' gte="0"' if cfvo[:gte] == false
              parts << "<cfvo #{cfvo_attrs}/>"
            end
            parts << "</iconSet>"
            parts << "</cfRule>"
          else
            parts << "<cfRule #{rule_attrs}/>"
          end
        end

        # Emits a <color> element for CF rules, accepting either a plain RGB string or a hash.
        # : (untyped color) -> (untyped | ::String)
        def emit_cf_color_xml(color)
          if color.is_a?(Hash)
            emit_color_xml(color)
          else
            %(<color rgb="#{color}"/>)
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
          parts = [
            XML_HEADER,
            %(<cp:coreProperties xmlns:cp="#{CP_NS}" xmlns:dc="#{DC_NS}" xmlns:dcterms="#{DCTERMS_NS}" xmlns:xsi="#{XSI_NS}">)
          ]
          parts << "<dc:title>#{xml_escape(@core_properties[:title])}</dc:title>" if @core_properties[:title]
          parts << "<dc:subject>#{xml_escape(@core_properties[:subject])}</dc:subject>" if @core_properties[:subject]
          parts << "<dc:creator>#{xml_escape(@core_properties[:creator])}</dc:creator>" if @core_properties[:creator]
          parts << "<cp:keywords>#{xml_escape(@core_properties[:keywords])}</cp:keywords>" if @core_properties[:keywords]
          parts << "<dc:description>#{xml_escape(@core_properties[:description])}</dc:description>" if @core_properties[:description]
          parts << "<cp:lastModifiedBy>#{xml_escape(@core_properties[:last_modified_by])}</cp:lastModifiedBy>" if @core_properties[:last_modified_by]
          parts << "<cp:revision>#{xml_escape(@core_properties[:revision])}</cp:revision>" if @core_properties[:revision]
          parts << %(<dcterms:created xsi:type="dcterms:W3CDTF">#{xml_escape(@core_properties[:created])}</dcterms:created>) if @core_properties[:created]
          parts << %(<dcterms:modified xsi:type="dcterms:W3CDTF">#{xml_escape(@core_properties[:modified])}</dcterms:modified>) if @core_properties[:modified]
          parts << "<cp:category>#{xml_escape(@core_properties[:category])}</cp:category>" if @core_properties[:category]
          parts << "<cp:contentStatus>#{xml_escape(@core_properties[:content_status])}</cp:contentStatus>" if @core_properties[:content_status]
          parts << "<dc:language>#{xml_escape(@core_properties[:language])}</dc:language>" if @core_properties[:language]
          parts << "</cp:coreProperties>"
          parts.join
        end

        # : () -> untyped
        def generate_app_properties_xml
          parts = [
            XML_HEADER,
            %(<Properties xmlns="#{APP_NS}" xmlns:vt="#{VT_NS}">)
          ]
          parts << "<Application>#{xml_escape(@app_properties[:application])}</Application>" if @app_properties[:application]
          parts << "<AppVersion>#{xml_escape(@app_properties[:app_version])}</AppVersion>" if @app_properties[:app_version]
          if @app_properties[:heading_pairs] && @app_properties[:titles_of_parts]
            hp = @app_properties[:heading_pairs]
            tp = @app_properties[:titles_of_parts]
            parts << "<HeadingPairs>"
            parts << %(<vt:vector size="#{hp.size * 2}" baseType="variant">)
            hp.each do |label, count|
              parts << "<vt:variant><vt:lpstr>#{xml_escape(label)}</vt:lpstr></vt:variant>"
              parts << "<vt:variant><vt:i4>#{count}</vt:i4></vt:variant>"
            end
            parts << "</vt:vector>"
            parts << "</HeadingPairs>"
            parts << "<TitlesOfParts>"
            parts << %(<vt:vector size="#{tp.size}" baseType="lpstr">)
            tp.each { |t| parts << "<vt:lpstr>#{xml_escape(t)}</vt:lpstr>" }
            parts << "</vt:vector>"
            parts << "</TitlesOfParts>"
          end
          parts << "</Properties>"
          parts.join
        end

        # : () -> untyped
        def generate_custom_properties_xml
          custom_ns = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
          parts = [
            XML_HEADER,
            %(<Properties xmlns="#{custom_ns}" xmlns:vt="#{VT_NS}">)
          ]
          @custom_properties.each_with_index do |prop, idx|
            fmtid = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
            pid = idx + 2 # pids start at 2
            parts << %(<property fmtid="#{fmtid}" pid="#{pid}" name="#{xml_escape(prop[:name])}">)
            parts << case prop[:type]
                     when :number, :int, :i4
                       "<vt:i4>#{prop[:value]}</vt:i4>"
                     when :float, :r8
                       "<vt:r8>#{prop[:value]}</vt:r8>"
                     when :bool
                       "<vt:bool>#{prop[:value] ? "true" : "false"}</vt:bool>"
                     when :date, :filetime
                       "<vt:filetime>#{prop[:value]}</vt:filetime>"
                     else
                       "<vt:lpwstr>#{xml_escape(prop[:value].to_s)}</vt:lpwstr>"
                     end
            parts << "</property>"
          end
          parts << "</Properties>"
          parts.join
        end
      end
    end
  end
end
