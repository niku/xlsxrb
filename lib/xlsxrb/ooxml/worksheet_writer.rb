# frozen_string_literal: true

require "stringio"
require_relative "xml_builder"

module Xlsxrb
  module Ooxml
    # Generates worksheet XML for a list of rows.
    # Supports streaming: rows can be written one at a time.
    class WorksheetWriter
      SSML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

      def initialize(io)
        @io = io
        @builder = XmlBuilder.new(@io)
        @started = false
        @finished = false
      end

      # Write the worksheet header. Call once before writing rows.
      # Options for pre-sheetData elements:
      #   sheet_properties: Hash of sheet-level properties (:tab_color, etc.)
      #   freeze_pane: { row:, col:, state: :frozen }
      #   split_pane: { x_split:, y_split:, top_left_cell: }
      #   selection: { active_cell:, sqref:, pane: }
      #   sheet_view: Hash of sheet view properties
      def start(columns: [], sheet_properties: nil, freeze_pane: nil, split_pane: nil, selection: nil, sheet_view: nil)
        return if @started

        @started = true
        @builder.declaration
        @builder.open_tag("worksheet", { xmlns: SSML_NS, "xmlns:r": DOC_REL_NS })

        write_sheet_properties(sheet_properties) if sheet_properties && !sheet_properties.empty?
        write_sheet_views(freeze_pane: freeze_pane, split_pane: split_pane, selection: selection, sheet_view: sheet_view) if freeze_pane || split_pane || selection || (sheet_view && !sheet_view.empty?)
        write_columns(columns) unless columns.empty?

        @builder.open_tag("sheetData")
      end

      # Write a single row. Automatically calls start if needed.
      def write_row(row_index, cells, attrs: {}, unmapped: [])
        start unless @started

        row_attrs = { r: (row_index + 1).to_s } # convert 0-based to 1-based
        row_attrs[:ht] = attrs[:height].to_s if attrs[:height]
        row_attrs[:hidden] = "1" if attrs[:hidden]
        row_attrs[:customHeight] = "1" if attrs[:custom_height]

        @builder.open_tag("row", row_attrs)

        cells.each do |cell|
          write_cell(cell)
        end

        unmapped.each { |node| @builder.write_unmapped(node) }

        @builder.close_tag("row")
      end

      # Highly optimized row writing for StreamWriter that avoids allocating intermediate Hashes.
      def write_row_values(row_index, values, styles: nil, style_map: nil, sst: nil, sst_index: nil, attrs: nil)
        start unless @started

        row_num = row_index + 1
        row_num_str = row_num.to_s
        style_lookup_enabled = styles && style_map && (styles.is_a?(Array) || styles.is_a?(Hash))
        io = @io
        io.write("<row r=\"")
        io.write(row_num_str)
        io.write('"')
        if attrs
          if attrs[:height]
            io.write(' ht="')
            io.write(attrs[:height].to_s)
            io.write('" customHeight="1"')
          end
          io.write(' hidden="1"') if attrs[:hidden]
        end
        io.write(">")

        col_index = 0
        while col_index < values.length
          value = values[col_index]
          style_id = nil
          if style_lookup_enabled
            style_name = styles[col_index]
            style_id = style_map[style_name] if style_name
          end

          col_ref = column_letter(col_index)

          if value.nil?
            if style_id
              io.write('<c r="')
              io.write(col_ref)
              io.write(row_num_str)
              io.write('" s="')
              io.write(style_id.to_s)
              io.write('"/>')
            end
            col_index += 1
            next
          end

          xml_val = value
          type = nil

          case value
          when String
            idx = sst_index[value]
            unless idx
              sst << value
              idx = sst.size - 1
              sst_index[value] = idx
            end
            xml_val = idx
            type = "s"
          when true
            xml_val = "1"
            type = "b"
          when false
            xml_val = "0"
            type = "b"
          when Date
            xml_val = Xlsxrb::Ooxml::Utils.date_to_serial(value)
          when Time
            xml_val = Xlsxrb::Ooxml::Utils.datetime_to_serial(value)
          end

          io.write('<c r="')
          io.write(col_ref)
          io.write(row_num_str)
          io.write('"')
          if style_id
            io.write(' s="')
            io.write(style_id.to_s)
            io.write('"')
          end
          if type
            io.write(' t="')
            io.write(type)
            io.write('"')
          end
          io.write("><v>")
          io.write(xml_val.to_s)
          io.write("</v></c>")

          col_index += 1
        end

        io.write("</row>")
      end

      # Write the worksheet footer. Call once after all rows.
      # Options for post-sheetData elements (in OOXML order):
      def finish(drawing_rid: nil, sheet_protection: nil, auto_filter: nil,
                 filter_columns: nil, sort_state: nil, merge_cells: nil,
                 conditional_formats: nil, data_validations: nil,
                 hyperlinks: nil, print_options: nil, page_margins: nil,
                 page_setup: nil, header_footer: nil, row_breaks: nil,
                 col_breaks: nil, tables: nil, table_start_rid: nil,
                 legacy_drawing_rid: nil)
        return if @finished

        start unless @started
        @finished = true
        @builder.close_tag("sheetData")

        # Elements must appear in OOXML specification order after sheetData
        write_sheet_protection(sheet_protection) if sheet_protection
        write_auto_filter(auto_filter, filter_columns, sort_state) if auto_filter
        write_merge_cells(merge_cells) if merge_cells && !merge_cells.empty?
        write_conditional_formatting(conditional_formats) if conditional_formats && !conditional_formats.empty?
        write_data_validations(data_validations) if data_validations && !data_validations.empty?
        write_hyperlinks(hyperlinks) if hyperlinks && !hyperlinks.empty?
        write_print_options(print_options) if print_options && !print_options.empty?
        write_page_margins(page_margins) if page_margins
        write_page_setup(page_setup) if page_setup && !page_setup.empty?
        write_header_footer(header_footer) if header_footer && !header_footer.empty?
        write_row_breaks(row_breaks) if row_breaks && !row_breaks.empty?
        write_col_breaks(col_breaks) if col_breaks && !col_breaks.empty?
        @builder.empty_tag("drawing", { "r:id": drawing_rid }) if drawing_rid
        @builder.empty_tag("legacyDrawing", { "r:id": legacy_drawing_rid }) if legacy_drawing_rid
        write_table_parts(tables, table_start_rid) if tables && !tables.empty?
        @builder.close_tag("worksheet")
      end

      private

      # --- Pre-sheetData elements ---

      def write_sheet_properties(props)
        attrs = {}
        has_children = props[:tab_color]
        if has_children
          @builder.open_tag("sheetPr", attrs)
          @builder.empty_tag("tabColor", { rgb: props[:tab_color] }) if props[:tab_color]
          @builder.close_tag("sheetPr")
        else
          @builder.empty_tag("sheetPr", attrs) unless attrs.empty?
        end
      end

      def write_sheet_views(freeze_pane: nil, split_pane: nil, selection: nil, sheet_view: nil)
        @builder.open_tag("sheetViews")
        sv_attrs = { tabSelected: "1", workbookViewId: "0" }
        if sheet_view
          sv_attrs[:showGridLines] = "0" if sheet_view[:show_grid_lines] == false
          sv_attrs[:showRowColHeaders] = "0" if sheet_view[:show_row_col_headers] == false
          sv_attrs[:rightToLeft] = "1" if sheet_view[:right_to_left]
          sv_attrs[:zoomScale] = sheet_view[:zoom_scale].to_s if sheet_view[:zoom_scale]
        end
        @builder.open_tag("sheetView", sv_attrs)
        if freeze_pane
          pane_attrs = {}
          pane_attrs[:xSplit] = freeze_pane[:col].to_s if freeze_pane[:col]&.positive?
          pane_attrs[:ySplit] = freeze_pane[:row].to_s if freeze_pane[:row]&.positive?
          top_left_col = column_letter(freeze_pane[:col] || 0)
          top_left_row = (freeze_pane[:row] || 0) + 1
          pane_attrs[:topLeftCell] = "#{top_left_col}#{top_left_row}"
          pane_attrs[:state] = "frozen"
          # Determine active pane
          pane_attrs[:activePane] = if (freeze_pane[:col] || 0).positive? && (freeze_pane[:row] || 0).positive?
                                      "bottomRight"
                                    elsif (freeze_pane[:col] || 0).positive?
                                      "topRight"
                                    else
                                      "bottomLeft"
                                    end
          @builder.empty_tag("pane", pane_attrs)
        elsif split_pane
          pane_attrs = {}
          pane_attrs[:xSplit] = split_pane[:x_split].to_s if split_pane[:x_split]&.positive?
          pane_attrs[:ySplit] = split_pane[:y_split].to_s if split_pane[:y_split]&.positive?
          pane_attrs[:topLeftCell] = split_pane[:top_left_cell] if split_pane[:top_left_cell]
          @builder.empty_tag("pane", pane_attrs)
        end
        if selection
          sel_attrs = {}
          sel_attrs[:activeCell] = selection[:active_cell] if selection[:active_cell]
          sel_attrs[:sqref] = selection[:sqref] || selection[:active_cell] if selection[:active_cell]
          sel_attrs[:pane] = selection[:pane] if selection[:pane]
          @builder.empty_tag("selection", sel_attrs)
        end
        @builder.close_tag("sheetView")
        @builder.close_tag("sheetViews")
      end

      # --- Post-sheetData elements (in OOXML order) ---

      def write_sheet_protection(opts)
        attrs = {}
        attrs[:sheet] = "1" if opts[:sheet] != false
        attrs[:objects] = "1" if opts[:objects]
        attrs[:scenarios] = "1" if opts[:scenarios]
        attrs[:formatCells] = "0" if opts[:format_cells] == false
        attrs[:formatColumns] = "0" if opts[:format_columns] == false
        attrs[:formatRows] = "0" if opts[:format_rows] == false
        attrs[:insertColumns] = "0" if opts[:insert_columns] == false
        attrs[:insertRows] = "0" if opts[:insert_rows] == false
        attrs[:insertHyperlinks] = "0" if opts[:insert_hyperlinks] == false
        attrs[:deleteColumns] = "0" if opts[:delete_columns] == false
        attrs[:deleteRows] = "0" if opts[:delete_rows] == false
        attrs[:selectLockedCells] = "1" if opts[:select_locked_cells]
        attrs[:sort] = "0" if opts[:sort] == false
        attrs[:autoFilter] = "0" if opts[:auto_filter] == false
        attrs[:pivotTables] = "0" if opts[:pivot_tables] == false
        attrs[:selectUnlockedCells] = "1" if opts[:select_unlocked_cells]
        attrs[:password] = opts[:password] if opts[:password]
        attrs[:algorithmName] = opts[:algorithm_name] if opts[:algorithm_name]
        attrs[:hashValue] = opts[:hash_value] if opts[:hash_value]
        attrs[:saltValue] = opts[:salt_value] if opts[:salt_value]
        attrs[:spinCount] = opts[:spin_count].to_s if opts[:spin_count]
        @builder.empty_tag("sheetProtection", attrs)
      end

      def write_auto_filter(range, filter_columns, sort_state)
        if (filter_columns && !filter_columns.empty?) || sort_state
          @builder.open_tag("autoFilter", { ref: range })
          filter_columns&.each do |col_id, filter|
            write_filter_column(col_id, filter)
          end
          if sort_state
            ss_attrs = { ref: sort_state[:ref] }
            ss_attrs[:columnSort] = "1" if sort_state[:column_sort]
            ss_attrs[:caseSensitive] = "1" if sort_state[:case_sensitive]
            @builder.open_tag("sortState", ss_attrs)
            (sort_state[:sort_conditions] || []).each do |sc|
              sc_attrs = { ref: sc[:ref] }
              sc_attrs[:descending] = "1" if sc[:descending]
              @builder.empty_tag("sortCondition", sc_attrs)
            end
            @builder.close_tag("sortState")
          end
          @builder.close_tag("autoFilter")
        else
          @builder.empty_tag("autoFilter", { ref: range })
        end
      end

      def write_filter_column(col_id, filter)
        @builder.open_tag("filterColumn", { colId: col_id.to_s })
        case filter[:type]
        when :filters
          f_attrs = {}
          f_attrs[:blank] = "1" if filter[:blank]
          @builder.open_tag("filters", f_attrs)
          (filter[:values] || []).each do |val|
            @builder.empty_tag("filter", { val: val.to_s })
          end
          @builder.close_tag("filters")
        when :custom
          if filter[:filters]
            c_attrs = {}
            c_attrs[:and] = "1" if filter[:and]
            @builder.open_tag("customFilters", c_attrs)
            filter[:filters].each do |cf|
              @builder.empty_tag("customFilter", { operator: cf[:operator], val: cf[:val].to_s })
            end
          else
            @builder.open_tag("customFilters")
            @builder.empty_tag("customFilter", { operator: filter[:operator], val: filter[:val].to_s })
          end
          @builder.close_tag("customFilters")
        when :dynamic
          @builder.empty_tag("dynamicFilter", { type: filter[:dynamic_type] })
        when :top10
          t_attrs = {}
          t_attrs[:top] = filter[:top] ? "1" : "0" unless filter[:top].nil?
          t_attrs[:percent] = filter[:percent] ? "1" : "0" unless filter[:percent].nil?
          t_attrs[:val] = filter[:val].to_s if filter[:val]
          @builder.empty_tag("top10", t_attrs)
        end
        @builder.close_tag("filterColumn")
      end

      def write_merge_cells(ranges)
        @builder.open_tag("mergeCells", { count: ranges.size.to_s })
        ranges.each do |range|
          @builder.empty_tag("mergeCell", { ref: range })
        end
        @builder.close_tag("mergeCells")
      end

      def write_conditional_formatting(rules)
        # Group rules by sqref
        grouped = {}
        rules.each do |rule|
          sqref = rule[:sqref]
          grouped[sqref] ||= []
          grouped[sqref] << rule
        end
        grouped.each do |sqref, sqref_rules|
          @builder.open_tag("conditionalFormatting", { sqref: sqref })
          sqref_rules.each_with_index do |rule, idx|
            type = rule[:type]
            cf_type = case type
                      when :cell_is then "cellIs"
                      when :expression then "expression"
                      when :color_scale then "colorScale"
                      when :data_bar then "dataBar"
                      when :icon_set then "iconSet"
                      else type.to_s
                      end
            r_attrs = { type: cf_type, priority: (rule[:priority] || (idx + 1)).to_s }
            r_attrs[:operator] = rule[:operator].to_s if rule[:operator]
            r_attrs[:dxfId] = rule[:format_id].to_s if rule[:format_id]
            @builder.open_tag("cfRule", r_attrs)
            @builder.tag("formula") { |b| b.text(rule[:formula]) } if rule[:formula]
            (rule[:formulas] || []).each do |f|
              @builder.tag("formula") { |b| b.text(f) }
            end
            @builder.close_tag("cfRule")
          end
          @builder.close_tag("conditionalFormatting")
        end
      end

      def write_data_validations(validations)
        @builder.open_tag("dataValidations", { count: validations.size.to_s })
        validations.each do |dv|
          dv_attrs = { sqref: dv[:sqref] }
          dv_type = dv[:type]
          dv_attrs[:type] = dv_type.to_s if dv_type
          dv_attrs[:operator] = dv[:operator].to_s if dv[:operator]
          dv_attrs[:allowBlank] = "1" if dv[:allow_blank]
          dv_attrs[:showInputMessage] = "1" if dv[:show_input_message]
          dv_attrs[:showErrorMessage] = "1" if dv[:show_error_message]
          dv_attrs[:errorStyle] = dv[:error_style].to_s if dv[:error_style]
          dv_attrs[:errorTitle] = dv[:error_title] if dv[:error_title]
          dv_attrs[:error] = dv[:error] if dv[:error]
          dv_attrs[:promptTitle] = dv[:prompt_title] if dv[:prompt_title]
          dv_attrs[:prompt] = dv[:prompt] if dv[:prompt]

          has_formulas = dv[:formula1] || dv[:formula2]
          if has_formulas
            @builder.open_tag("dataValidation", dv_attrs)
            @builder.tag("formula1") { |b| b.text(dv[:formula1].to_s) } if dv[:formula1]
            @builder.tag("formula2") { |b| b.text(dv[:formula2].to_s) } if dv[:formula2]
            @builder.close_tag("dataValidation")
          else
            @builder.empty_tag("dataValidation", dv_attrs)
          end
        end
        @builder.close_tag("dataValidations")
      end

      def write_hyperlinks(links)
        @builder.open_tag("hyperlinks")
        links.each_with_index do |link, idx|
          h_attrs = { ref: link[:cell] }
          h_attrs[:"r:id"] = "rId#{link[:_rid] || (idx + 1)}" if link[:url]
          h_attrs[:location] = link[:location] if link[:location]
          h_attrs[:display] = link[:display] if link[:display]
          h_attrs[:tooltip] = link[:tooltip] if link[:tooltip]
          @builder.empty_tag("hyperlink", h_attrs)
        end
        @builder.close_tag("hyperlinks")
      end

      def write_print_options(opts)
        attrs = {}
        attrs[:gridLines] = "1" if opts[:grid_lines]
        attrs[:headings] = "1" if opts[:headings]
        attrs[:horizontalCentered] = "1" if opts[:horizontal_centered]
        attrs[:verticalCentered] = "1" if opts[:vertical_centered]
        @builder.empty_tag("printOptions", attrs) unless attrs.empty?
      end

      def write_page_margins(margins)
        attrs = {
          left: (margins[:left] || 0.7).to_s,
          right: (margins[:right] || 0.7).to_s,
          top: (margins[:top] || 0.75).to_s,
          bottom: (margins[:bottom] || 0.75).to_s,
          header: (margins[:header] || 0.3).to_s,
          footer: (margins[:footer] || 0.3).to_s
        }
        @builder.empty_tag("pageMargins", attrs)
      end

      def write_page_setup(opts)
        attrs = {}
        attrs[:orientation] = opts[:orientation].to_s if opts[:orientation]
        attrs[:paperSize] = opts[:paper_size].to_s if opts[:paper_size]
        attrs[:scale] = opts[:scale].to_s if opts[:scale]
        attrs[:fitToWidth] = opts[:fit_to_width].to_s if opts[:fit_to_width]
        attrs[:fitToHeight] = opts[:fit_to_height].to_s if opts[:fit_to_height]
        attrs[:firstPageNumber] = opts[:first_page_number].to_s if opts[:first_page_number]
        attrs[:pageOrder] = opts[:page_order].to_s if opts[:page_order]
        attrs[:blackAndWhite] = "1" if opts[:black_and_white]
        attrs[:draft] = "1" if opts[:draft]
        @builder.empty_tag("pageSetup", attrs) unless attrs.empty?
      end

      def write_header_footer(opts)
        @builder.open_tag("headerFooter")
        @builder.tag("oddHeader") { |b| b.text(opts[:odd_header]) } if opts[:odd_header]
        @builder.tag("oddFooter") { |b| b.text(opts[:odd_footer]) } if opts[:odd_footer]
        @builder.tag("evenHeader") { |b| b.text(opts[:even_header]) } if opts[:even_header]
        @builder.tag("evenFooter") { |b| b.text(opts[:even_footer]) } if opts[:even_footer]
        @builder.close_tag("headerFooter")
      end

      def write_row_breaks(breaks)
        @builder.open_tag("rowBreaks", { count: breaks.size.to_s, manualBreakCount: breaks.size.to_s })
        breaks.each do |brk|
          @builder.empty_tag("brk", { id: brk.to_s, max: "16383", man: "1" })
        end
        @builder.close_tag("rowBreaks")
      end

      def write_col_breaks(breaks)
        @builder.open_tag("colBreaks", { count: breaks.size.to_s, manualBreakCount: breaks.size.to_s })
        breaks.each do |brk|
          @builder.empty_tag("brk", { id: brk.to_s, max: "1048575", man: "1" })
        end
        @builder.close_tag("colBreaks")
      end

      def write_table_parts(tables, start_rid)
        @builder.open_tag("tableParts", { count: tables.size.to_s })
        tables.each_with_index do |_tbl, idx|
          rid = start_rid ? "rId#{start_rid + idx}" : "rId#{idx + 1}"
          @builder.empty_tag("tablePart", { "r:id": rid })
        end
        @builder.close_tag("tableParts")
      end

      # --- Columns ---

      def write_columns(columns)
        @builder.open_tag("cols")
        columns.each do |col|
          attrs = {
            min: ((col[:index] || col[:min] || 0) + 1).to_s,
            max: ((col[:index] || col[:max] || col[:min] || 0) + 1).to_s
          }
          attrs[:width] = col[:width].to_s if col[:width]
          attrs[:hidden] = "1" if col[:hidden]
          attrs[:customWidth] = "1" if col[:custom_width] || col[:width]
          @builder.empty_tag("col", attrs)
        end
        @builder.close_tag("cols")
      end

      def write_cell(cell)
        ref = cell[:ref] || cell_ref(cell[:row_index], cell[:column_index])
        attrs = { r: ref }

        value = cell[:value]
        type = cell[:type] || cell_type(value)
        attrs[:t] = type if type
        attrs[:s] = cell[:style_index].to_s if cell[:style_index]

        formula = cell[:formula]

        if value.nil? && formula.nil?
          @builder.empty_tag("c", attrs)
          return
        end

        @builder.open_tag("c", attrs)
        @builder.tag("f") { |b| b.text(formula) } if formula
        @builder.tag("v") { |b| b.text(xml_cell_value(value, type)) } unless value.nil?
        @builder.close_tag("c")
      end

      def cell_type(value)
        case value
        when String then "s" # will be shared string index
        when true, false then "b"
        end
      end

      def xml_cell_value(value, _type)
        case value
        when true then "1"
        when false then "0"
        else value.to_s
        end
      end

      def cell_ref(row_index, col_index)
        col_letter = column_letter(col_index)
        "#{col_letter}#{row_index + 1}"
      end

      def column_letter(index)
        Xlsxrb::Elements::Cell.column_letter(index)
      end
    end
  end
end
