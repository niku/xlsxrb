# frozen_string_literal: true

# rbs_inline: enabled

require "stringio"
require_relative "xml_builder"

module Xlsxrb
  module Ooxml
    # Generates worksheet XML for a list of rows.
    # Supports streaming: rows can be written one at a time.
    class WorksheetWriter
      SSML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

      COLUMN_LETTERS = (0...16_384).map do |index|
        result = +""
        i = index
        loop do
          result.prepend(("A".ord + (i % 26)).chr)
          i = (i / 26) - 1
          break if i.negative?
        end
        result.freeze
      end.freeze

      INTEGER_STRINGS = (0..65_535).map(&:to_s).freeze

      def initialize(io)
        @io = io
        @builder = XmlBuilder.new(@io)
        @row_buffer = String.new(capacity: 65_536)
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
      def write_row(row_index, cells, attrs: {}, unmapped: [], sst_index: nil)
        start unless @started

        row_num = row_index + 1
        row_num_str = row_num.to_s
        buf = @row_buffer ||= String.new(capacity: 65_536)
        buf << '<row r="' << row_num_str << '"'
        if attrs[:height]
          buf << ' ht="' << attrs[:height].to_s << '" customHeight="1"'
        elsif attrs[:custom_height]
          buf << ' customHeight="1"'
        end
        buf << ' hidden="1"' if attrs[:hidden]
        buf << ' outlineLevel="' << attrs[:outline_level].to_s << '"' if attrs[:outline_level]
        buf << ">"

        cells.each do |cell|
          if cell.is_a?(Elements::Cell)
            value = cell.value
            style_id = cell.style_index
            col_ref = cell.ref || "#{column_letter(cell.column_index)}#{row_num_str}"
            formula = cell.formula
            formula_ca = false
            cell_type_val = nil
          elsif cell.is_a?(Hash)
            value = cell[:value]
            style_id = cell[:style_index]
            col_ref = cell[:ref] || "#{column_letter(cell[:column_index])}#{row_num_str}"
            formula = cell[:formula]
            formula_ca = cell[:formula_ca]
            cell_type_val = cell[:type]
          else
            value = cell
            style_id = nil
            col_ref = nil
            formula = nil
            formula_ca = false
            cell_type_val = nil
          end

          # Fast path for common unstyled cells
          if !formula && !style_id && !cell_type_val
            case value
            when Integer, Float
              buf << '<c r="' << col_ref << '"><v>' << value.to_s << "</v></c>"
              next
            when String
              if !value.start_with?("=") && sst_index && (idx = sst_index[value])
                buf << '<c r="' << col_ref << '" t="s"><v>' << idx.to_s << "</v></c>"
                next
              end
            when true
              buf << '<c r="' << col_ref << '" t="b"><v>1</v></c>'
              next
            when false
              buf << '<c r="' << col_ref << '" t="b"><v>0</v></c>'
              next
            when nil
              buf << '<c r="' << col_ref << '"/>'
              next
            when Date
              serial = Xlsxrb::Ooxml::Utils.date_to_serial(value)
              buf << '<c r="' << col_ref << '"><v>' << serial.to_s << "</v></c>"
              next
            end
          end

          if value.nil? && formula.nil?
            buf << '<c r="' << col_ref << '"'
            buf << ' s="' << style_id.to_s << '"' if style_id
            buf << "/>"
            next
          end

          if value.is_a?(String) && value.start_with?("=") && !formula
            formula = value
            value = nil
          end

          xml_val = value
          type = cell_type_val
          formula_expr = nil

          if formula
            if formula.is_a?(Xlsxrb::Elements::Formula)
              formula_expr = formula.expression
              formula_ca = formula.calculate_always
              xml_val = formula.cached_value || value || nil
            else
              formula_expr = formula
              xml_val = value || nil
            end
            formula_expr = formula_expr[1..] if formula_expr.start_with?("=")
          end

          if !formula_expr && !type
            case value
            when String
              if sst_index && (idx = sst_index[value])
                xml_val = idx
                type = "s"
              else
                type = "inlineStr"
              end
            when Xlsxrb::Elements::RichText
              if sst_index && (idx = sst_index[value])
                xml_val = idx
                type = "s"
              else
                type = "inlineStr"
                xml_val = value
              end
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
            when BigDecimal
              xml_val = value.to_s("F")
            when Xlsxrb::Elements::CellError
              xml_val = value.code
              type = "e"
            end
          end

          buf << '<c r="' << col_ref << '"'
          buf << ' s="' << style_id.to_s << '"' if style_id
          buf << ' t="' << type << '"' if type

          if type == "inlineStr"
            if xml_val.is_a?(Xlsxrb::Elements::RichText)
              buf << "><is>"
              xml_val.runs.each do |run|
                font = run[:font]
                if font && !font.empty?
                  buf << "<r><rPr>"
                  buf << "<b/>" if font[:bold]
                  buf << "<i/>" if font[:italic]
                  buf << "<strike/>" if font[:strike]
                  if font[:underline]
                    if font[:underline] == true
                      buf << "<u/>"
                    else
                      buf << '<u val="' << font[:underline].to_s << '"/>'
                    end
                  end
                  buf << '<vertAlign val="' << font[:vert_align].to_s << '"/>' if font[:vert_align]
                  buf << '<sz val="' << font[:sz].to_s << '"/>' if font[:sz]
                  if font[:color]
                    buf << '<color rgb="' << font[:color].to_s << '"/>'
                  elsif font[:theme]
                    tint_attr = font[:tint] ? " tint=\"#{font[:tint]}\"" : ""
                    buf << '<color theme="' << font[:theme].to_s << '"' << tint_attr << "/>"
                  end
                  buf << '<rFont val="' << escape_xml(font[:name]) << '"/>' if font[:name]
                  buf << '<family val="' << font[:family].to_s << '"/>' if font[:family]
                  buf << '<scheme val="' << font[:scheme].to_s << '"/>' if font[:scheme]
                  buf << "</rPr><t>"
                else
                  buf << "<r><t>"
                end
                buf << escape_xml(run[:text]) << "</t></r>"
              end
              buf << "</is></c>"
            else
              buf << "><is><t>" << escape_xml(xml_val.to_s) << "</t></is></c>"
            end
          elsif formula_expr
            buf << if formula_ca
                     '><f ca="1">'
                   else
                     "><f>"
                   end
            buf << escape_xml(formula_expr) << "</f>"
            if xml_val
              buf << "<v>" << xml_val.to_s << "</v></c>"
            else
              buf << "</c>"
            end
          else
            buf << "><v>" << xml_val.to_s << "</v></c>"
          end
        end

        buf << "</row>"
        if buf.bytesize >= 32_768
          @io.write(buf)
          buf.clear
        end

        unmapped.each { |node| @builder.write_unmapped(node) }
      end

      # Highly optimized row writing for StreamWriter that avoids allocating intermediate Hashes.
      def write_row_values(row_index, values, styles: nil, style_map: nil, sst: nil, sst_index: nil, attrs: nil)
        start unless @started

        row_num = row_index + 1
        row_num_str = row_num < 65_536 ? INTEGER_STRINGS[row_num] : row_num.to_s
        buf = @row_buffer ||= String.new(capacity: 65_536)

        if styles.nil? && attrs.nil?
          buf << "<row r=\"#{row_num_str}\">"
          col_index = 0
          max_len = values.length
          while col_index < max_len
            value = values[col_index]
            col_ref = COLUMN_LETTERS[col_index] || column_letter(col_index)
            col_index += 1

            next if value.nil?

            case value
            when Integer
              val_str = value >= 0 && value < 65_536 ? INTEGER_STRINGS[value] : value.to_s
              buf << "<c r=\"#{col_ref}#{row_num_str}\"><v>#{val_str}</v></c>"
            when Float
              buf << "<c r=\"#{col_ref}#{row_num_str}\"><v>#{value}</v></c>"
            when String
              unless value.start_with?("=")
                raise ArgumentError, "Cell text length #{value.length} exceeds Excel limit of 32,767 characters" if @strict_excel_mode && value.length > 32_767

                idx = (sst_index[value] ||= begin
                  sst << value
                  sst.size - 1
                end)
                idx_str = idx < 65_536 ? INTEGER_STRINGS[idx] : idx.to_s
                buf << "<c r=\"#{col_ref}#{row_num_str}\" t=\"s\"><v>#{idx_str}</v></c>"
                next
              end
              formula_expr = value[1..]
              buf << "<c r=\"#{col_ref}#{row_num_str}\"><f>#{escape_xml(formula_expr)}</f></c>"
            when true
              buf << "<c r=\"#{col_ref}#{row_num_str}\" t=\"b\"><v>1</v></c>"
            when false
              buf << "<c r=\"#{col_ref}#{row_num_str}\" t=\"b\"><v>0</v></c>"
            when Date
              date_style_id = style_map ? style_map["__xlsxrb_date"] : nil
              buf << if date_style_id
                       "<c r=\"#{col_ref}#{row_num_str}\" s=\"#{date_style_id}\"><v>#{Xlsxrb::Ooxml::Utils.date_to_serial(value)}</v></c>"
                     else
                       "<c r=\"#{col_ref}#{row_num_str}\"><v>#{Xlsxrb::Ooxml::Utils.date_to_serial(value)}</v></c>"
                     end
            when Time
              time_style_id = style_map ? style_map["__xlsxrb_time"] : nil
              buf << if time_style_id
                       "<c r=\"#{col_ref}#{row_num_str}\" s=\"#{time_style_id}\"><v>#{Xlsxrb::Ooxml::Utils.datetime_to_serial(value)}</v></c>"
                     else
                       "<c r=\"#{col_ref}#{row_num_str}\"><v>#{Xlsxrb::Ooxml::Utils.datetime_to_serial(value)}</v></c>"
                     end
            when Xlsxrb::Elements::Formula
              formula_expr = value.expression
              formula_expr = formula_expr[1..] if formula_expr.start_with?("=")
              buf << if value.cached_value
                       "<c r=\"#{col_ref}#{row_num_str}\"><f>#{escape_xml(formula_expr)}</f><v>#{value.cached_value}</v></c>"
                     else
                       "<c r=\"#{col_ref}#{row_num_str}\"><f>#{escape_xml(formula_expr)}</f></c>"
                     end
            when Hash
              if value.key?(:formula)
                formula_expr = value[:formula]
                formula_expr = formula_expr[1..] if formula_expr.start_with?("=")
                xml_val = value[:value]
                buf << if xml_val
                         "<c r=\"#{col_ref}#{row_num_str}\"><f>#{escape_xml(formula_expr)}</f><v>#{xml_val}</v></c>"
                       else
                         "<c r=\"#{col_ref}#{row_num_str}\"><f>#{escape_xml(formula_expr)}</f></c>"
                       end
              end
            when Xlsxrb::Elements::CellError
              buf << "<c r=\"#{col_ref}#{row_num_str}\" t=\"e\"><v>#{value.code}</v></c>"
            when BigDecimal
              buf << "<c r=\"#{col_ref}#{row_num_str}\"><v>#{value.to_s("F")}</v></c>"
            else
              idx = (sst_index[value] ||= begin
                sst << value
                sst.size - 1
              end)
              idx_str = idx < 65_536 ? INTEGER_STRINGS[idx] : idx.to_s
              buf << "<c r=\"#{col_ref}#{row_num_str}\" t=\"s\"><v>#{idx_str}</v></c>"
            end
          end

          buf << "</row>"
          return unless buf.bytesize >= 32_768

          @io.write(buf)
          buf.clear
          return
        end

        is_styles_collection = styles && (styles.is_a?(Array) || styles.is_a?(Hash))
        single_style_id = nil
        single_style_id = style_map[styles] if styles && style_map && !is_styles_collection

        buf << '<row r="' << row_num_str << '"'
        if attrs
          buf << ' ht="' << attrs[:height].to_s << '" customHeight="1"' if attrs[:height]
          buf << ' hidden="1"' if attrs[:hidden]
          buf << ' outlineLevel="' << attrs[:outline_level].to_s << '"' if attrs[:outline_level]
        end
        buf << ">"

        max_len = values.length
        if is_styles_collection
          styles_len = styles.is_a?(Array) ? styles.length : (styles.keys.max || -1) + 1
          max_len = [max_len, styles_len].max
        end

        col_index = 0
        while col_index < max_len
          value = col_index < values.length ? values[col_index] : nil
          style_id = single_style_id
          if is_styles_collection && style_map
            style_name = if styles.is_a?(Array)
                           col_index < styles.length ? styles[col_index] : nil
                         else
                           styles[col_index]
                         end
            style_id = style_map[style_name] if style_name
          end

          col_ref = COLUMN_LETTERS[col_index] || column_letter(col_index)

          if value.nil?
            buf << '<c r="' << col_ref << row_num_str << '" s="' << style_id.to_s << '"/>' if style_id
            col_index += 1
            next
          end

          # Fast path: unstyled numbers, booleans, and simple strings (majority of cells)
          if style_id.nil? && !value.is_a?(Xlsxrb::Elements::Formula) && !value.is_a?(Hash)
            case value
            when Integer
              val_str = INTEGER_STRINGS[value] || value.to_s
              buf << '<c r="' << col_ref << row_num_str << '"><v>' << val_str << "</v></c>"
              col_index += 1
              next
            when Float
              buf << '<c r="' << col_ref << row_num_str << '"><v>' << value.to_s << "</v></c>"
              col_index += 1
              next
            when String
              unless value.start_with?("=")
                idx = sst_index[value]
                unless idx
                  sst << value
                  idx = sst.size - 1
                  sst_index[value] = idx
                end
                idx_str = INTEGER_STRINGS[idx] || idx.to_s
                buf << '<c r="' << col_ref << row_num_str << '" t="s"><v>' << idx_str << "</v></c>"
                col_index += 1
                next
              end
            when true
              buf << '<c r="' << col_ref << row_num_str << '" t="b"><v>1</v></c>'
              col_index += 1
              next
            when false
              buf << '<c r="' << col_ref << row_num_str << '" t="b"><v>0</v></c>'
              col_index += 1
              next
            when Date
              buf << '<c r="' << col_ref << row_num_str << '"><v>' << Xlsxrb::Ooxml::Utils.date_to_serial(value).to_s << "</v></c>"
              col_index += 1
              next
            when Time
              buf << '<c r="' << col_ref << row_num_str << '"><v>' << Xlsxrb::Ooxml::Utils.datetime_to_serial(value).to_s << "</v></c>"
              col_index += 1
              next
            end
          end

          # General path: styled, formula, rich text, or complex cell
          formula_expr = nil
          formula_ca = false
          xml_val = value
          type = nil

          case value
          when Xlsxrb::Elements::Formula
            formula_expr = value.expression
            formula_ca = value.calculate_always
            xml_val = value.cached_value
            case value.cached_value
            when String
              type = "str"
            when true
              type = "b"
              xml_val = "1"
            when false
              type = "b"
              xml_val = "0"
            end
          when Hash
            if value.key?(:formula)
              formula_expr = value[:formula]
              formula_ca = value[:calculate_always] || false
              xml_val = value[:value]
              case xml_val
              when String
                type = "str"
              when true
                type = "b"
                xml_val = "1"
              when false
                type = "b"
                xml_val = "0"
              end
            end
          when String
            if value.start_with?("=") && value.length > 1
              formula_expr = value
              xml_val = nil
            else
              idx = sst_index[value]
              unless idx
                sst << value
                idx = sst.size - 1
                sst_index[value] = idx
              end
              xml_val = idx
              type = "s"
            end
          when Xlsxrb::Elements::RichText
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
          when BigDecimal
            xml_val = value.to_s("F")
          when Xlsxrb::Elements::CellError
            xml_val = value.code
            type = "e"
          end

          formula_expr = formula_expr[1..] if formula_expr&.start_with?("=")

          buf << '<c r="' << col_ref << row_num_str << '"'
          buf << ' s="' << style_id.to_s << '"' if style_id
          buf << ' t="' << type << '"' if type
          if formula_expr
            buf << if formula_ca
                     '><f ca="1">'
                   else
                     "><f>"
                   end
            buf << escape_xml(formula_expr) << "</f>"
            if xml_val
              buf << "<v>" << xml_val.to_s << "</v></c>"
            else
              buf << "</c>"
            end
          else
            buf << "><v>" << xml_val.to_s << "</v></c>"
          end

          col_index += 1
        end

        buf << "</row>"
        return unless buf.bytesize >= 32_768

        @io.write(buf)
        buf.clear
      end

      # Write the worksheet footer. Call once after all rows.
      # Options for post-sheetData elements (in OOXML order):
      def finish(drawing_rid: nil, sheet_protection: nil, auto_filter: nil,
                 filter_columns: nil, sort_state: nil, merge_cells: nil,
                 conditional_formats: nil, data_validations: nil,
                 hyperlinks: nil, print_options: nil, page_margins: nil,
                 page_setup: nil, header_footer: nil, row_breaks: nil,
                 col_breaks: nil, tables: nil, table_start_rid: nil,
                 legacy_drawing_rid: nil, sparkline_groups: nil)
        return if @finished

        start unless @started
        @finished = true
        if @row_buffer && !@row_buffer.empty?
          @io.write(@row_buffer)
          @row_buffer.clear
        end
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
        write_sparklines(sparkline_groups) if sparkline_groups && !sparkline_groups.empty?
        @builder.close_tag("worksheet")
      end

      private

      # --- Pre-sheetData elements ---

      def write_sheet_properties(props)
        attrs = {}
        has_children = props[:tab_color] || !props[:fit_to_page].nil? || !props[:outline_below].nil? || !props[:outline_right].nil?
        if has_children
          @builder.open_tag("sheetPr", attrs)
          @builder.empty_tag("tabColor", { rgb: props[:tab_color] }) if props[:tab_color]

          outline_attrs = {}
          outline_attrs[:summaryBelow] = props[:outline_below] ? "1" : "0" unless props[:outline_below].nil?
          outline_attrs[:summaryRight] = props[:outline_right] ? "1" : "0" unless props[:outline_right].nil?
          @builder.empty_tag("outlinePr", outline_attrs) unless outline_attrs.empty?

          unless props[:fit_to_page].nil?
            @builder.empty_tag("pageSetUpPr", { fitToPage: props[:fit_to_page] ? "1" : "0" })
          end

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
            if type.is_a?(String) || type.is_a?(Symbol)
              t_str = type.to_s
              snake = t_str.gsub(/([A-Z]+)([A-Z][a-z])/, '\1_\2')
                           .gsub(/([a-z\d])([A-Z])/, '\1_\2')
                           .downcase.to_sym
              type = snake if %i[cell_is expression color_scale data_bar icon_set above_average top10 duplicate_values unique_values contains_text not_contains_text begins_with ends_with contains_blanks not_contains_blanks time_period].include?(snake)
            end

            cf_type = case type
                      when :cell_is then "cellIs"
                      when :expression then "expression"
                      when :color_scale then "colorScale"
                      when :data_bar then "dataBar"
                      when :icon_set then "iconSet"
                      when :above_average then "aboveAverage"
                      when :top10 then "top10"
                      when :duplicate_values then "duplicateValues"
                      when :unique_values then "uniqueValues"
                      when :contains_text then "containsText"
                      when :not_contains_text then "notContainsText"
                      when :begins_with then "beginsWith"
                      when :ends_with then "endsWith"
                      when :contains_blanks then "containsBlanks"
                      when :not_contains_blanks then "notContainsBlanks"
                      when :time_period then "timePeriod"
                      else type.to_s
                      end

            r_attrs = { type: cf_type, priority: (rule[:priority] || (idx + 1)).to_s }
            r_attrs[:operator] = rule[:operator].to_s if rule[:operator]
            r_attrs[:dxfId] = rule[:format_id].to_s if rule[:format_id]
            r_attrs[:stopIfTrue] = "1" if rule[:stop_if_true]
            r_attrs[:aboveAverage] = "0" if rule[:above_average] == false
            r_attrs[:equalAverage] = "1" if rule[:equal_average]
            r_attrs[:rank] = rule[:rank].to_s if rule[:rank]
            r_attrs[:percent] = "1" if rule[:percent]
            r_attrs[:bottom] = "1" if rule[:bottom]
            r_attrs[:text] = rule[:text].to_s if rule[:text]
            r_attrs[:timePeriod] = rule[:time_period].to_s if rule[:time_period]
            r_attrs[:stdDev] = rule[:std_dev].to_s if rule[:std_dev]

            case type
            when :color_scale
              @builder.open_tag("cfRule", r_attrs)
              cs = rule[:color_scale] || {}
              cfvos = cs[:cfvo] || [
                { type: "min" },
                { type: "percent", val: "50" },
                { type: "max" }
              ]
              colors = cs[:colors] || [
                "FFF8696B", # Red
                "FFFFEB84", # Yellow
                "FF63BE7B"  # Green
              ]
              @builder.tag("colorScale") do |b|
                cfvos.each do |cfvo|
                  cfvo_attrs = { type: cfvo[:type] }
                  cfvo_attrs[:val] = cfvo[:val].to_s if cfvo[:val]
                  cfvo_attrs[:gte] = "0" if cfvo[:gte] == false
                  b.empty_tag("cfvo", cfvo_attrs)
                end
                colors.each do |c|
                  if c.is_a?(Hash)
                    c_attrs = {}
                    c_attrs[:auto] = "1" if c[:auto]
                    c_attrs[:indexed] = c[:indexed].to_s if c[:indexed]
                    c_attrs[:rgb] = c[:rgb] if c[:rgb]
                    c_attrs[:theme] = c[:theme].to_s if c[:theme]
                    c_attrs[:tint] = c[:tint].to_s if c[:tint]
                    b.empty_tag("color", c_attrs)
                  else
                    b.empty_tag("color", { rgb: c.to_s })
                  end
                end
              end
              @builder.close_tag("cfRule")
            when :data_bar
              @builder.open_tag("cfRule", r_attrs)
              db = rule[:data_bar] || {}
              db_color = db[:color] || rule[:color] || "FF5A8DD4"
              cfvos = db[:cfvo] || [
                { type: "min" },
                { type: "max" }
              ]
              db_attrs = {}
              db_attrs[:minLength] = db[:min_length].to_s if db[:min_length]
              db_attrs[:maxLength] = db[:max_length].to_s if db[:max_length]
              db_attrs[:showValue] = db[:show_value] ? "1" : "0" unless db[:show_value].nil?

              @builder.tag("dataBar", db_attrs) do |b|
                cfvos.each do |cfvo|
                  cfvo_attrs = { type: cfvo[:type] }
                  cfvo_attrs[:val] = cfvo[:val].to_s if cfvo[:val]
                  cfvo_attrs[:gte] = "0" if cfvo[:gte] == false
                  b.empty_tag("cfvo", cfvo_attrs)
                end
                if db_color.is_a?(Hash)
                  c_attrs = {}
                  c_attrs[:rgb] = db_color[:rgb] if db_color[:rgb]
                  c_attrs[:theme] = db_color[:theme].to_s if db_color[:theme]
                  b.empty_tag("color", c_attrs)
                else
                  b.empty_tag("color", { rgb: db_color.to_s })
                end
              end
              @builder.close_tag("cfRule")
            when :icon_set
              @builder.open_tag("cfRule", r_attrs)
              is = rule[:icon_set] || {}
              icon_style = is[:icon_set] || rule[:icon_style] || "3Arrows"
              cfvos = is[:cfvo] || [
                { type: "percent", val: "0" },
                { type: "percent", val: "33" },
                { type: "percent", val: "67" }
              ]
              is_attrs = { iconSet: icon_style }
              is_attrs[:reverse] = is[:reverse] ? "1" : "0" unless is[:reverse].nil?
              is_attrs[:percent] = is[:percent] ? "1" : "0" unless is[:percent].nil?
              is_attrs[:showValue] = is[:show_value] ? "1" : "0" unless is[:show_value].nil?

              @builder.tag("iconSet", is_attrs) do |b|
                cfvos.each do |cfvo|
                  cfvo_attrs = { type: cfvo[:type] }
                  cfvo_attrs[:val] = cfvo[:val].to_s if cfvo[:val]
                  cfvo_attrs[:gte] = "0" if cfvo[:gte] == false
                  b.empty_tag("cfvo", cfvo_attrs)
                end
              end
              @builder.close_tag("cfRule")
            else
              formulas = rule[:formulas] || [rule[:formula]].compact
              if formulas.empty?
                @builder.empty_tag("cfRule", r_attrs)
              else
                @builder.open_tag("cfRule", r_attrs)
                formulas.each do |f|
                  @builder.tag("formula") { |b| b.text(f) }
                end
                @builder.close_tag("cfRule")
              end
            end
          end
          @builder.close_tag("conditionalFormatting")
        end
      end

      def write_data_validations(validations)
        @builder.open_tag("dataValidations", { count: validations.size.to_s })
        validations.each do |dv|
          if dv[:in].is_a?(Array)
            dv = dv.dup
            dv[:type] = "list"
            dv[:formula1] = %("#{dv[:in].join(",")}")
          end
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

      # --- Sparklines (extLst) ---

      def write_sparklines(groups)
        @builder.open_tag("extLst")
        @builder.open_tag("ext", {
                            uri: "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}",
                            "xmlns:x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
                          })
        @builder.open_tag("x14:sparklineGroups", {
                            "xmlns:xm": "http://schemas.microsoft.com/office/excel/2006/main"
                          })
        groups.each do |group|
          sg_attrs = {}
          sg_attrs[:type] = group[:type] if group[:type]
          sg_attrs[:displayEmptyCellsAs] = group[:display_empty] if group[:display_empty]
          sg_attrs[:markers] = "1" if group[:markers]
          sg_attrs[:high] = "1" if group[:high]
          sg_attrs[:low] = "1" if group[:low]
          sg_attrs[:first] = "1" if group[:first]
          sg_attrs[:last] = "1" if group[:last]
          sg_attrs[:negative] = "1" if group[:negative]
          sg_attrs[:manualMax] = group[:max].to_s if group[:max]
          sg_attrs[:manualMin] = group[:min].to_s if group[:min]
          sg_attrs[:lineWeight] = group[:line_weight].to_s if group[:line_weight]

          @builder.open_tag("x14:sparklineGroup", sg_attrs)

          write_sparkline_color("x14:colorSeries", group[:color_series]) if group[:color_series]
          write_sparkline_color("x14:colorNegative", group[:color_negative]) if group[:color_negative]
          write_sparkline_color("x14:colorAxis", group[:color_axis]) if group[:color_axis]
          write_sparkline_color("x14:colorMarkers", group[:color_markers]) if group[:color_markers]
          write_sparkline_color("x14:colorFirst", group[:color_first]) if group[:color_first]
          write_sparkline_color("x14:colorLast", group[:color_last]) if group[:color_last]
          write_sparkline_color("x14:colorHigh", group[:color_high]) if group[:color_high]
          write_sparkline_color("x14:colorLow", group[:color_low]) if group[:color_low]

          @builder.open_tag("x14:sparklines")
          sparklines = group[:sparklines] || []
          sparklines.each do |sp|
            @builder.open_tag("x14:sparkline")
            @builder.open_tag("xm:f")
            @builder.text(sp[:data_ref])
            @builder.close_tag("xm:f")
            @builder.open_tag("xm:sqref")
            @builder.text(sp[:location_ref])
            @builder.close_tag("xm:sqref")
            @builder.close_tag("x14:sparkline")
          end
          @builder.close_tag("x14:sparklines")
          @builder.close_tag("x14:sparklineGroup")
        end
        @builder.close_tag("x14:sparklineGroups")
        @builder.close_tag("ext")
        @builder.close_tag("extLst")
      end

      def write_sparkline_color(tag, color)
        attrs = {}
        if color.is_a?(String) || color.is_a?(Symbol)
          resolved = Xlsxrb::StyleBuilder.new.resolve_color(color)
          attrs[:rgb] = resolved
        else
          attrs[:rgb] = color[:rgb] if color[:rgb]
          attrs[:theme] = color[:theme].to_s if color[:theme]
          attrs[:tint] = color[:tint].to_s if color[:tint]
        end
        @builder.empty_tag(tag, attrs)
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
          attrs[:outlineLevel] = col[:outline_level].to_s if col[:outline_level]
          @builder.empty_tag("col", attrs)
        end
        @builder.close_tag("cols")
      end

      def write_cell(cell)
        ref = cell[:ref] || cell_ref(cell[:row_index], cell[:column_index])
        attrs = { r: ref }

        value = cell[:value]
        formula = cell[:formula]
        # For formula cells with cached values, don't infer type from value
        type = formula ? cell[:type] : (cell[:type] || cell_type(value))
        attrs[:t] = type if type
        attrs[:s] = cell[:style_index].to_s if cell[:style_index]

        if value.nil? && formula.nil?
          @builder.empty_tag("c", attrs)
          return
        end

        @builder.open_tag("c", attrs)
        if formula
          f_attrs = {}
          f_attrs[:ca] = "1" if cell[:formula_ca]
          if f_attrs.empty?
            @builder.tag("f") { |b| b.text(formula) }
          else
            @builder.tag("f", f_attrs) { |b| b.text(formula) }
          end
        end
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

      INVALID_XML_CHARS_RE = /[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD\u{10000}-\u{10FFFF}]/
      ESCAPE_RE = /[&<>"']/
      ESCAPE_MAP = { "&" => "&amp;", "<" => "&lt;", ">" => "&gt;", '"' => "&quot;", "'" => "&apos;" }.freeze
      private_constant :INVALID_XML_CHARS_RE, :ESCAPE_RE, :ESCAPE_MAP

      def escape_xml(value)
        str = value.to_s
        str = str.gsub(INVALID_XML_CHARS_RE, "") if str.match?(INVALID_XML_CHARS_RE)
        str.match?(ESCAPE_RE) ? str.gsub(ESCAPE_RE, ESCAPE_MAP) : str
      end
    end
  end
end
