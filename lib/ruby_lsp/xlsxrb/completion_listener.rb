# frozen_string_literal: true

module RubyLsp
  module Xlsxrb
    # CompletionListener provides intelligent, context-aware method completions
    # for all xlsxrb public block variables (StreamWriter, WorkbookBuilder, WorksheetProxy,
    # StreamSheet, Elements::Workbook, Elements::Worksheet, Elements::Row, Elements::Cell,
    # ChartBuilder, StyleBuilder).
    #
    # Note: This is an interim enhancer for Ruby LSP until native RBS block argument
    # type propagation is fully supported upstream in Ruby LSP / language server tools.
    class CompletionListener
      STREAM_WRITER_ITEMS = [
        {
          label: "sheet",
          detail: "(name = nil, **opts) { |sheet| ... } -> void",
          documentation: "Adds a new worksheet to the streaming XLSX writer and yields a `WorksheetProxy`.\n\n```ruby\nwb.sheet(\"Sales\") do |s|\n  s.row([\"Product\", \"Price\"], styles: :bold)\nend\n```"
        },
        {
          label: "style",
          detail: "(name, **options) -> Symbol",
          documentation: "Registers a reusable named cell style (font, fill, border, alignment, number_format).\n\n```ruby\nwb.style(:bold_blue, font: { bold: true, color: \"0000FF\" })\n```"
        },
        {
          label: "workbook_property",
          detail: "(name, value) -> void",
          documentation: "Sets a workbook-level property (e.g. `:update_links`).\n\n```ruby\nwb.workbook_property(:update_links, \"never\")\n```"
        },
        {
          label: "protect_workbook",
          detail: "(**opts) -> void",
          documentation: "Configures workbook protection options (e.g. `structure: true`, `windows: true`)."
        },
        {
          label: "defined_name",
          detail: "(name, value, sheet: nil, hidden: false) -> void",
          documentation: "Adds a workbook-level defined named range or formula expression."
        },
        {
          label: "print_area",
          detail: "(range, sheet: nil) -> void",
          documentation: "Sets the print area range for a sheet (e.g. `\"A1:F50\"`)."
        },
        {
          label: "print_titles",
          detail: "(rows: nil, cols: nil, sheet: nil) -> void",
          documentation: "Sets repeating title rows or columns for printing (e.g. `rows: \"1:2\"`)."
        },
        {
          label: "properties",
          detail: "(core: nil, app: nil) -> void",
          documentation: "Sets multiple core and application metadata properties at once."
        },
        {
          label: "core_property",
          detail: "(name, value) -> void",
          documentation: "Sets core document metadata (e.g. `:title`, `:creator`, `:subject`, `:keywords`)."
        },
        {
          label: "app_property",
          detail: "(name, value) -> void",
          documentation: "Sets application document metadata (e.g. `:company`, `:manager`)."
        },
        {
          label: "custom_property",
          detail: "(name, value, type: :string) -> void",
          documentation: "Adds a custom document property (type: `:string`, `:number`, `:bool`, `:date`)."
        },
        {
          label: "close",
          detail: "() -> void",
          documentation: "Finalizes the streaming writer and writes the completed XLSX archive."
        }
      ].freeze

      WORKBOOK_BUILDER_ITEMS = [
        {
          label: "sheet",
          detail: "(name = nil, **opts) { |sheet_builder| ... } -> WorksheetBuilder",
          documentation: "Adds a new worksheet to the workbook builder and yields a `WorksheetBuilder`.\n\n```ruby\nwb.sheet(\"Summary\") do |s|\n  s.row([\"Item\", \"Total\"])\nend\n```"
        },
        {
          label: "style",
          detail: "(name, **opts) { |style_builder| ... } -> Symbol",
          documentation: "Defines a named style in the workbook builder."
        },
        {
          label: "build",
          detail: "() -> Elements::Workbook",
          documentation: "Builds and returns the immutable `Elements::Workbook` DOM hierarchy."
        },
        {
          label: "workbook_property",
          detail: "(name, value) -> void",
          documentation: "Sets a workbook property."
        },
        {
          label: "protect_workbook",
          detail: "(**opts) -> void",
          documentation: "Sets workbook protection options."
        },
        {
          label: "defined_name",
          detail: "(name, value, sheet: nil, hidden: false) -> void",
          documentation: "Adds a defined name or formula to the workbook."
        },
        {
          label: "print_area",
          detail: "(range, sheet: nil) -> void",
          documentation: "Sets the print area for a sheet."
        },
        {
          label: "print_titles",
          detail: "(rows: nil, cols: nil, sheet: nil) -> void",
          documentation: "Sets repeating print titles."
        },
        {
          label: "properties",
          detail: "(core: nil, app: nil) -> void",
          documentation: "Sets multiple metadata properties."
        },
        {
          label: "core_property",
          detail: "(name, value) -> void",
          documentation: "Sets core document metadata."
        },
        {
          label: "app_property",
          detail: "(name, value) -> void",
          documentation: "Sets application document metadata."
        },
        {
          label: "custom_property",
          detail: "(name, value, type: :string) -> void",
          documentation: "Adds a custom document metadata property."
        }
      ].freeze

      WORKSHEET_PROXY_ITEMS = [
        {
          label: "row",
          detail: "(values, styles: nil, height: nil, hidden: false, outline_level: nil) -> void",
          documentation: "Writes a single row of cells with optional style name, height, and outline level.\n\n```ruby\ns.row([\"Product\", \"Price\"], styles: :bold)\ns.row([\"Laptop\", 1200])\n```"
        },
        {
          label: "column",
          detail: "(col_index, width: nil, hidden: false, outline_level: nil, custom_width: false) -> void",
          documentation: "Configures column formatting (width in characters, visibility, grouping)."
        },
        {
          label: "merge",
          detail: "(range = nil, row: nil, col_start: nil, col_end: nil) -> void",
          documentation: "Merges a range of cells.\n\n```ruby\ns.merge(\"A1:C1\")\ns.merge(row: 0, col_start: 0, col_end: 2)\n```"
        },
        {
          label: "freeze_pane",
          detail: "(row: 0, col: 0) -> void",
          documentation: "Freezes rows and columns from scrolling (0-based indices).\n\n```ruby\ns.freeze_pane(row: 1, col: 0) # Freeze top header row\n```"
        },
        {
          label: "auto_filter",
          detail: "(range) -> void",
          documentation: "Enables Excel auto-filter dropdown arrows over a cell range (e.g. `\"A1:D100\"`)."
        },
        {
          label: "hyperlink",
          detail: "(cell, url = nil, display: nil, tooltip: nil, location: nil) -> void",
          documentation: "Adds a clickable hyperlink to a cell reference.\n\n```ruby\ns.hyperlink(\"A1\", \"https://example.com\", display: \"Example\")\n```"
        },
        {
          label: "chart",
          detail: "(type = nil, **opts) { |chart| ... } -> void",
          documentation: "Adds a chart (bar, column, line, pie, scatter, area, doughnut, radar) to the worksheet."
        },
        {
          label: "table",
          detail: "(range, name: \"Table1\", **opts) -> void",
          documentation: "Adds a formatted Excel Table to the range."
        },
        {
          label: "conditional_format",
          detail: "(range, type: \"cellIs\", **opts) -> void",
          documentation: "Adds conditional formatting rules (cellIs, colorScale, dataBar, expression)."
        },
        {
          label: "validate_data",
          detail: "(range, type: \"list\", **opts) -> void",
          documentation: "Adds data validation rules (dropdown list, whole number, decimal, date range, custom formula)."
        },
        {
          label: "image",
          detail: "(file_data, ext: \"png\", from_col: 0, from_row: 0, to_col: 5, to_row: 10, **opts) -> void",
          documentation: "Inserts an image into the worksheet."
        },
        {
          label: "comment",
          detail: "(cell, text, author: nil) -> void",
          documentation: "Adds an author comment / note to a cell."
        },
        {
          label: "sparkline_group",
          detail: "(range:, type: \"line\", sparklines:, **opts) -> void",
          documentation: "Adds in-cell sparklines (line, column, win_loss)."
        },
        {
          label: "pivot_table",
          detail: "(range, name:, row_fields:, data_fields:, **opts) -> void",
          documentation: "Creates a native Excel Pivot Table from worksheet data."
        },
        {
          label: "shape",
          detail: "(preset: \"rect\", text: nil, from_col: 0, from_row: 0, to_col: 5, to_row: 5, **opts) -> void",
          documentation: "Adds a drawing shape to the sheet."
        },
        {
          label: "page_setup",
          detail: "(orientation: \"portrait\", paper_size: 9, **opts) -> void",
          documentation: "Configures page orientation (portrait/landscape), paper size (e.g. 9 for A4), and scaling."
        },
        {
          label: "page_margins",
          detail: "(left: nil, right: nil, top: nil, bottom: nil, header: nil, footer: nil) -> void",
          documentation: "Configures print page margins in inches."
        },
        {
          label: "header_footer",
          detail: "(**opts) -> void",
          documentation: "Configures print header and footer text."
        },
        {
          label: "print_options",
          detail: "(name, value) -> void",
          documentation: "Configures print options (e.g. `s.print_options(:grid_lines, true)`)."
        },
        {
          label: "protect_sheet",
          detail: "(**opts) -> void",
          documentation: "Protects worksheet with optional password hashing and capability flags."
        },
        {
          label: "sheet_view",
          detail: "(name, value) -> void",
          documentation: "Sets view settings (e.g. `s.sheet_view(:show_grid_lines, false)`)."
        },
        {
          label: "sheet_properties",
          detail: "(name, value) -> void",
          documentation: "Sets sheet properties (e.g. `s.sheet_properties(:tab_color, \"FF0000\")`)."
        },
        {
          label: "select_cell",
          detail: "(active_cell, sqref: nil, pane: nil) -> void",
          documentation: "Sets the active / selected cell when opening the worksheet in Excel."
        },
        {
          label: "split_pane",
          detail: "(x_split: 0, y_split: 0, top_left_cell: nil) -> void",
          documentation: "Splits worksheet view into independent scrollable panes."
        },
        {
          label: "filter_column",
          detail: "(col_id, filter_values) -> void",
          documentation: "Sets filter criteria for a column within the auto-filter range."
        },
        {
          label: "sort_state",
          detail: "(ref, sort_conditions, **opts) -> void",
          documentation: "Configures sort state conditions on a cell range."
        },
        {
          label: "page_break_row",
          detail: "(row_num) -> void",
          documentation: "Adds a horizontal page break before the given row index."
        },
        {
          label: "page_break_col",
          detail: "(col_index) -> void",
          documentation: "Adds a vertical page break before the given column index."
        }
      ].freeze

      STREAM_SHEET_ITEMS = [
        {
          label: "each_row",
          detail: "() { |row| ... } -> void",
          documentation: "Yields each row in the streaming sheet as an `Elements::Row`.\n\n```ruby\nsheet.each_row do |row|\n  puts row.to_a.inspect\nend\n```"
        },
        {
          label: "each",
          detail: "() { |row| ... } -> void",
          documentation: "Yields each row in the streaming sheet."
        },
        {
          label: "name",
          detail: "() -> String",
          documentation: "Returns the worksheet name."
        }
      ].freeze

      WORKBOOK_ELEMENT_ITEMS = [
        {
          label: "sheet",
          detail: "(index_or_name) -> Elements::Worksheet?",
          documentation: "Returns the worksheet by 0-based index or name.\n\n```ruby\nsheet = workbook.sheet(0)\n```"
        },
        {
          label: "[]",
          detail: "(index_or_name) -> Elements::Worksheet?",
          documentation: "Returns the worksheet by 0-based index or name.\n\n```ruby\nsheet = workbook[\"Sheet1\"]\n```"
        },
        {
          label: "update_sheet",
          detail: "(name_or_index, sheet) -> Elements::Workbook",
          documentation: "Replaces or updates a sheet and returns a new `Elements::Workbook`."
        },
        {
          label: "sheet_names",
          detail: "() -> Array<String>",
          documentation: "Returns the list of all sheet names in the workbook."
        },
        {
          label: "each",
          detail: "() { |sheet| ... } -> void",
          documentation: "Yields each `Elements::Worksheet` in the workbook.\n\n```ruby\nworkbook.each do |sheet|\n  puts sheet.name\nend\n```"
        },
        {
          label: "save",
          detail: "(target_path_or_io) -> void",
          documentation: "Writes the workbook DOM to a target XLSX file or IO stream."
        },
        {
          label: "valid?",
          detail: "() -> Boolean",
          documentation: "Returns true if all workbook elements and sheets are valid."
        },
        {
          label: "validate",
          detail: "() -> Array<String>",
          documentation: "Validates workbook structure and returns any validation errors."
        }
      ].freeze

      WORKSHEET_ELEMENT_ITEMS = [
        {
          label: "[]",
          detail: "(cell_ref) -> Elements::Cell?",
          documentation: "Finds and returns a cell by A1 reference.\n\n```ruby\ncell = sheet[\"B2\"]\n```"
        },
        {
          label: "cell_value",
          detail: "(cell_ref) -> Object?",
          documentation: "Returns the raw value of the cell at the given A1 reference."
        },
        {
          label: "each_row",
          detail: "() { |row| ... } -> void",
          documentation: "Yields each row in the worksheet as an `Elements::Row`.\n\n```ruby\nsheet.each_row do |row|\n  puts row.to_a.inspect\nend\n```"
        },
        {
          label: "each",
          detail: "() { |row| ... } -> void",
          documentation: "Yields each row in the worksheet as an `Elements::Row`."
        },
        {
          label: "each_cell",
          detail: "() { |cell| ... } -> void",
          documentation: "Yields every `Elements::Cell` across all rows."
        },
        {
          label: "update_cell",
          detail: "(ref, new_value) -> Elements::Worksheet",
          documentation: "Returns a new worksheet with the updated cell value."
        },
        {
          label: "row_at",
          detail: "(row_index) -> Elements::Row?",
          documentation: "Returns the row at the 0-based row index."
        },
        {
          label: "first_row",
          detail: "() -> Elements::Row?",
          documentation: "Returns the first row in the sheet."
        },
        {
          label: "last_row",
          detail: "() -> Elements::Row?",
          documentation: "Returns the last row in the sheet."
        },
        {
          label: "cells",
          detail: "() -> Array<Elements::Cell>",
          documentation: "Returns a flat array of all cells in the sheet."
        },
        {
          label: "cells_hash",
          detail: "() -> Hash<String, Elements::Cell>",
          documentation: "Returns a hash mapping A1 references to `Elements::Cell` objects."
        },
        {
          label: "name",
          detail: "() -> String",
          documentation: "Returns the worksheet name."
        },
        {
          label: "valid?",
          detail: "() -> Boolean",
          documentation: "Returns true if all cells and rows in the worksheet are valid."
        },
        {
          label: "validate",
          detail: "() -> Array<String>",
          documentation: "Returns any validation errors found in the worksheet."
        }
      ].freeze

      ROW_ELEMENT_ITEMS = [
        {
          label: "[]",
          detail: "(col_index_or_symbol) -> Elements::Cell? | Object?",
          documentation: "Returns the cell at 0-based column index, or access property symbol (`:cells`, `:index`, `:height`).\n\n```ruby\ncell = row[0]\n```"
        },
        {
          label: "to_a",
          detail: "() -> Array<Object?>",
          documentation: "Returns an array containing the values of all cells in this row.\n\n```ruby\nvalues = row.to_a\n```"
        },
        {
          label: "values",
          detail: "() -> Array<Object?>",
          documentation: "Returns an array of cell values in this row."
        },
        {
          label: "cell_at",
          detail: "(col_index) -> Elements::Cell?",
          documentation: "Returns the cell at the 0-based column index."
        },
        {
          label: "each",
          detail: "() { |cell| ... } -> void",
          documentation: "Yields each `Elements::Cell` in this row."
        },
        {
          label: "each_cell",
          detail: "() { |cell| ... } -> void",
          documentation: "Yields each `Elements::Cell` in this row."
        },
        {
          label: "cells",
          detail: "() -> Array<Elements::Cell>",
          documentation: "Returns the list of cells in this row."
        },
        {
          label: "index",
          detail: "() -> Integer",
          documentation: "Returns the 0-based row index."
        },
        {
          label: "height",
          detail: "() -> Float?",
          documentation: "Returns custom row height in points, or nil."
        },
        {
          label: "valid?",
          detail: "() -> Boolean",
          documentation: "Returns true if all cells in this row are valid."
        },
        {
          label: "validate",
          detail: "() -> Array<String>",
          documentation: "Validates all cells in this row and returns error messages."
        }
      ].freeze

      CELL_ELEMENT_ITEMS = [
        {
          label: "value",
          detail: "() -> Object?",
          documentation: "Returns the cell's typed value (String, Integer, Float, Date, Time, nil)."
        },
        {
          label: "ref",
          detail: "() -> String",
          documentation: "Returns the cell's A1 reference (e.g. `\"B2\"`)."
        },
        {
          label: "to_s",
          detail: "() -> String",
          documentation: "Returns string representation of cell value."
        },
        {
          label: "to_i",
          detail: "() -> Integer",
          documentation: "Converts cell value to Integer."
        },
        {
          label: "to_f",
          detail: "() -> Float",
          documentation: "Converts cell value to Float."
        },
        {
          label: "to_date",
          detail: "() -> Date?",
          documentation: "Converts cell value to Date."
        },
        {
          label: "to_time",
          detail: "() -> Time?",
          documentation: "Converts cell value to Time."
        },
        {
          label: "column_letter",
          detail: "() -> String",
          documentation: "Returns the column letters (e.g. `\"A\"`, `\"AA\"`)."
        },
        {
          label: "row_index",
          detail: "() -> Integer",
          documentation: "Returns the 0-based row index."
        },
        {
          label: "column_index",
          detail: "() -> Integer",
          documentation: "Returns the 0-based column index."
        },
        {
          label: "style_index",
          detail: "() -> Integer?",
          documentation: "Returns the 0-based index into stylesheet formatting."
        },
        {
          label: "content",
          detail: "() -> String",
          documentation: "Returns string content of cell value."
        },
        {
          label: "[]",
          detail: "(symbol) -> Object?",
          documentation: "Hash-style access to cell attributes (`:value`, `:ref`, `:style_index`)."
        },
        {
          label: "valid?",
          detail: "() -> Boolean",
          documentation: "Returns true if row_index, column_index, and value are valid."
        },
        {
          label: "validate",
          detail: "() -> Array<String>",
          documentation: "Returns validation errors if indices are negative or invalid."
        }
      ].freeze

      CHART_BUILDER_ITEMS = [
        {
          label: "title",
          detail: "(text) -> void",
          documentation: "Sets the chart title text."
        },
        {
          label: "series",
          detail: "(range, name: nil) -> void",
          documentation: "Adds data series value range."
        },
        {
          label: "categories",
          detail: "(range) -> void",
          documentation: "Sets category label range (e.g. `'Sheet1!$A$2:$A$10'`)."
        },
        {
          label: "legend",
          detail: "(position: 'r') -> void",
          documentation: "Configures legend position (`'r'`, `'l'`, `'t'`, `'b'`, `'none'`)."
        },
        {
          label: "plot_by",
          detail: "(grouping) -> void",
          documentation: "Sets series grouping (`'standard'`, `'stacked'`, `'percentStacked'`)."
        },
        {
          label: "dimension",
          detail: "(from_col: 0, from_row: 0, to_col: 8, to_row: 15) -> void",
          documentation: "Sets chart position and size on the worksheet."
        },
        {
          label: "gridlines",
          detail: "(horizontal: true, vertical: false) -> void",
          documentation: "Configures chart major gridlines."
        },
        {
          label: "style",
          detail: "(id) -> void",
          documentation: "Sets built-in chart style number (1-48)."
        }
      ].freeze

      STYLE_BUILDER_ITEMS = [
        {
          label: "font",
          detail: "(name: 'Calibri', size: 11, bold: false, italic: false, underline: false, color: nil) -> void",
          documentation: "Sets font typography, size, weight, and color."
        },
        {
          label: "fill",
          detail: "(pattern: 'solid', fg_color: nil, bg_color: nil) -> void",
          documentation: "Sets background fill pattern and foreground color."
        },
        {
          label: "border",
          detail: "(left: nil, right: nil, top: nil, bottom: nil) -> void",
          documentation: "Sets cell border lines and colors."
        },
        {
          label: "alignment",
          detail: "(horizontal: nil, vertical: nil, wrap_text: false, rotation: 0) -> void",
          documentation: "Sets cell text alignment and text wrap."
        },
        {
          label: "number_format",
          detail: "(format_code_or_id) -> void",
          documentation: "Sets custom number or date format code (e.g. `'#,##0.00'`, `'yyyy-mm-dd'`)."
        }
      ].freeze

      CLASS_NAMES = {
        stream_writer: "Xlsxrb::StreamWriter",
        workbook_builder: "Xlsxrb::WorkbookBuilder",
        worksheet_proxy: "Xlsxrb::StreamWriter::WorksheetProxy",
        stream_sheet: "Xlsxrb::StreamSheet",
        workbook: "Xlsxrb::Elements::Workbook",
        worksheet: "Xlsxrb::Elements::Worksheet",
        row: "Xlsxrb::Elements::Row",
        cell: "Xlsxrb::Elements::Cell",
        chart_builder: "Xlsxrb::ChartBuilder",
        style_builder: "Xlsxrb::StyleBuilder"
      }.freeze

      ITEMS_BY_TYPE = {
        stream_writer: STREAM_WRITER_ITEMS,
        workbook_builder: WORKBOOK_BUILDER_ITEMS,
        worksheet_proxy: WORKSHEET_PROXY_ITEMS,
        stream_sheet: STREAM_SHEET_ITEMS,
        workbook: WORKBOOK_ELEMENT_ITEMS,
        worksheet: WORKSHEET_ELEMENT_ITEMS,
        row: ROW_ELEMENT_ITEMS,
        cell: CELL_ELEMENT_ITEMS,
        chart_builder: CHART_BUILDER_ITEMS,
        style_builder: STYLE_BUILDER_ITEMS
      }.freeze

      def initialize(response_builder, node_context, dispatcher, global_state)
        @response_builder = response_builder
        @node_context = node_context
        @global_state = global_state

        dispatcher.register(self, :on_call_node_enter)
      end

      def on_call_node_enter(node)
        receiver = node.receiver
        receiver_name = receiver.name if receiver.is_a?(Prism::LocalVariableReadNode) || (receiver.is_a?(Prism::CallNode) && receiver.receiver.nil?)
        return unless receiver_name

        target_type = infer_target_type(receiver_name)
        return unless target_type

        items = ITEMS_BY_TYPE[target_type]
        return unless items

        detail_prefix = "(#{CLASS_NAMES[target_type]}) "
        items.each_with_index do |item, index|
          sort_key = "#{index.to_s.rjust(3, "0")}_#{item[:label]}"
          @response_builder << ::RubyLsp::Interface::CompletionItem.new(
            label: item[:label],
            kind: ::RubyLsp::Constant::CompletionItemKind::METHOD,
            detail: "#{detail_prefix}#{item[:detail]}",
            sort_text: sort_key,
            documentation: {
              kind: ::RubyLsp::Constant::MarkupKind::MARKDOWN,
              value: item[:documentation]
            }
          )
        end
      end

      private

      def infer_target_type(receiver_name)
        parent = @node_context&.parent

        # 1. AST Context Analysis: Inspect enclosing block caller
        if parent.is_a?(Prism::CallNode)
          method_name = parent.name
          caller_receiver = parent.receiver

          if caller_receiver_is_xlsxrb?(caller_receiver)
            case method_name
            when :write then return :stream_writer
            when :read then return :stream_sheet
            when :build then return :workbook_builder
            when :modify then return :workbook
            end
          end

          case method_name
          when :sheet then return :worksheet_proxy
          when :chart then return :chart_builder
          when :style then return :style_builder
          when :each_row then return :row
          when :each_cell then return :cell
          when :each
            return infer_each_block_target(caller_receiver)
          end
        end

        # 2. Variable Name Heuristics (fallback if parent node wasn't direct)
        infer_from_variable_name(receiver_name)
      end

      def infer_each_block_target(caller_receiver)
        if receiver_matches?(caller_receiver, %i[workbook wb])
          :worksheet
        elsif receiver_matches?(caller_receiver, %i[row r stream_row])
          :cell
        else
          :row
        end
      end

      def infer_from_variable_name(name)
        case name
        when :wb, :workbook
          :workbook_builder
        when :s, :sheet, :ws, :worksheet
          :worksheet_proxy
        when :stream_sheet, :ss
          :stream_sheet
        when :stream_writer, :sw
          :stream_writer
        when :r, :row
          :row
        when :c, :cell
          :cell
        when :chart, :chart_builder, :cb
          :chart_builder
        when :style, :style_builder, :sb
          :style_builder
        end
      end

      def caller_receiver_is_xlsxrb?(receiver)
        return false unless receiver

        if receiver.is_a?(Prism::ConstantReadNode) || receiver.is_a?(Prism::ConstantPathNode)
          receiver.slice == "Xlsxrb"
        else
          false
        end
      end

      def receiver_matches?(receiver, names)
        return false unless receiver

        rec_name = receiver.name if receiver.is_a?(Prism::LocalVariableReadNode) || (receiver.is_a?(Prism::CallNode) && receiver.receiver.nil?)
        names.include?(rec_name)
      end
    end
  end
end
