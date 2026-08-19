# frozen_string_literal: true

# rbs_inline: enabled

require "tempfile"
require_relative "style_builder"
require_relative "chart_builder"
require_relative "ooxml/writer"
require_relative "ooxml/worksheet_writer"
require_relative "ooxml/workbook_writer"

module Xlsxrb
  # High-performance streaming writer that outputs XLSX files with O(1) constant memory.
  #
  # @example Streaming write using StreamWriter
  #   Xlsxrb.write("output.xlsx") do |writer|
  #     writer.sheet("Sales") do |s|
  #       s.row(["Item", "Price"])
  #       s.row(["Coffee", 3.50])
  #     end
  #   end
  #
  # @api public
  class StreamWriter
    # @return [String, nil] Name of the currently active worksheet.
    #: String?
    attr_reader :current_sheet

    # Initializes a streaming writer context.
    #
    # @param target [String, IO, StringIO] Destination file path or writable IO stream.
    # @param strict_excel_mode [Boolean] Whether to enforce Microsoft Excel specification limits.
    #: (untyped target, ?strict_excel_mode: bool) -> void
    def initialize(target, strict_excel_mode: true)
      @target = target
      @strict_excel_mode = strict_excel_mode
      @sst = []
      @sst_index = {}
      @sheets = []
      @current_sheet = nil
      @current_row_index = 0
      @tempfiles = []
      @current_tempfile = nil
      @current_row_writer = nil
      @current_columns = []
      @current_charts = []
      @current_hyperlinks = []
      @current_auto_filter = nil
      @current_filter_columns = {}
      @current_sort_state = nil
      @current_data_validations = []
      @current_conditional_formats = []
      @current_tables = []
      @current_comments = []
      @current_merge_cells = []
      @current_freeze_pane = nil
      @current_split_pane = nil
      @current_selection = nil
      @current_page_margins = nil
      @current_page_setup = {}
      @current_header_footer = {}
      @current_print_options = {}
      @current_sheet_protection = nil
      @current_images = []
      @current_shapes = []
      @current_sheet_properties = {}
      @current_sheet_view = {}
      @current_row_breaks = []
      @current_col_breaks = []
      @styles = {} # { style_name => StyleBuilder }
      @style_writer = Ooxml::Writer.new
      @style_name_to_id = {}
      # Workbook-level settings
      @defined_names = []
      @core_properties = {}
      @app_properties = {}
      @custom_properties = []
      @workbook_protection = nil
      @workbook_properties = { update_links: "never" }
    end

    # Sets a workbook property.
    #
    # @note **SECURITY WARNING:** If you set `:update_links` to anything other than `"never"`,
    #   you may expose end-users to malicious external reference vulnerabilities (e.g., CSV/DDE Injection)
    #   when they open the generated Excel file. Ensure you fully trust the exported data.
    #
    # @param name [Symbol] Property name (e.g. :update_links).
    # @param value [String, Integer, Boolean] Property value.
    # @return [void]
    #: (Symbol name, String | Integer | bool value) -> void
    def workbook_property(name, value)
      @workbook_properties[name] = value
    end

    # Defines or configures a named cell style.
    #
    # @param name [String, Symbol] The name of the style.
    # @param opts [Hash] Style options (e.g. bold: true, fill_color: "FF0000").
    # @yield [style_builder]
    # @yieldparam style_builder [Xlsxrb::StyleBuilder]
    # @return [StyleBuilder]
    #: (String | Symbol name, **untyped opts) ?{ (StyleBuilder) -> void } -> StyleBuilder
    def style(name, **opts)
      style_name = name.to_s
      style_builder = StyleBuilder.new(style_name)
      style_builder.apply_options!(**opts) unless opts.empty?
      yield style_builder if block_given?
      @styles[style_name] = style_builder

      # Register immediately with low-level style writer
      @style_name_to_id[style_name] = style_builder.register_with(@style_writer)

      style_builder
    end

    # Proxy object yielded by the `sheet` method to prevent writing to inactive sheets in streaming mode.
    #
    # @api public
    class WorksheetProxy
      # @param writer [StreamWriter]
      # @param sheet_name [String]
      def initialize(writer, sheet_name)
        @writer = writer
        @sheet_name = sheet_name
      end

      # Defines or configures a named cell style on the active sheet.
      #
      # @example
      #   s.style(:header, bold: true, fill_color: "4F81BD", font_color: "FFFFFF")
      #
      # @param name [String, Symbol] The name of the style.
      # @param opts [Hash] Style options (e.g. bold: true, fill_color: "FF0000").
      # @yield [style_builder]
      # @yieldparam style_builder [Xlsxrb::StyleBuilder]
      # @return [Xlsxrb::StyleBuilder]
      # @api public
      #: (String | Symbol name, **untyped opts) ?{ (Xlsxrb::StyleBuilder) -> void } -> Xlsxrb::StyleBuilder
      def style(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.style(...)
      end

      # Merges a range of cells into a single cell.
      #
      # @example Merge with cell reference string
      #   s.merge("A1:C1")
      #
      # @example Merge with coordinates
      #   s.merge(row: 0, col_start: 0, col_end: 2)
      #
      # @param range [String, nil] The cell range (e.g. "A1:B2").
      # @param row [Integer, nil] 0-based row index.
      # @param col_start [Integer, String, nil] 0-based start column index or letter.
      # @param col_end [Integer, String, nil] 0-based end column index or letter.
      # @param row_start [Integer, nil] 0-based start row index.
      # @param row_end [Integer, nil] 0-based end row index.
      # @return [void]
      # @api public
      #: (?String? range, ?row: Integer | nil, ?col_start: (Integer | String)?, ?col_end: (Integer | String)?, ?row_start: Integer | nil, ?row_end: Integer | nil) -> void
      def merge(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.merge(...)
      end

      # Adds a drawing shape to the sheet.
      #
      # @example
      #   s.shape(preset: "ellipse", text: "Circle", from_col: 1, from_row: 1, to_col: 4, to_row: 5)
      #
      # @param preset [String] Preset shape type (e.g. "rect", "ellipse").
      # @param text [String, nil] Shape label text.
      # @param from_col [Integer] Starting column index (0-based).
      # @param from_row [Integer] Starting row index (0-based).
      # @param to_col [Integer] Ending column index (0-based).
      # @param to_row [Integer] Ending row index (0-based).
      # @param opts [Hash] Additional shape formatting options.
      # @return [void]
      # @api public
      #: (?preset: String, ?text: String?, ?from_col: Integer, ?from_row: Integer, ?to_col: Integer, ?to_row: Integer, **untyped opts) -> void
      def shape(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.shape(...)
      end

      # simplecov:disable
      # Edge case / untested delegation block
      #: (*untyped args, **untyped kwargs) ?{ (*untyped) -> untyped } -> untyped
      def internal_sheet_setup(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.internal_sheet_setup(...)
      end
      # simplecov:enable

      # Appends a row of cells to the active sheet.
      #
      # @example Write an array of values
      #   s.row(["Name", "Age", "City"])
      #
      # @example Write with explicit column keys and styles
      #   s.row({ A: "Header", C: 100 }, styles: { A: :bold })
      #
      # @param values [Array, Hash] The cell values (e.g. `[1, 2, 3]` or `{ A: 1, C: 3 }`).
      # @param styles [String, Symbol, Array, Hash, nil] Style names or hashes to apply.
      # @param height [Float, Integer, nil] The row height in points (0 - 409).
      # @param hidden [Boolean] Whether the row is hidden.
      # @param custom_height [Boolean] Whether to flag as custom height.
      # @param outline_level [Integer, nil] Grouping/outline level (0 - 7).
      # @return [void]
      # @api public
      #: (Array[untyped] | Hash[untyped, untyped] values, ?styles: untyped, ?height: Float | Integer | nil, ?hidden: bool, ?custom_height: bool, ?outline_level: Integer | nil) -> void
      def row(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.row(...)
      end

      # Configures column width and properties.
      #
      # @example Set column A width
      #   s.column(0, width: 25.0)
      #
      # @param col_index [Integer, String, Symbol] 0-based column index or letter (e.g. 0 or "A" or :A).
      # @param width [Float, Integer, nil] Column width in characters.
      # @param hidden [Boolean] Whether the column is hidden.
      # @param best_fit [Boolean] Whether the column automatically fits content.
      # @param custom_width [Boolean] Whether to flag as custom width.
      # @param outline_level [Integer, nil] Grouping/outline level (0 - 7).
      # @param collapsed [Boolean] Whether the outline group is collapsed.
      # @return [void]
      # @api public
      #: (Integer | String | Symbol | Range[Integer | String] | Array[Integer | String] col_index, ?width: Float | Integer | nil, ?hidden: bool, ?best_fit: bool, ?custom_width: bool, ?outline_level: Integer | nil, ?collapsed: bool) -> void
      def column(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.column(...)
      end

      # Adds a chart to the active sheet.
      #
      # @example
      #   s.chart(:bar) do |chart_builder|
      #     chart_builder.title("Quarterly Sales")
      #     chart_builder.series(values: "Sheet1!$B$2:$B$5", categories: "Sheet1!$A$2:$A$5", name: "Revenue")
      #   end
      #
      # @param type [Symbol, String, nil] The chart type (:bar, :col, :line, :pie, :scatter, :area, :doughnut, :radar).
      # @param opts [Hash] Additional chart options.
      # @yield [chart_builder]
      # @yieldparam chart_builder [Xlsxrb::ChartBuilder]
      # @return [void]
      # @api public
      #: (?Symbol | String? type, **untyped opts) ?{ (Xlsxrb::ChartBuilder) -> void } -> void
      def chart(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.chart(...)
      end

      # Adds a hyperlink to a cell.
      #
      # @example Positional URL
      #   s.hyperlink("A1", "https://example.com", display: "Example")
      #
      # @example Keyword location
      #   s.hyperlink("A1", location: "https://example.com", tooltip: "Go to Example")
      #
      # @param cell [String] The cell reference (e.g. "A1").
      # @param url [String, nil] The target URL or URI.
      # @param display [String, nil] Display text for the link.
      # @param tooltip [String, nil] Tooltip text when hovering.
      # @param location [String, nil] Destination location / URL (keyword alternative).
      # @return [void]
      # @api public
      #: (String cell, ?String? url, ?display: String?, ?tooltip: String?, ?location: String?) -> void
      def hyperlink(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.hyperlink(...)
      end

      # Sets the auto-filter range on the active sheet.
      #
      # @example
      #   s.auto_filter("A1:D100")
      #
      # @param ref [String] The cell range (e.g. "A1:D10").
      # @return [void]
      # @api public
      #: (String ref) -> void
      def auto_filter(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.auto_filter(...)
      end

      # Sets filter criteria for a column in the auto-filter.
      #
      # @example Simple values filter
      #   s.filter_column(0, ["Active", "Pending"])
      #
      # @param col_id [Integer] 0-based column index relative to auto-filter range.
      # @param filter_values [Array<String>, Hash] Values or filter specification hash.
      # @return [void]
      # @api public
      #: (Integer col_id, Array[String] | Hash[Symbol, untyped] filter_values) -> void
      def filter_column(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.filter_column(...)
      end

      # Configures sort state on a range.
      #
      # @example
      #   s.sort_state("A1:A10", [{ ref: "A1:A10", descending: true }])
      #
      # @param ref [String] The range to sort.
      # @param sort_conditions [Array<Hash>, Hash] Sort conditions array or options hash.
      # @param opts [Hash] Additional sort options.
      # @return [void]
      # @api public
      #: (String ref, Array[Hash[Symbol, untyped]] | Hash[Symbol, untyped] sort_conditions, **untyped opts) -> void
      def sort_state(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.sort_state(...)
      end

      # Adds data validation rules to a range.
      #
      # @example Dropdown list validation
      #   s.validate_data("B2:B100", type: "list", formula1: '"High,Medium,Low"')
      #
      # @param range [String] The cell range (e.g. "B2:B10").
      # @param type [String, Symbol] Validation type ("list", "whole", "decimal", "date", "time", "textLength", "custom").
      # @param opts [Hash] Validation options.
      # @return [void]
      # @api public
      #: (String range, ?type: String | Symbol, **untyped opts) -> void
      def validate_data(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.validate_data(...)
      end

      # Adds conditional formatting to a range.
      #
      # @example Highlight values greater than 100
      #   s.conditional_format("A1:A10", type: "cellIs", operator: "greaterThan", formula: 100, style: :highlight)
      #
      # @param range [String] The cell range (e.g. "A1:A10").
      # @param type [String, Symbol] Rule type ("cellIs", "colorScale", "dataBar", "expression").
      # @param opts [Hash] Rule options.
      # @return [void]
      # @api public
      #: (String range, ?type: String | Symbol, **untyped opts) -> void
      def conditional_format(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.conditional_format(...)
      end

      # Adds a formatted Excel Table to the sheet.
      #
      # @example
      #   s.table("A1:C10", columns: ["ID", "Name", "Total"], name: "SalesTable", style: "TableStyleMedium9")
      #
      # @param ref [String] The cell range for the table (e.g. "A1:D10").
      # @param columns [Array<String>, Array<Hash>] Column names or definitions.
      # @param name [String, nil] Table name.
      # @param display_name [String, nil] Display name.
      # @param style [String, nil] Table style name.
      # @param opts [Hash] Additional options.
      # @return [void]
      # @api public
      #: (String ref, columns: untyped, ?name: String?, ?display_name: String?, ?style: String?, **untyped opts) -> void
      def table(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.table(...)
      end

      # simplecov:disable
      # Edge case / untested delegation block
      #: () -> void
      def cleanup!(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.cleanup!(...)
      end
      # simplecov:enable

      # Adds a comment to a cell.
      #
      # @example
      #   s.comment("A1", "Reviewed and approved", author: "Auditor")
      #
      # @param cell [String, Integer] The cell reference (e.g. "A1").
      # @param text [String] The comment text.
      # @param author [String] The author name.
      # @return [void]
      # @api public
      #: (String | Integer cell, String text, ?author: String) -> void
      def comment(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.comment(...)
      end

      # Adds a Pivot Table to the sheet.
      #
      # @param source_ref [String] Source data range reference (e.g. "Sheet1!A1:D100").
      # @param row_fields [Array<String>] Field names for rows.
      # @param data_fields [Array<String>] Field names for data values.
      # @param col_fields [Array<String>] Field names for columns.
      # @param dest_ref [String] Target top-left cell reference (default: "E1").
      # @param name [String, nil] Pivot table name.
      # @param field_names [Array<String>, nil] Override field names.
      # @param items [Array, nil] Items configuration.
      # @param opts [Hash] Additional options.
      # @return [void]
      # @api public
      #: (String source_ref, row_fields: untyped, data_fields: untyped, ?col_fields: untyped, ?dest_ref: String, ?name: String?, ?field_names: untyped, ?items: untyped, **untyped opts) -> void
      def pivot_table(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.pivot_table(...)
      end

      # Adds sparklines to the sheet.
      #
      # @example
      #   s.sparkline_group(sparklines: [{ data_ref: "A1:E1", location_ref: "F1" }], type: "line")
      #
      # @param sparklines [Array<Hash>] Array of { data_ref:, location_ref: } hashes.
      # @param type [String, Symbol, nil] "line" (default), "column", or "stacked".
      # @param opts [Hash] Additional sparkline options.
      # @return [void]
      # @api public
      #: (sparklines: Array[Hash[Symbol, untyped]], ?type: (String | Symbol)?, **untyped opts) -> void
      def sparkline_group(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.sparkline_group(...)
      end

      # Sets workbook-level properties.
      #
      # @param opts [Hash] Workbook property options.
      # @return [void]
      # @api public
      #: (**untyped opts) -> void
      def workbook_property(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.workbook_property(...)
      end

      # Sets sheet properties (e.g. tab color, page setup flags).
      #
      # @example
      #   s.sheet_properties(:tab_color, "FF0000")
      #
      # @param name [Symbol, String] Property name.
      # @param value [Object] Property value.
      # @return [void]
      # @api public
      #: (Symbol | String name, untyped value) -> void
      def sheet_properties(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.sheet_properties(...)
      end

      # Adds a defined named range or formula.
      #
      # @example
      #   s.defined_name("TaxRate", "0.10")
      #
      # @param name [String] The name.
      # @param formula [String] The formula or range expression.
      # @param sheet_id [Integer, nil] Optional sheet scope.
      # @param hidden [Boolean] Whether the name is hidden.
      # @return [void]
      # @api public
      #: (String name, String formula, ?sheet_id: Integer | nil, ?hidden: bool) -> void
      def defined_name(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.defined_name(...)
      end

      # Freezes rows and/or columns for scrolling.
      #
      # @example Freeze top row
      #   s.freeze_pane(row: 1)
      #
      # @example Freeze first column and top 2 rows
      #   s.freeze_pane(row: 2, col: 1)
      #
      # @param row [Integer, nil] Number of rows to freeze.
      # @param col [Integer, nil] Number of columns to freeze.
      # @return [void]
      # @api public
      #: (?row: Integer | nil, ?col: Integer | nil) -> void
      def freeze_pane(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.freeze_pane(...)
      end

      # simplecov:disable
      # Edge case / untested delegation block
      # Sets the print area range for the sheet.
      #
      # @param ref [String] Range reference.
      # @return [void]
      # @api public
      #: (String ref) -> void
      def print_area(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.print_area(...)
      end
      # simplecov:enable

      # simplecov:disable
      # Edge case / untested delegation block
      # Configures repeating title rows and columns for printing.
      #
      # @param rows [String, nil] Row range to repeat (e.g. "1:2").
      # @param cols [String, nil] Column range to repeat (e.g. "A:B").
      # @return [void]
      # @api public
      #: (?rows: String?, ?cols: String?) -> void
      def print_titles(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.print_titles(...)
      end
      # simplecov:enable

      # Splits sheet view into panes.
      #
      # @param x_split [Numeric, nil] Horizontal split position.
      # @param y_split [Numeric, nil] Vertical split position.
      # @param top_left_cell [String, nil] Top-left visible cell in bottom-right pane.
      # @param active_pane [String, nil] Active pane identifier.
      # @param state [String, nil] Split state.
      # @return [void]
      # @api public
      #: (?x_split: Numeric | nil, ?y_split: Numeric | nil, ?top_left_cell: String?, ?active_pane: String?, ?state: String?) -> void
      def split_pane(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.split_pane(...)
      end

      # simplecov:disable
      # Edge case / untested delegation block
      # Protects the workbook structure.
      #
      # @param opts [Hash] Protection options.
      # @return [void]
      # @api public
      #: (**untyped opts) -> void
      def protect_workbook(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.protect_workbook(...)
      end
      # simplecov:enable

      # simplecov:disable
      # Edge case / untested delegation block
      # Sets core metadata property.
      #
      # @param name [String, Symbol] Property name.
      # @param value [Object] Property value.
      # @return [void]
      # @api public
      #: (String | Symbol name, untyped value) -> void
      def core_property(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.core_property(...)
      end
      # simplecov:enable

      # Sets the active/selected cell on the sheet.
      #
      # @param active_cell [String] Cell reference (e.g. "A1").
      # @param sqref [String, nil] Selection range.
      # @param pane [String, Symbol, nil] Pane identifier.
      # @return [void]
      # @api public
      #: (String active_cell, ?sqref: String?, ?pane: (String | Symbol)?) -> void
      def select_cell(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.select_cell(...)
      end

      # Configures page margins for printing.
      #
      # @param left [Float, nil] Left margin in inches.
      # @param right [Float, nil] Right margin in inches.
      # @param top [Float, nil] Top margin in inches.
      # @param bottom [Float, nil] Bottom margin in inches.
      # @param header [Float, nil] Header margin in inches.
      # @param footer [Float, nil] Footer margin in inches.
      # @return [void]
      # @api public
      #: (?left: Float | nil, ?right: Float | nil, ?top: Float | nil, ?bottom: Float | nil, ?header: Float | nil, ?footer: Float | nil) -> void
      def page_margins(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.page_margins(...)
      end

      # Configures page orientation, paper size, and print setup.
      #
      # @param orientation [String, Symbol, nil] "portrait" or "landscape" (or :portrait, :landscape).
      # @param paper_size [Integer, nil] Paper size index (e.g. 9 for A4, 1 for Letter).
      # @param opts [Hash] Additional options (scale, fit_to_width, fit_to_height).
      # @return [void]
      # @api public
      #: (?orientation: (String | Symbol)?, ?paper_size: Integer | nil, **untyped opts) -> void
      def page_setup(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.page_setup(...)
      end

      # Configures header and footer text for printing.
      #
      # @param opts [Hash] Header and footer specifications.
      # @return [void]
      # @api public
      #: (**untyped opts) -> void
      def header_footer(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.header_footer(...)
      end

      # Configures print options (e.g. gridlines, headings).
      #
      # @param name [Symbol, String] Print option name.
      # @param value [Object] Print option value.
      # @return [void]
      # @api public
      #: (Symbol | String name, untyped value) -> void
      def print_options(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.print_options(...)
      end

      # simplecov:disable
      # Edge case / untested delegation block
      # Sets document metadata properties (core, app, custom).
      #
      # @param core [Hash, nil] Core properties (title, creator, subject, etc.).
      # @param app [Hash, nil] App properties (company, manager).
      # @param custom [Hash, nil] Custom properties.
      # @return [void]
      # @api public
      #: (?core: Hash[untyped, untyped]?, ?app: Hash[untyped, untyped]?, ?custom: Hash[untyped, untyped]?) -> void
      def properties(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.properties(...)
      end
      # simplecov:enable

      # simplecov:disable
      # Edge case / untested delegation block
      # Sets app metadata property.
      #
      # @param name [String, Symbol] Property name.
      # @param value [Object] Property value.
      # @return [void]
      # @api public
      #: (String | Symbol name, untyped value) -> void
      def app_property(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.app_property(...)
      end
      # simplecov:enable

      # Protects the worksheet against modifications.
      #
      # @param opts [Hash] Protection options.
      # @return [void]
      # @api public
      #: (**untyped opts) -> void
      def protect_sheet(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.protect_sheet(...)
      end

      # simplecov:disable
      # Edge case / untested delegation block
      # Sets custom metadata property.
      #
      # @param name [String, Symbol] Property name.
      # @param value [Object] Property value.
      # @return [void]
      # @api public
      #: (String | Symbol name, untyped value) -> void
      def custom_property(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.custom_property(...)
      end
      # simplecov:enable

      # Inserts an image into the sheet.
      #
      # @param file_data [String] Binary image data or file content.
      # @param ext [String] Image extension ("png", "jpeg", etc.).
      # @param from_col [Integer] Starting column index (0-based).
      # @param from_row [Integer] Starting row index (0-based).
      # @param to_col [Integer] Ending column index (0-based).
      # @param to_row [Integer] Ending row index (0-based).
      # @param opts [Hash] Additional anchor and sizing options.
      # @return [void]
      # @api public
      #: (String file_data, ?ext: String, ?from_col: Integer, ?from_row: Integer, ?to_col: Integer, ?to_row: Integer, **untyped opts) -> void
      def image(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.image(...)
      end

      # Configures sheet view settings (zoom scale, grid lines visibility).
      #
      # @param name [Symbol, String] View setting name.
      # @param value [Object] View setting value.
      # @return [void]
      # @api public
      #: (Symbol | String name, untyped value) -> void
      def sheet_view(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.sheet_view(...)
      end

      # simplecov:disable
      # Edge case / untested delegation block
      # Adds a horizontal page break after the given row index.
      #
      # @param row_index [Integer] 0-based row index.
      # @return [void]
      # @api public
      #: (Integer row_index) -> void
      def page_break_row(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.page_break_row(...)
      end
      # simplecov:enable

      # simplecov:disable
      # Edge case / untested delegation block
      # Adds a vertical page break after the given column index.
      #
      # @param col_index [Integer] 0-based column index.
      # @return [void]
      # @api public
      #: (Integer col_index) -> void
      def page_break_col(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.page_break_col(...)
      end
      # simplecov:enable
    end

    # Adds a new sheet to the workbook and starts streaming rows into it.
    #
    # @param name [String, nil] Sheet name (max 31 characters).
    # @param opts [Hash] Sheet-level configuration.
    # @yield [sheet_proxy]
    # @yieldparam sheet_proxy [WorksheetProxy] The streaming sheet proxy.
    # @return [String] The sheet name.
    # @api public
    #: (?String? name, **untyped opts) ?{ (WorksheetProxy) -> void } -> untyped
    def sheet(name = nil, **opts)
      name ||= "Sheet#{@sheets.size + 1}"
      raise ArgumentError, "Sheet name '#{name}' must be <= 31 characters (Excel limitation)" if @strict_excel_mode && name.length > 31
      raise ArgumentError, "Sheet name '#{name}' contains invalid characters (ECMA-376 OOXML specification)" if name.match?(%r{[\[\]*?/\\]})
      raise ArgumentError, "Sheet name '#{name}' is already used. Excel requires unique sheet names." if @strict_excel_mode && @sheets.map { |s| s.respond_to?(:name) ? s.name.downcase : s.to_s.downcase }.include?(name.downcase)

      internal_sheet_setup(name)
      opts.each { |k, v| sheet_properties(k, v) }

      yield WorksheetProxy.new(self, @current_sheet) if block_given?
      @current_sheet
    end

    # Internal: Start or switch to a named sheet (internal helper).
    #: (?String? name) ?{ (WorksheetProxy) -> void } -> (WorksheetProxy | nil)
    def internal_sheet_setup(name = nil)
      flush_current_sheet
      name ||= "Sheet#{@sheets.size + 1}"
      @current_sheet = name
      @current_row_index = 0
      @current_tempfile = Tempfile.new(["xlsxrb_rows", ".xml"])
      @current_tempfile.binmode
      @current_row_writer = Ooxml::WorksheetWriter.new(@current_tempfile)
      @current_row_writer.instance_variable_set(:@started, true)

      @current_columns = []
      @current_charts = []
      @current_hyperlinks = []
      @current_auto_filter = nil
      @current_filter_columns = {}
      @current_sort_state = nil
      @current_data_validations = []
      @current_conditional_formats = []
      @current_tables = []
      @current_pivot_tables = []
      @current_sparkline_groups = []
      @current_comments = []
      @current_merge_cells = []
      @current_freeze_pane = nil
      @current_split_pane = nil
      @current_selection = nil
      @current_page_margins = nil
      @current_page_setup = {}
      @current_header_footer = {}
      @current_print_options = {}
      @current_sheet_protection = nil
      @current_images = []
      @current_shapes = []
      @current_sheet_properties = {}
      @current_sheet_view = {}
      @current_row_breaks = []
      @current_col_breaks = []
      @current_cells = {}

      return unless block_given?

      # simplecov:disable
      # Edge case / untested delegation block
      yield self
      flush_current_sheet
      # simplecov:enable
    end

    # Appends a row of values to the active sheet.
    #
    # @param values [Array<Object>, Hash] The cell values.
    # @param styles [String, Symbol, Array, Hash, nil] Style names or inline style definitions.
    # @param height [Float, Integer, nil] The row height in points (0 - 409).
    # @param hidden [Boolean] Whether the row is hidden.
    # @param custom_height [Boolean] Whether custom row height is set.
    # @param outline_level [Integer, nil] Grouping/outline level (0 - 7).
    # @return [void]
    # @api public
    #: (Array[untyped] | Hash[untyped, untyped] values, ?styles: untyped, ?height: Float | Integer | nil, ?hidden: bool, ?custom_height: bool, ?outline_level: Integer | nil) -> void
    def row(values, styles: nil, height: nil, hidden: false, custom_height: false, outline_level: nil)
      sheet if @current_sheet.nil?

      row_index = @current_row_index
      # See: https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
      if @strict_excel_mode
        raise ArgumentError, "Row index #{row_index} exceeds Excel limit of 1,048,576 rows" if row_index >= 1_048_576
        raise ArgumentError, "Row height #{height} must be between 0 and 409 points (Excel limitation)" if height && (height.negative? || height > 409)
      end
      @current_row_index += 1

      if values.is_a?(Hash)
        max_col = values.keys.map { |k| Elements::Cell.column_index(k) }.max || -1
        cells_array = Array.new(max_col + 1)
        values.each do |k, v|
          idx = Elements::Cell.column_index(k)
          cells_array[idx] = v
        end
        values = cells_array
      end

      if styles.is_a?(Hash)
        expanded_styles = {}
        styles.each do |k, v|
          if k.is_a?(Range) || k.is_a?(Array)
            k.each { |idx| expanded_styles[Elements::Cell.column_index(idx)] = v }
          else
            expanded_styles[Elements::Cell.column_index(k)] = v
          end
        end
        max_col_style = expanded_styles.keys.max || -1
        styles_array = Array.new(max_col_style + 1)
        expanded_styles.each do |idx, v|
          styles_array[idx] = v
        end
        styles = styles_array
      end

      # Auto-detect Date / Time for built-in styles
      values.each_with_index do |val, idx|
        cell_style = styles.is_a?(Array) ? styles[idx] : styles

        if val.is_a?(Date) && cell_style.nil?
          style("__xlsxrb_date", number_format: "yyyy-mm-dd") unless @styles.key?("__xlsxrb_date")
          styles = [] if styles.nil?
          styles = Array.new(values.size, styles) unless styles.is_a?(Array)
          styles[idx] = "__xlsxrb_date"
        elsif val.is_a?(Time) && cell_style.nil?
          style("__xlsxrb_time", number_format: "yyyy-mm-dd hh:mm:ss") unless @styles.key?("__xlsxrb_time")
          styles = [] if styles.nil?
          styles = Array.new(values.size, styles) unless styles.is_a?(Array)
          styles[idx] = "__xlsxrb_time"
        end
      end

      has_charts = @current_charts && !@current_charts.empty?
      row_num = row_index + 1 if has_charts

      max_len = values.size
      max_len = [max_len, styles.size].max if styles.is_a?(Array)
      # See: https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
      raise ArgumentError, "Row contains #{max_len} columns, exceeding Excel limit of 16_384 columns" if @strict_excel_mode && max_len > 16_384

      if has_charts || @strict_excel_mode
        @current_cells ||= {} if has_charts
        max_len.times do |col_idx|
          val = col_idx < values.size ? values[col_idx] : nil
          next if val.nil?

          raise ArgumentError, "Invalid cell value type or value: #{val.class} for value #{val.inspect}" unless val.nil? || val.is_a?(String) || (val.is_a?(Numeric) && !(val.is_a?(Float) && (val.infinite? || val.nan?))) || val.is_a?(TrueClass) || val.is_a?(FalseClass) || val.is_a?(Date) || val.is_a?(Time) || val.is_a?(Elements::Formula) || (val.is_a?(Hash) && val.key?(:formula)) || val.is_a?(Elements::RichText) || val.is_a?(Elements::CellError)

          # See: https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
          raise ArgumentError, "Cell text length #{val.length} exceeds Excel limit of 32,767 characters" if @strict_excel_mode && val.is_a?(String) && val.length > 32_767

          if has_charts
            addr = "#{Elements::Cell.column_letter(col_idx)}#{row_num}"
            @current_cells[addr] = val
          end
        end
      end

      attrs = nil
      if height || hidden || outline_level
        attrs = {}
        attrs[:height] = height if height
        attrs[:hidden] = true if hidden
        attrs[:custom_height] = custom_height || !height.nil?
        attrs[:outline_level] = outline_level if outline_level
      end

      @current_row_writer.write_row_values(row_index, values, styles: styles, style_map: @style_name_to_id, sst: @sst, sst_index: @sst_index, attrs: attrs)
    end

    # Sets column formatting and properties for one or multiple columns.
    #
    # @param index [Integer, String, Range, Array] Column index (0-based) or letter ("A".."D").
    # @param width [Float, Integer, nil] Column width in character units (0 - 255).
    # @param hidden [Boolean] Whether the column is hidden.
    # @param custom_width [Boolean] Whether custom width is set.
    # @param outline_level [Integer, nil] Grouping/outline level.
    # @return [void]
    # @api public
    #: (Integer | String | Range[Integer | String] | Array[Integer | String] index, ?width: Float | Integer | nil, ?hidden: bool, ?custom_width: bool, ?outline_level: Integer | nil) -> void
    def column(index, width: nil, hidden: false, custom_width: false, outline_level: nil)
      raise ArgumentError, "Column width #{width} must be between 0 and 255 characters (Excel limitation)" if @strict_excel_mode && width && (width.negative? || width > 255)

      indices = case index
                when Range, Array
                  index.map { |i| Elements::Cell.column_index(i) }
                else
                  [Elements::Cell.column_index(index)]
                end

      sheet if @current_sheet.nil?

      indices.each do |idx|
        @current_columns << { index: idx, width: width, hidden: hidden, custom_width: custom_width || !width.nil?, outline_level: outline_level }
      end
    end

    # Adds a chart to the current worksheet.
    #
    # @param options [Hash] Chart options.
    # @yield [builder]
    # @yieldparam builder [Xlsxrb::ChartBuilder]
    # @return [void]
    # @api public
    #: (**untyped options) ?{ (ChartBuilder) -> void } -> void
    def chart(**options)
      sheet if @current_sheet.nil?

      if block_given?
        builder = ChartBuilder.new
        yield builder
        options = builder.options.merge(options)
      end

      @current_charts << options
    end

    # Adds a hyperlink on a cell.
    #
    # @param cell [String, Integer] Cell coordinate (e.g. "A1").
    # @param url [String, nil] Target URL.
    # @param display [String, nil] Display text.
    # @param tooltip [String, nil] Tooltip text.
    # @param location [String, nil] Internal location.
    # @return [void]
    # @api public
    #: (String | Integer cell, ?String? url, ?display: String?, ?tooltip: String?, ?location: String?) -> void
    def hyperlink(cell, url = nil, display: nil, tooltip: nil, location: nil)
      sheet if @current_sheet.nil?
      link = { cell: cell }
      link[:url] = url if url
      link[:display] = display if display
      link[:tooltip] = tooltip if tooltip
      link[:location] = location if location
      @current_hyperlinks << link
    end

    # Sets the auto-filter range on the active sheet.
    #
    # @param range [String] Cell range (e.g. "A1:E100").
    # @return [void]
    # @api public
    #: (String range) -> void
    def auto_filter(range)
      sheet if @current_sheet.nil?
      @current_auto_filter = range
    end

    # Sets filter criteria on a column in the auto-filter.
    #
    # @param col_id [Integer] 0-based column index.
    # @param filter [Hash] Filter criteria.
    # @return [void]
    # @api public
    #: (untyped col_id, untyped filter) -> untyped
    def filter_column(col_id, filter)
      sheet if @current_sheet.nil?
      @current_filter_columns[col_id] = filter
    end

    # Configures column sort state on the active sheet.
    #
    # @param ref [String] The sorted range.
    # @param sort_conditions [Array<Hash>] Sort conditions.
    # @param opts [Hash] Additional options.
    # @return [void]
    # @api public
    #: (untyped ref, untyped sort_conditions, **untyped opts) -> untyped
    def sort_state(ref, sort_conditions, **opts)
      sheet if @current_sheet.nil?
      @current_sort_state = { ref: ref, sort_conditions: sort_conditions }.merge(opts)
    end

    # Adds a data validation rule.
    #
    # @param sqref [String] Cell range.
    # @param opts [Hash] Validation options.
    # @return [void]
    # @api public
    #: (untyped sqref, **untyped opts) -> void
    def validate_data(sqref, **opts)
      sheet if @current_sheet.nil?
      @current_data_validations << opts.merge(sqref: sqref)
    end

    # Adds a conditional formatting rule.
    #
    # @param sqref [String] Cell range.
    # @param opts [Hash] Rule options.
    # @return [void]
    # @api public
    #: (untyped sqref, **untyped opts) -> void
    def conditional_format(sqref, **opts)
      sheet if @current_sheet.nil?
      @current_conditional_formats << opts.merge(sqref: sqref)
    end

    # Adds an Excel Table (ListObject).
    #
    # @param ref [String] Table range.
    # @param columns [Array<String>] Column names.
    # @param name [String, nil] Table name.
    # @param display_name [String, nil] Display name.
    # @param style [String, nil] Table style.
    # @param opts [Hash] Additional options.
    # @return [void]
    # @api public
    #: (untyped ref, columns: untyped, ?name: untyped, ?display_name: untyped, ?style: untyped, **untyped opts) -> void
    def table(ref, columns:, name: nil, display_name: nil, style: nil, **opts)
      sheet if @current_sheet.nil?
      tbl = { ref: ref, columns: columns }
      tbl[:name] = name if name
      tbl[:display_name] = display_name if display_name
      tbl[:style] = style if style
      tbl.merge!(opts)
      @current_tables << tbl
    end

    # Adds a Pivot Table.
    #
    # @param source_ref [String] Source data range.
    # @param row_fields [Array<Integer>] 0-based field indices for rows.
    # @param data_fields [Array<Hash>] Data fields.
    # @param col_fields [Array<Integer>] 0-based field indices for columns.
    # @param dest_ref [String] Top-left destination cell (default: "E1").
    # @param name [String, nil] Pivot table name.
    # @param field_names [Array<String>, nil] Override field names.
    # @param items [Array, nil] Items configuration.
    # @param opts [Hash] Additional options.
    # @return [void]
    # @api public
    #: (untyped source_ref, row_fields: untyped, data_fields: untyped, ?col_fields: untyped, ?dest_ref: untyped, ?name: untyped, ?field_names: untyped, ?items: untyped, **untyped opts) -> void
    def pivot_table(source_ref, row_fields:, data_fields:, col_fields: [], dest_ref: "E1", name: nil, field_names: nil, items: nil)
      sheet if @current_sheet.nil?
      @current_pivot_tables ||= []
      @current_pivot_tables << {
        source_ref: source_ref, row_fields: row_fields,
        data_fields: data_fields, col_fields: col_fields,
        dest_ref: dest_ref, name: name,
        field_names: field_names, items: items
      }
    end

    # Adds a comment to a cell.
    #
    # @param cell [String, Integer] Cell coordinate (e.g. "A1").
    # @param text [String] Comment text.
    # @param author [String] Author name.
    # @return [void]
    # @api public
    #: (String | Integer cell, String text, ?author: ::String) -> void
    def comment(cell, text, author: "Author")
      sheet if @current_sheet.nil?
      @current_comments << { cell: cell, text: text, author: author }
    end

    # Adds a sparkline group.
    #
    # @param sparklines [Array<Hash>] Sparkline definitions.
    # @param type [String, nil] "line", "column", or "stacked".
    # @param opts [Hash] Additional options.
    # @return [void]
    # @api public
    #: (sparklines: untyped, ?type: untyped, **untyped opts) -> void
    def sparkline_group(sparklines:, type: nil, **opts)
      sheet if @current_sheet.nil?
      group = { sparklines: sparklines }
      group[:type] = type if type
      group.merge!(opts)
      @current_sparkline_groups << group
    end

    # Merges a range of cells into a single cell.
    #
    # @param range [String, Hash, nil] Cell range (e.g. "A1:B2") or hash of coordinates.
    # @param row [Integer, nil] Single row index.
    # @param col_start [Integer, String, nil] Starting column.
    # @param col_end [Integer, String, nil] Ending column.
    # @param row_start [Integer, nil] Starting row index.
    # @param row_end [Integer, nil] Ending row index.
    # @return [void]
    # @api public
    #: (?(String | Hash[Symbol, Integer | String])? range, ?row: Integer?, ?col_start: (Integer | String)?, ?col_end: (Integer | String)?, ?row_start: Integer?, ?row_end: Integer?) -> void
    def merge(range = nil, row: nil, col_start: nil, col_end: nil, row_start: nil, row_end: nil)
      sheet if @current_sheet.nil?
      if range.is_a?(Hash)
        row = range[:row]
        row_start = range[:row_start]
        row_end = range[:row_end]
        col_start = range[:col_start]
        col_end = range[:col_end]
        range = nil
      end

      if range
        raise ArgumentError, "Invalid merge range format: '#{range}'. Expected format like 'A1:B2'." if @strict_excel_mode && !range.match?(/^[A-Za-z]{1,3}\d+(:[A-Za-z]{1,3}\d+)?$/)

        @current_merge_cells << range
      else
        r_start = row || row_start || 0
        r_end = row || row_end || 0
        c_start = Elements::Cell.column_index(col_start || 0)
        c_end = Elements::Cell.column_index(col_end || 0)
        start_ref = "#{Xlsxrb::Elements::Cell.column_letter(c_start)}#{r_start + 1}"
        end_ref = "#{Xlsxrb::Elements::Cell.column_letter(c_end)}#{r_end + 1}"
        @current_merge_cells << "#{start_ref}:#{end_ref}"
      end
    end

    # Freezes window panes at the given row and column.
    #
    # @param row [Integer] Number of rows to freeze.
    # @param col [Integer, String] Number of columns to freeze.
    # @return [void]
    # @api public
    #: (?row: Integer, ?col: (Integer | String)) -> void
    def freeze_pane(row: 0, col: 0)
      col = Elements::Cell.column_index(col)
      sheet if @current_sheet.nil?
      @current_freeze_pane = { row: row, col: col }
    end

    # Splits window panes without freezing.
    #
    # @param x_split [Integer] Horizontal split in points.
    # @param y_split [Integer] Vertical split in points.
    # @param top_left_cell [String, nil] Top-left cell reference in bottom-right pane.
    # @return [void]
    # @api public
    #: (?x_split: ::Integer, ?y_split: ::Integer, ?top_left_cell: String?) -> void
    def split_pane(x_split: 0, y_split: 0, top_left_cell: nil)
      sheet if @current_sheet.nil?
      @current_split_pane = { x_split: x_split, y_split: y_split, top_left_cell: top_left_cell }
    end

    # Sets the active cell selection on the current sheet.
    #
    # @param active_cell [String] Active cell reference (e.g. "A1").
    # @param sqref [String, nil] Selection range.
    # @param pane [String, Symbol, nil] Target pane.
    # @return [void]
    # @api public
    #: (String active_cell, ?sqref: String?, ?pane: (String | Symbol)?) -> void
    def select_cell(active_cell, sqref: nil, pane: nil)
      sheet if @current_sheet.nil?
      @current_selection = { active_cell: active_cell, sqref: sqref || active_cell }
      @current_selection[:pane] = pane if pane
    end

    # Sets page margins in inches for printing.
    #
    # @param left [Float, nil] Left margin.
    # @param right [Float, nil] Right margin.
    # @param top [Float, nil] Top margin.
    # @param bottom [Float, nil] Bottom margin.
    # @param header [Float, nil] Header margin.
    # @param footer [Float, nil] Footer margin.
    # @return [void]
    # @api public
    #: (?left: Float?, ?right: Float?, ?top: Float?, ?bottom: Float?, ?header: Float?, ?footer: Float?) -> void
    def page_margins(left: nil, right: nil, top: nil, bottom: nil, header: nil, footer: nil)
      sheet if @current_sheet.nil?
      @current_page_margins = { left: left, right: right, top: top, bottom: bottom, header: header, footer: footer }.compact
    end

    # Sets page setup configuration.
    #
    # @param opts [Hash] Page setup options (e.g. orientation: :landscape).
    # @return [void]
    # @api public
    #: (**untyped opts) -> void
    def page_setup(**opts)
      sheet if @current_sheet.nil?
      @current_page_setup.merge!(opts)
    end

    # Configures headers and footers for printing.
    #
    # @param opts [Hash] Header/footer options.
    # @return [void]
    # @api public
    #: (**untyped opts) -> void
    def header_footer(**opts)
      sheet if @current_sheet.nil?
      @current_header_footer.merge!(opts)
    end

    # Sets a print option (e.g. grid_lines: true).
    #
    # @param name [Symbol] Option name.
    # @param value [Object] Option value.
    # @return [void]
    # @api public
    #: (Symbol name, untyped value) -> void
    def print_options(name, value)
      sheet if @current_sheet.nil?
      @current_print_options[name] = value
    end

    # Sets sheet-level protection with optional password hashing.
    #
    # @param opts [Hash] Protection options.
    # @return [void]
    # @api public
    #: (**untyped opts) -> void
    def protect_sheet(**opts)
      sheet if @current_sheet.nil?
      normalized = opts.dup
      plain_password = normalized[:password]
      needs_hash = plain_password.is_a?(String) && !plain_password.empty? &&
                   normalized[:algorithm_name].nil? && normalized[:hash_value].nil? &&
                   normalized[:salt_value].nil? && normalized[:spin_count].nil? &&
                   !plain_password.match?(/\A[0-9A-Fa-f]{4}\z/)
      if needs_hash
        normalized.delete(:password)
        normalized.merge!(Xlsxrb::Ooxml::Utils.hash_password(plain_password))
      end
      @current_sheet_protection = normalized
    end

    # Inserts an embedded image into the active sheet.
    #
    # @param file_data [String] Binary image data.
    # @param ext [String] File extension ("png", "jpeg", etc.).
    # @param from_col [Integer] Top-left starting column.
    # @param from_row [Integer] Top-left starting row.
    # @param to_col [Integer] Bottom-right ending column.
    # @param to_row [Integer] Bottom-right ending row.
    # @param opts [Hash] Additional options.
    # @return [void]
    # @api public
    #: (String file_data, ?ext: ::String, ?from_col: ::Integer, ?from_row: ::Integer, ?to_col: ::Integer, ?to_row: ::Integer, **untyped opts) -> void
    def image(file_data, ext: "png", from_col: 0, from_row: 0, to_col: 5, to_row: 10, **opts)
      sheet if @current_sheet.nil?
      img = { file_data: file_data, ext: ext, from_col: from_col, from_row: from_row, to_col: to_col, to_row: to_row }
      img.merge!(opts)
      @current_images << img
    end

    # Inserts a drawing shape.
    #
    # @param preset [String] Preset shape name.
    # @param text [String, nil] Label text.
    # @param from_col [Integer] Top-left column.
    # @param from_row [Integer] Top-left row.
    # @param to_col [Integer] Bottom-right column.
    # @param to_row [Integer] Bottom-right row.
    # @param opts [Hash] Additional options.
    # @return [void]
    # @api public
    #: (**untyped opts) -> void
    def shape(preset: "rect", text: nil, from_col: 0, from_row: 0, to_col: 5, to_row: 5, **opts)
      sheet if @current_sheet.nil?
      shape = { preset: preset, text: text, from_col: from_col, from_row: from_row, to_col: to_col, to_row: to_row }
      shape[:name] = opts.delete(:name) || "Shape #{@current_shapes.size + 1}"
      shape.merge!(opts)
      @current_shapes << shape
    end

    # Sets a sheet-level property (e.g. tab_color: "FF0000").
    #
    # @param name [Symbol] Property name.
    # @param value [Object] Property value.
    # @return [void]
    # @api public
    #: (Symbol name, untyped value) -> void
    def sheet_properties(name, value)
      sheet if @current_sheet.nil?
      @current_sheet_properties[name] = value
    end

    # Sets a sheet view property (e.g. zoom_scale: 120).
    #
    # @param name [Symbol] View property name.
    # @param value [Object] View property value.
    # @return [void]
    # @api public
    #: (Symbol name, untyped value) -> void
    def sheet_view(name, value)
      sheet if @current_sheet.nil?
      @current_sheet_view[name] = value
    end

    # Inserts a horizontal page break before a row.
    #
    # @param row_num [Integer] 1-based row number.
    # @return [void]
    # @api public
    #: (Integer row_num) -> void
    def page_break_row(row_num)
      # simplecov:disable
      # Edge case / untested delegation block
      sheet if @current_sheet.nil?
      @current_row_breaks << row_num
      # simplecov:enable
    end

    # Inserts a vertical page break before a column.
    #
    # @param col_index [Integer, String] 0-based column index or letter ("B").
    # @return [void]
    # @api public
    #: (Integer col_index) -> void
    def page_break_col(col_index)
      # simplecov:disable
      # Edge case / untested delegation block
      col_index = Elements::Cell.column_index(col_index)
      sheet if @current_sheet.nil?
      @current_col_breaks << col_index
      # simplecov:enable
    end

    # --- Workbook-Level Methods ---

    # Adds a defined name.
    #
    # @param name [String] The defined name.
    # @param value [String] The formula or value expression.
    # @param sheet [String, nil] Local sheet name.
    # @param hidden [Boolean] Whether the defined name is hidden.
    # @return [void]
    # @api public
    #: (String name, String value, ?sheet: String?, ?hidden: bool) -> void
    def defined_name(name, value, sheet: nil, hidden: false)
      entry = { name: name, value: value, hidden: hidden }
      if sheet
        # local_sheet_id will be resolved at close time
        entry[:local_sheet_name] = sheet
      end
      @defined_names << entry
    end

    # Sets the print area for the current or named sheet.
    #
    # @param range [String] Cell range (e.g. "A1:G50").
    # @param sheet [String, nil] Target sheet name.
    # @return [void]
    # @api public
    #: (String range, ?sheet: String?) -> void
    def print_area(range, sheet: nil)
      # simplecov:disable
      # Edge case / untested delegation block
      sheet_name = sheet || @current_sheet || "Sheet1"
      value = "'#{sheet_name}'!#{absolute_range(range)}"
      @defined_names.reject! { |dn| dn[:name] == "_xlnm.Print_Area" && dn[:local_sheet_name] == sheet_name }
      defined_name("_xlnm.Print_Area", value, sheet: sheet_name)
      # simplecov:enable
    end

    # Sets print titles for the current or named sheet.
    #
    # @param rows [String, nil] Repeating row range (e.g. "1:2").
    # @param cols [String, nil] Repeating column range (e.g. "A:B").
    # @param sheet [String, nil] Target sheet name.
    # @return [void]
    # @api public
    #: (?rows: String?, ?cols: String?, ?sheet: String?) -> void
    def print_titles(rows: nil, cols: nil, sheet: nil)
      sheet_name = sheet || @current_sheet || "Sheet1"
      parts = []
      parts << "'#{sheet_name}'!$#{cols.sub(":", ":$")}" if cols
      parts << "'#{sheet_name}'!$#{rows.sub(":", ":$")}" if rows
      value = parts.join(",")
      @defined_names.reject! { |dn| dn[:name] == "_xlnm.Print_Titles" && dn[:local_sheet_name] == sheet_name }
      defined_name("_xlnm.Print_Titles", value, sheet: sheet_name)
    end

    # Sets workbook protection.
    #
    # @param opts [Hash] Protection options.
    # @return [void]
    # @api public
    #: (**String | Integer | bool | nil opts) -> void
    def protect_workbook(**opts)
      @workbook_protection = opts
    end

    # Sets a core document metadata property.
    #
    # @param name [Symbol] Property name.
    # @param value [String, Integer, Time] Property value.
    # @return [void]
    # @api public
    #: (Symbol name, String | Integer | Time value) -> void
    def core_property(name, value)
      @core_properties[name] = value
    end

    # Sets an app document property.
    #
    # @param name [Symbol] Property name.
    # @param value [String, Integer, Time] Property value.
    # @return [void]
    # @api public
    #: (Symbol name, String | Integer | Time value) -> void
    def app_property(name, value)
      # simplecov:disable
      # Edge case / untested delegation block
      @app_properties[name] = value
      # simplecov:enable
    end

    # Sets multiple core and/or app properties.
    #
    # @param core [Hash, nil] Core properties.
    # @param app [Hash, nil] App properties.
    # @param custom [Hash, nil] Custom properties.
    # @return [void]
    # @api public
    #: (?core: Hash[Symbol, String | Integer | Time]?, ?app: Hash[Symbol, String | Integer | Time]?, ?custom: Hash[String | Symbol, untyped]?) -> void
    def properties(core: nil, app: nil, custom: nil)
      # simplecov:disable
      # Edge case / untested delegation block
      core&.each { |k, v| core_property(k, v) }
      app&.each { |k, v| app_property(k, v) }
      custom&.each { |k, v| custom_property(k.to_s, v) }
      # simplecov:enable
    end

    # Adds a custom document property.
    #
    # @param name [String] Property name.
    # @param value [String, Integer, Float, Boolean, Time] Property value.
    # @param type [Symbol] Value type.
    # @return [void]
    # @api public
    #: (String name, String | Integer | Float | bool | Time value, ?type: ::Symbol) -> void
    def custom_property(name, value, type: :string)
      # simplecov:disable
      # Edge case / untested delegation block
      @custom_properties << { name: name, value: value, type: type }
      # simplecov:enable
    end

    # Finalizes and writes all streaming sheet contents to the destination target.
    #
    # @return [void]
    # @api public
    #: () -> untyped
    def close
      raise ArgumentError, "Workbook must contain at least one sheet (Excel limitation)" if @strict_excel_mode && @sheets.empty? && @current_sheet.nil?

      Xlsxrb.in_span("StreamWriter#close") do
        flush_current_sheet

        styles_definition = {
          fonts: @style_writer.fonts.dup,
          fills: @style_writer.fills.dup,
          borders: @style_writer.borders.dup,
          xf_entries: @style_writer.xf_entries.dup,
          num_fmts: @style_writer.num_fmts.dup
        }

        resolved_names = resolve_defined_names(@defined_names, @sheets)

        Ooxml::WorkbookWriter.write(
          @target,
          sheets: @sheets,
          shared_strings: @sst,
          styles: styles_definition,
          defined_names: resolved_names.empty? ? nil : resolved_names,
          core_properties: @core_properties.empty? ? nil : @core_properties,
          app_properties: @app_properties.empty? ? nil : @app_properties,
          custom_properties: @custom_properties.empty? ? nil : @custom_properties,
          workbook_protection: @workbook_protection
        )
      end
    ensure
      cleanup!
    end

    # Explicitly removes any temporary files created during streaming.
    #
    # @return [void]
    # @api public
    #: () -> void
    def cleanup!
      @tempfiles.each do |tmp|
        tmp.close
        tmp.unlink
      end
      @tempfiles.clear
    end

    private

    #: (untyped range) -> untyped
    def absolute_range(range)
      # simplecov:disable
      # Edge case / untested delegation block
      range.gsub(/([A-Z]+)(\d+)/, '$\1$\2')
      # simplecov:enable
    end

    #: (untyped names, untyped sheets) -> untyped
    def resolve_defined_names(names, sheets)
      sheet_names = sheets.map { |s| s[:name] }
      names.map do |dn|
        resolved = dn.dup
        if dn[:local_sheet_name]
          idx = sheet_names.index(dn[:local_sheet_name])
          resolved[:local_sheet_id] = idx if idx
          resolved.delete(:local_sheet_name)
        end
        resolved
      end
    end

    #: () -> (nil | untyped)
    def flush_current_sheet
      return unless @current_sheet

      @current_tempfile.close

      sheet_data = { name: @current_sheet, rows_tmp_path: @current_tempfile.path, columns: @current_columns }
      sheet_data[:cells] = @current_cells if @current_cells && !@current_cells.empty?
      @current_cells = nil
      sheet_data[:charts] = @current_charts unless @current_charts.empty?
      sheet_data[:hyperlinks] = @current_hyperlinks unless @current_hyperlinks.empty?
      sheet_data[:auto_filter] = @current_auto_filter if @current_auto_filter
      sheet_data[:filter_columns] = @current_filter_columns unless @current_filter_columns.empty?
      sheet_data[:sort_state] = @current_sort_state if @current_sort_state
      sheet_data[:data_validations] = @current_data_validations unless @current_data_validations.empty?
      sheet_data[:conditional_formats] = @current_conditional_formats unless @current_conditional_formats.empty?
      sheet_data[:tables] = @current_tables unless @current_tables.empty?
      sheet_data[:pivot_tables] = @current_pivot_tables unless @current_pivot_tables.empty?
      sheet_data[:sparkline_groups] = @current_sparkline_groups unless @current_sparkline_groups.empty?
      sheet_data[:comments] = @current_comments unless @current_comments.empty?
      sheet_data[:merge_cells] = @current_merge_cells unless @current_merge_cells.empty?
      sheet_data[:freeze_pane] = @current_freeze_pane if @current_freeze_pane
      sheet_data[:split_pane] = @current_split_pane if @current_split_pane
      sheet_data[:selection] = @current_selection if @current_selection
      sheet_data[:page_margins] = @current_page_margins if @current_page_margins
      sheet_data[:page_setup] = @current_page_setup unless @current_page_setup.empty?
      sheet_data[:header_footer] = @current_header_footer unless @current_header_footer.empty?
      sheet_data[:print_options] = @current_print_options unless @current_print_options.empty?
      sheet_data[:sheet_protection] = @current_sheet_protection if @current_sheet_protection
      sheet_data[:images] = @current_images unless @current_images.empty?
      sheet_data[:shapes] = @current_shapes unless @current_shapes.empty?
      sheet_data[:sheet_properties] = @current_sheet_properties unless @current_sheet_properties.empty?
      sheet_data[:sheet_view] = @current_sheet_view unless @current_sheet_view.empty?
      sheet_data[:row_breaks] = @current_row_breaks unless @current_row_breaks.empty?
      sheet_data[:col_breaks] = @current_col_breaks unless @current_col_breaks.empty?
      @sheets << sheet_data

      @tempfiles << @current_tempfile
      @current_sheet = nil
      @current_tempfile = nil
      @current_row_writer = nil
    end
  end
end
