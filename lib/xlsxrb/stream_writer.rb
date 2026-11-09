# frozen_string_literal: true

# rbs_inline: enabled

require "tempfile"
require_relative "style_builder"
require_relative "chart_builder"
require_relative "dsl_helpers"
require_relative "worksheet_builder"
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
    OMITTED = Object.new.freeze
    private_constant :OMITTED

    # @return [String, nil] Name of the currently active worksheet.
    # @api private
    #: String?
    attr_reader :current_sheet

    # Sets whether to treat strings starting with formula prefixes as plain text.
    #
    # @param val [Boolean]
    # @return [Boolean]
    #: (bool) -> bool
    attr_writer :escape_formulas

    # Returns whether formula escaping is enabled.
    #
    # @return [Boolean]
    # @api public
    #: () -> bool
    def escape_formulas?
      @escape_formulas ? true : false
    end

    # Returns whether the shared strings table is enabled.
    #
    # @return [Boolean]
    # @api public
    #: () -> bool
    def shared_strings?
      @shared_strings ? true : false
    end

    # @return [Integer] The ZIP deflate compression level (0..9).
    #: Integer
    attr_reader :compression_level

    # @return [Hash{String => StyleBuilder}]
    #: Hash[String, StyleBuilder]
    attr_reader :styles

    # Initializes a streaming writer context.
    #
    # @param target [String, IO, StringIO] Destination file path or writable IO stream.
    # @param strict_excel_mode [Boolean] Whether to enforce Microsoft Excel specification limits.
    # @param escape_formulas [Boolean] Whether to treat strings starting with formula prefixes as plain text.
    # @param default_font [Hash, nil] Optional default font options (e.g. `{ name: "Arial", size: 12 }`).
    # @param shared_strings [Boolean] Whether to intern strings into a shared strings table (`xl/sharedStrings.xml`).
    #   When set to `false`, strings are written directly as inline strings (`<c t="inlineStr">`), maintaining flat
    #   $O(1)$ constant memory regardless of row count or unique string cardinality. Note that disabling shared strings
    #   may slightly increase archive size (typically 3-5%) on datasets with high repetition.
    # @param compression_level [Integer, nil] ZIP deflate compression level (0..9, or Zlib::DEFAULT_COMPRESSION
    #   for zlib's level 6). Defaults to {Xlsxrb::DEFAULT_COMPRESSION_LEVEL}. Higher levels write smaller files more slowly;
    #   level 0 stores entries without compression. See docs/adr/0009-default-zip-compression-level.md.
    #: (untyped target, ?strict_excel_mode: bool, ?escape_formulas: bool, ?default_font: untyped?, ?shared_strings: bool, ?compression_level: Integer | nil) -> void
    def initialize(target, strict_excel_mode: true, escape_formulas: false, default_font: nil, shared_strings: true, compression_level: DEFAULT_COMPRESSION_LEVEL)
      @target = target
      @strict_excel_mode = strict_excel_mode
      @escape_formulas = escape_formulas ? true : false
      @shared_strings = shared_strings ? true : false
      @compression_level = compression_level || DEFAULT_COMPRESSION_LEVEL
      @io = target.is_a?(String) ? File.open(target, "wb") : target
      @owns_io = target.is_a?(String)
      @zip = Ooxml::ZipWriter.new(@io, compression_level: @compression_level)
      if @shared_strings
        @sst = []
        @sst_index = {}
      else
        @sst = nil
        @sst_index = nil
      end
      @sheets = []
      @current_sheet = nil
      @current_sheet_index = 0
      @current_row_index = 0
      @sheet_entry_started = false
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
      @current_pivot_tables = []
      @current_sparkline_groups = []
      @current_comments = []
      @current_threaded_comments = []
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
      @current_sheet_format = {}
      @current_sheet_view = {}
      @current_row_breaks = []
      @current_col_breaks = []
      @current_protected_ranges = []
      @current_cell_watches = []
      @current_ignored_errors = []
      @current_phonetic_properties = nil
      @current_sheet_calc_pr = nil
      @current_background_image = nil
      @current_print_area = nil
      @current_print_titles = nil
      @current_row_outline_level = nil
      @current_row_collapsed = false
      @current_row_hidden = false
      @current_cells = {}
      @styles = {} # { style_name => StyleBuilder }
      @style_writer = Ooxml::Writer.new(default_font: default_font)
      @style_name_to_id = {}
      style("__xlsxrb_date", number_format: DslHelpers::DATE_NUMBER_FORMAT)
      style("__xlsxrb_time", number_format: DslHelpers::TIME_NUMBER_FORMAT)
      # Workbook-level settings
      @defined_names = []
      @core_properties = {}
      @app_properties = {}
      @custom_properties = []
      @workbook_protection = nil
      @workbook_properties = { update_links: "never" }
      @active_sheet = nil
      @first_sheet = nil
      @show_sheet_tabs = nil
      @tab_ratio = nil
    end

    # Gets or sets the active worksheet tab in the workbook window.
    #
    # @param name_or_index [String, Symbol, Integer, nil] Sheet name or 0-based sheet index.
    # @return [String, Symbol, Integer, nil]
    # @api public
    # Active worksheet tab in the workbook window.
    # @return [String, Symbol, Integer, nil]
    # @api public
    attr_accessor :active_sheet #: (String | Symbol | Integer | nil)

    # First sheet tab visible in the sheet tab scroll area.
    # @return [String, Symbol, Integer, nil]
    # @api public
    attr_accessor :first_sheet #: (String | Symbol | Integer | nil)

    # Sets whether sheet tabs are visible in the workbook window.
    # @param visible [Boolean, nil]
    # @return [Boolean, nil]
    # @api public
    attr_writer :show_sheet_tabs #: (bool?) -> bool?

    # Returns whether sheet tabs are visible in the workbook window.
    # @return [Boolean]
    # @api public
    #: () -> bool
    def show_sheet_tabs?
      @show_sheet_tabs != false
    end

    # Ratio between the sheet tabs bar and the horizontal scroll bar (0-1000).
    # @return [Integer, nil]
    # @api public
    attr_accessor :tab_ratio #: Integer?

    # Configures or queries workbook-level properties.
    #
    # When called with no arguments, returns the {Elements::WorkbookProperties} value object.
    # When called with keyword options, configures properties and returns nil.
    #
    # @param opts [Hash] Workbook properties map.
    # @return [Elements::WorkbookProperties, void]
    # @api public
    #: () -> Elements::WorkbookProperties
    #: (**untyped opts) -> void
    def workbook_properties(**opts)
      return Elements::WorkbookProperties.new(**@workbook_properties) if opts.empty?

      @workbook_properties.merge!(opts)
      nil
    end

    # Sets workbook-level properties.
    #
    # @param props [Hash, Elements::MetadataRecord]
    # @return [void]
    # @api public
    #: (untyped props) -> void
    def workbook_properties=(props)
      h = props.respond_to?(:to_h) ? props.to_h : (props || {})
      @workbook_properties = h.transform_keys(&:to_sym)
      nil
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

    # Sets or configures the workbook-wide default font.
    #
    # @param name [String, Hash, nil] Font name or font options Hash.
    # @param size [Float, Integer, nil] Font size in points.
    # @param opts [Hash] Additional font options (e.g. `bold:`, `italic:`, `family:`, `charset:`, `color:`).
    # @return [void]
    # @api public
    #: (?String | Hash[Symbol, untyped]? name, ?Float | Integer | nil size, **untyped opts) -> void
    def default_font(name = nil, size = nil, **opts)
      @style_writer.default_font(name, size, **opts)
    end

    # Creates an {Elements::Formula} object.
    # @api public
    #: (String expression, ?cached_value: (String | Numeric | bool | Elements::CellError | nil), ?type: (Symbol | String)?, ?ref: String?, ?shared_index: Integer?) -> Elements::Formula
    def formula(...)
      Xlsxrb.formula(...)
    end

    # Creates an array or dynamic array {Elements::Formula} object.
    # @api public
    #: (String expression, ?ref: String?, ?cached_value: (String | Numeric | bool | Elements::CellError | nil), **untyped attrs) -> Elements::Formula
    def array_formula(...)
      Xlsxrb.array_formula(...)
    end

    # Creates a shared formula {Elements::Formula} object.
    # @api public
    #: (?String? expression, ?shared_index: Integer?, ?ref: String?, ?cached_value: (String | Numeric | bool | Elements::CellError | nil)) -> Elements::Formula
    def shared_formula(...)
      Xlsxrb.shared_formula(...)
    end

    # Creates an {Elements::RichText} object.
    # @api public
    #: (*(Hash[Symbol, untyped] | Elements::RichTextRun | String | Array[untyped] | untyped) runs, ?text: String?, ?base_font: untyped, **untyped font_props) -> Elements::RichText
    def rich_text(...)
      Xlsxrb.rich_text(...)
    end

    # Creates or returns an {Elements::CellError} object.
    # @api public
    #: (String | Symbol | Elements::CellError code) -> Elements::CellError
    def cell_error(...)
      Xlsxrb.cell_error(...)
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

      # Returns the worksheet name.
      #
      # @return [String]
      # @api public
      #: () -> String
      def name
        @sheet_name
      end

      # Creates an {Elements::Formula} object.
      # @api public
      #: (String expression, ?cached_value: (String | Numeric | bool | Elements::CellError | nil), ?type: (Symbol | String)?, ?ref: String?, ?shared_index: Integer?) -> Elements::Formula
      def formula(...)
        Xlsxrb.formula(...)
      end

      # Creates an array or dynamic array {Elements::Formula} object.
      # @api public
      #: (String expression, ?ref: String?, ?cached_value: (String | Numeric | bool | Elements::CellError | nil), **untyped attrs) -> Elements::Formula
      def array_formula(...)
        Xlsxrb.array_formula(...)
      end

      # Creates a shared formula {Elements::Formula} object.
      # @api public
      #: (?String? expression, ?shared_index: Integer?, ?ref: String?, ?cached_value: (String | Numeric | bool | Elements::CellError | nil)) -> Elements::Formula
      def shared_formula(...)
        Xlsxrb.shared_formula(...)
      end

      # Creates an {Elements::RichText} object.
      # @api public
      #: (*(Hash[Symbol, untyped] | Elements::RichTextRun | String | Array[untyped] | untyped) runs, ?text: String?, ?base_font: untyped, **untyped font_props) -> Elements::RichText
      def rich_text(...)
        Xlsxrb.rich_text(...)
      end

      # Creates or returns an {Elements::CellError} object.
      # @api public
      #: (String | Symbol | Elements::CellError code) -> Elements::CellError
      def cell_error(...)
        Xlsxrb.cell_error(...)
      end

      # Returns whether formula escaping is enabled.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def escape_formulas?
        @writer.escape_formulas?
      end

      # Sets whether to treat strings starting with formula prefixes as plain text.
      #
      # @param val [Boolean]
      # @return [Boolean]
      # @api public
      #: (bool val) -> bool
      def escape_formulas=(val, ...)
        @writer.escape_formulas = val
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
      # @param range [String] The cell range (e.g. "A1:B2").
      # @return [void]
      # @api public
      #: (String range) -> void
      def merge(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.merge(...)
      end

      # Adds a drawing shape to the sheet.
      #
      # @example
      #   s.shape(preset: "ellipse", text: "Circle", from_column: 1, from_row: 1, to_column: 4, to_row: 5)
      #
      # @param preset [String] Preset shape type (e.g. "rect", "ellipse").
      # @param text [String, nil] Shape label text.
      # @param from_column [Integer, nil] Starting column index (0-based).
      # @param from_row [Integer] Starting row index (0-based).
      # @param to_column [Integer, nil] Ending column index (0-based).
      # @param to_row [Integer] Ending row index (0-based).
      # @param opts [Hash] Additional shape formatting options.
      # @return [void]
      # @api public
      #: (?preset: String, ?text: String?, ?from_column: Integer?, ?from_row: Integer, ?to_column: Integer?, ?to_row: Integer, **untyped opts) -> void
      def shape(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.shape(...)
      end

      # Appends a row of cells to the active sheet.
      #
      # @example Write an array of values
      #   s.row(["Name", "Age", "City"])
      #
      # @example Write with explicit column keys and styles
      #   s.row({ A: "Header", C: 100 }, cell_styles: { A: :bold })
      #
      # @param values [Array, Hash] The cell values (e.g. `[1, 2, 3]` or `{ A: 1, C: 3 }`).
      # @param cell_styles [String, Symbol, Array, Hash, nil] Style names or hashes to apply to cells.
      # @param row_style [Integer, Symbol, String, Hash, nil] Row style format or index.
      # @param height [Float, Integer, nil] The row height in points (0 - 409).
      # @param hidden [Boolean] Whether the row is hidden.
      # @param custom_height [Boolean] Whether to flag as custom height.
      # @param outline_level [Integer, nil] Grouping/outline level (0 - 7).
      # @param collapsed [Boolean] Whether the row is collapsed.
      # @return [void]
      # @api public
      #: (Array[untyped] | Hash[untyped, untyped] values, ?cell_styles: untyped, ?row_style: untyped, ?height: Float | Integer | nil, ?hidden: bool, ?custom_height: bool, ?collapsed: bool, ?outline_level: Integer | nil) -> void
      def row(values, cell_styles: nil, row_style: nil, height: nil, hidden: false, custom_height: false, collapsed: false, outline_level: nil)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.row(values, cell_styles: cell_styles, row_style: row_style, height: height, hidden: hidden, custom_height: custom_height, collapsed: collapsed, outline_level: outline_level)
      end

      # Appends a row to the worksheet and returns self for chaining.
      #
      # @example
      #   sheet << ["A", 1] << ["B", 2]
      #
      # @param values [Array, Hash] Row cell values.
      # @return [self]
      # @api public
      #: (Array[untyped] | Hash[untyped, untyped] values) -> self
      def <<(values)
        row(values)
        self
      end

      # Advances the row cursor by count without writing empty row elements.
      #
      # @example Skip multiple rows
      #   s.skip_rows(3)
      #
      # @param count [Integer] Number of rows to skip (must be >= 1).
      # @return [void]
      # @api public
      #: (?Integer count) -> void
      def skip_rows(count = 1)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.skip_rows(count)
      end

      # Configures column width and properties.
      #
      # @example Set column A width
      #   s.column(0, width: 25.0)
      #
      # @param index [Integer, String, Symbol, Range, Array] 0-based column index, letter ("A", :A), range, or array.
      # @param width [Float, Integer, nil] Column width in characters.
      # @param hidden [Boolean] Whether the column is hidden.
      # @param custom_width [Boolean] Whether to flag as custom width.
      # @param outline_level [Integer, nil] Grouping/outline level (0 - 7).
      # @param collapsed [Boolean] Whether the outline group is collapsed.
      # @param style [Integer, Symbol, String, Hash, nil] Column style format or index.
      # @param best_fit [Boolean] Whether the column automatically fits content.
      # @return [void]
      # @api public
      #: (Integer | String | Symbol | Range[Integer | String | Symbol] | Array[Integer | String | Symbol] index, ?width: Float | Integer | nil, ?hidden: bool, ?custom_width: bool, ?outline_level: Integer | nil, ?collapsed: bool, ?style: (Integer | Symbol | String | Hash[Symbol, untyped])?, ?best_fit: bool) -> void
      def column(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.column(...)
      end

      # Groups multiple columns together with outline hierarchy.
      #
      # @param index [Integer, String, Range, Array] Column index, letter, range, or array.
      # @param level [Integer] Grouping outline level (1 to 7).
      # @param collapsed [Boolean] Whether the group is collapsed.
      # @param hidden [Boolean] Whether the columns are hidden.
      # @param width [Float, Integer, nil] Optional column width.
      # @param style [Integer, Symbol, String, nil] Optional column style.
      # @return [void]
      # @api public
      #: (Integer | String | Range[Integer | String] | Array[Integer | String] index, ?level: Integer, ?collapsed: bool, ?hidden: bool, ?width: Float | Integer | nil, ?style: untyped) -> void
      def group_columns(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.group_columns(...)
      end

      # Groups multiple rows with outline hierarchy with a block.
      #
      # @param rows [Range, Array, Integer, nil] Unused in streaming write.
      # @param level [Integer] Grouping outline level (1 to 7).
      # @param collapsed [Boolean] Whether the group is collapsed.
      # @param hidden [Boolean] Whether the rows are hidden.
      # @yield Optional block executed with active group outline settings.
      # @return [void]
      # @api public
      #: (?untyped rows, ?level: Integer, ?collapsed: bool, ?hidden: bool) ?{ (self) -> void } -> void
      def group_rows(rows = nil, level: 1, collapsed: false, hidden: false)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        if block_given?
          @writer.group_rows(rows, level: level, collapsed: collapsed, hidden: hidden) do
            yield self
          end
        else
          @writer.group_rows(rows, level: level, collapsed: collapsed, hidden: hidden)
        end
      end

      # Configures or queries worksheet-level outline and grouping view properties.
      #
      # @param summary_below [Boolean, nil] Summary rows below detail rows.
      # @param summary_right [Boolean, nil] Summary columns to the right of detail columns.
      # @param show_outline_symbols [Boolean, nil] Outline symbols visible.
      # @param apply_styles [Boolean, nil] Automatic styles applied.
      # @param opts [Hash] Additional options.
      # @return [Hash<Symbol, untyped>, void]
      # @api public
      #: () -> Hash[Symbol, untyped]
      #: (?summary_below: bool?, ?summary_right: bool?, ?show_outline_symbols: bool?, ?apply_styles: bool?, **untyped opts) -> void
      def outline_properties(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.outline_properties(...)
      end

      # Sets worksheet-level outline and grouping view properties.
      #
      # @param opts [Hash] Outline properties.
      # @return [void]
      # @api public
      #: (Hash[Symbol, untyped] opts) -> void
      def outline_properties=(opts)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.outline_properties = opts
      end

      # Adds a chart to the active sheet.
      #
      # @example
      #   s.chart(:bar) do |chart_builder|
      #     chart_builder.title = "Quarterly Sales"
      #     chart_builder.series(values: "Sheet1!$B$2:$B$5", categories: "Sheet1!$A$2:$A$5", name: "Revenue")
      #   end
      #
      # @param type [Symbol, String, nil] The chart type (:bar, :col, :line, :pie, :scatter, :area, :doughnut, :radar).
      # @param options [Hash] Additional chart options.
      # @yield [chart_builder]
      # @yieldparam chart_builder [Xlsxrb::ChartBuilder]
      # @return [void]
      # @api public
      #: (?Symbol | String? type, **untyped options) ?{ (ChartBuilder) -> void } -> void
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
      #: (String | Integer cell, ?(String | Elements::Hyperlink)? url, ?display: String?, ?tooltip: String?, ?location: String?) -> void
      def hyperlink(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.hyperlink(...)
      end

      # Returns the auto-filter range on the active sheet, or nil if none.
      #
      # @return [String, nil]
      # @api public
      #: () -> String?
      def auto_filter
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.auto_filter
      end

      # Sets the auto-filter range on the active sheet.
      #
      # @param range [String, nil] The cell range (e.g. "A1:D10").
      # @return [String, nil]
      # @api public
      #: (String? range) -> String?
      def auto_filter=(range)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.auto_filter = range
      end

      # Sets filter criteria for a column in the auto-filter.
      #
      # @example Simple values filter
      #   s.filter_column(0, ["Active", "Pending"])
      #
      # @param col_id [Integer, String, Symbol] 0-based column index relative to auto-filter range or column letter.
      # @param filter [Hash] Values or filter specification hash.
      # @return [void]
      # @api public
      #: (Integer | String | Symbol col_id, (Hash[untyped, untyped] | Array[untyped]) filter) -> void
      def filter_column(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.filter_column(...)
      end

      # Configures or queries sort state on a range.
      #
      # When called with no arguments, returns the {Elements::SortState} value object.
      # When called with arguments, updates sort state and returns nil.
      #
      # @param ref [String, nil] The range to sort.
      # @param sort_conditions [Array<Hash>, nil] Sort conditions array.
      # @param opts [Hash] Additional sort options.
      # @return [Elements::SortState, nil, void]
      # @api public
      #: () -> Elements::SortState?
      #: (?String? ref, ?Array[Hash[untyped, untyped]]? sort_conditions, **untyped opts) -> void
      def sort_state(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.sort_state(...)
      end

      # Sets sort state configuration for the active sheet.
      #
      # @param opts [Elements::SortState, Hash, nil] Sort state options.
      # @return [void]
      # @api public
      #: (Elements::SortState | Hash[Symbol, untyped] | nil opts) -> void
      def sort_state=(opts)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.sort_state = opts
      end

      # Adds data validation rules to a range.
      #
      # @example Dropdown list validation
      #   s.data_validation("B2:B100", type: "list", formula1: '"High,Medium,Low"')
      #
      # @example Block syntax
      #   s.data_validation("B2:B100") do |dv|
      #     dv.type = :list
      #     dv.formula1 = '"High,Medium,Low"'
      #   end
      #
      # @param range [String, nil] The cell range (e.g. "B2:B10").
      # @param type [String, Symbol, nil] Validation type ("list", "whole", "decimal", "date", "time", "textLength", "custom").
      # @param opts [Hash] Validation options.
      # @yield [data_validation_builder] Optional configuration block.
      # @return [void]
      # @api public
      #: (?String? range, ?type: (String | Symbol)?, **untyped opts) ?{ (DataValidationBuilder) -> void } -> void
      def data_validation(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.data_validation(...)
      end

      # Adds conditional formatting to a range.
      #
      # @example Highlight values greater than 100
      #   s.conditional_format("A1:A10", type: "cellIs", operator: "greaterThan", formula: 100, style: :highlight)
      #
      # @example Hierarchical differential format (dxf)
      #   s.conditional_format("B2:B20", type: :cell_is, operator: "equal", formula: "0",
      #                        dxf: { font: { bold: false, color: { auto: true } }, fill: { pattern: "none", bg_color: "FFFF00" } })
      #
      # @param sqref [String, nil] The cell range (e.g. "A1:A10").
      # @param type [String, Symbol, nil] Rule type ("cellIs", "colorScale", "dataBar", "expression").
      # @param opts [Hash] Rule options.
      # @option opts [Hash] :dxf Hierarchical differential format specification.
      # @option opts [Integer] :format_id Pre-registered differential format index (dxfId).
      # @yield [conditional_format_builder] Optional block for configuration.
      # @return [void]
      # @api public
      #: (?String? sqref, ?type: String | Symbol, **untyped opts) ?{ (ConditionalFormatBuilder) -> void } -> void
      def conditional_format(sqref = nil, **opts, &)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.conditional_format(sqref, **opts, &)
      end

      # Adds a formatted Excel Table to the sheet.
      #
      # @example Standard keyword arguments
      #   s.table("A1:C10", columns: ["ID", "Name", "Total"], name: "SalesTable", style: "TableStyleMedium9")
      #
      # @example Block syntax
      #   s.table("A1:C10") do |t|
      #     t.columns = ["ID", "Name", "Total"]
      #     t.style = "TableStyleMedium9"
      #   end
      #
      # @param ref [String, nil] The cell range for the table (e.g. "A1:D10").
      # @param columns [Array<String>, Array<Hash>, nil] Column names or definitions.
      # @param name [String, nil] Table name.
      # @param display_name [String, nil] Display name.
      # @param style [String, nil] Table style name.
      # @param opts [Hash] Additional options.
      # @yield [table_builder] Optional configuration block.
      # @return [void]
      # @api public
      #: (?String? ref, ?columns: untyped, ?name: String?, ?display_name: String?, ?style: String?, **untyped opts) ?{ (TableBuilder) -> void } -> void
      def table(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.table(...)
      end

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

      # Adds a threaded comment to a cell.
      #
      # @param cell [String, Integer] The cell reference (e.g. "A1").
      # @param text [String] The comment text.
      # @param author [String] The author name.
      # @param opts [Hash] Options: timestamp, id, parent_id, done, user_id, provider_id.
      # @return [void]
      # @api public
      #: (String | Integer cell, String text, ?author: untyped, ?timestamp: untyped, ?id: String?, ?parent_id: String?, ?done: bool?, ?user_id: String?, ?provider_id: String?) -> void
      def threaded_comment(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.threaded_comment(...)
      end

      # Adds a Pivot Table to the sheet.
      #
      # @param source_ref [String] Source data range reference (e.g. "Sheet1!A1:D100").
      # @param row_fields [Array<String>] Field names for rows.
      # @param data_fields [Array<String>] Field names for data values.
      # @param column_fields [Array<String>, nil] Field names for columns.
      # @param ref [String, nil] Target top-left cell reference (default: "E1").
      # @param name [String, nil] Pivot table name.
      # @param field_names [Array<String>, nil] Override field names.
      # @param items [Array, nil] Items configuration.
      # @param opts [Hash] Additional options.
      # @return [void]
      # @api public
      #: (String source_ref, row_fields: untyped, data_fields: untyped, ?column_fields: untyped, ?ref: String?, ?name: String?, ?field_names: untyped, ?items: untyped, **untyped opts) -> void
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
      # Gets the sheet visibility state (:visible, :hidden, or :very_hidden).
      #
      # @return [Symbol]
      # @api public
      #: () -> Symbol
      def state
        @writer.state
      end

      # Sets the sheet visibility state (:visible, :hidden, or :very_hidden).
      #
      # @param new_state [Symbol, String] The new state.
      # @return [Symbol]
      # @api public
      #: (Symbol | String new_state) -> Symbol
      def state=(new_state)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.state = new_state
        @writer.state
      end

      # Gets the tab color of this worksheet.
      #
      # @return [String, Hash, nil]
      # @api public
      #: () -> (String | Hash[Symbol, untyped] | nil)
      def tab_color
        @writer.tab_color
      end

      # Sets the tab color of this worksheet.
      #
      # @param spec [String, Hash, Symbol, nil] Tab color hex string or hash options.
      # @return [String, Hash, nil]
      # @api public
      #: (String | Symbol | Hash[Symbol, untyped] | nil spec) -> (String | Hash[Symbol, untyped] | nil)
      def tab_color=(spec)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.tab_color = spec
      end

      # Configures or queries sheet properties.
      #
      # @param opts [Hash] Keyword options for sheet properties.
      # @return [Elements::SheetProperties, void]
      # @api public
      #: () -> Elements::SheetProperties
      #: (**untyped opts) -> void
      def sheet_properties(**opts)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.sheet_properties(**opts)
      end

      # Replaces worksheet properties.
      # @param opts [Hash, Elements::SheetProperties]
      # @return [void]
      # @api public
      #: (Hash[Symbol, untyped] | Elements::SheetProperties opts) -> void
      def sheet_properties=(opts)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.sheet_properties = opts
      end

      # Adds a defined name (named range or formula) scoped to this sheet or workbook.
      # Bare cell references and coordinate keywords are automatically formatted with sheet quotes and absolute $ references.
      #
      # @param name [String] Defined name identifier.
      # @param value [Object, nil] Value or formula expression.
      # @param hidden [Boolean] Whether the defined name is hidden.
      # @param columns [Integer, String, nil] Column index or letter for coordinate-based reference.
      # @param rows [Integer, Range, nil] Row or row range (0-based) for coordinate-based reference.
      # @param range [String, nil] Cell range reference.
      # @param sheet [String, nil] Reference sheet name (defaults to current sheet).
      # @param scope [String, Symbol, nil] Scope identifier (:workbook or sheet name).
      # @return [void]
      # @api public
      #: (String name, ?untyped value, ?hidden: bool, ?columns: untyped, ?rows: untyped, ?range: untyped, ?sheet: String?, ?scope: (String | Symbol)?) -> void
      def defined_name(name, value = nil, hidden: false, columns: nil, rows: nil, range: nil, sheet: nil, scope: nil)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        ref_sheet = sheet || @sheet_name
        local_sheet = if scope
                        scope == :workbook ? nil : scope.to_s
                      end
        @writer.defined_name(name, value, sheet: ref_sheet, scope: local_sheet, hidden: hidden, columns: columns, rows: rows, range: range)
      end

      # Configures or queries freeze panes (<pane state="frozen">).
      #
      # When called with no arguments, returns the {Elements::FreezePane} value object or nil.
      # When called with arguments, configures the freeze pane and returns nil.
      #
      # @example Freeze top row
      #   s.freeze_pane(row: 1)
      #
      # @example Freeze first column and top 2 rows
      #   s.freeze_pane(row: 2, column: 1)
      #
      # @param row [Integer, nil] Number of rows to freeze.
      # @param column [Integer, String, nil] Number of columns to freeze.
      # @return [Elements::FreezePane, nil, void]
      # @api public
      #: () -> Elements::FreezePane?
      #: (?row: Integer | nil, ?column: (Integer | String)?, ?top_left_cell: String?, ?active_pane: (String | Symbol)?) -> void
      def freeze_pane(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.freeze_pane(...)
      end

      # Sets or clears freeze pane configuration.
      #
      # @param opts [Elements::FreezePane, Hash, nil] Freeze pane options.
      # @return [void]
      # @api public
      #: (Elements::FreezePane | Hash[Symbol, untyped] | nil opts) -> void
      def freeze_pane=(opts)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.freeze_pane = opts
      end

      # Returns the print area range for this worksheet, or nil if none.
      #
      # @return [String, nil]
      # @api public
      #: () -> String?
      def print_area
        @writer.print_area(sheet: @sheet_name)
      end

      # Sets the print area for this worksheet.
      #
      # @param range [String, nil] Cell range (e.g. "A1:G50").
      # @return [String, nil]
      # @api public
      #: (String? range) -> String?
      def print_area=(range)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.send(:set_print_area, range, sheet: @sheet_name)
      end

      # Sets or gets repeating print titles (rows and/or columns) for pagination.
      #
      # When called with no arguments, returns the {Elements::PrintTitles} value object or nil.
      # When called with arguments, updates print titles and returns nil.
      #
      # @param rows [String, Integer, Range, nil] Repeating row range (e.g. "1:2").
      # @param columns [String, Integer, Symbol, Range, nil] Repeating column range (e.g. "A:B").
      # @return [Elements::PrintTitles, nil, void]
      # @api public
      #: () -> Elements::PrintTitles?
      #: (?rows: (String | Integer | Range[Integer])?, ?columns: (String | Integer | Symbol | Range[untyped])?) -> void
      def print_titles(rows: nil, columns: nil)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        if rows.nil? && columns.nil?
          @writer.print_titles(sheet: @sheet_name)
        else
          @writer.print_titles(rows: rows, columns: columns, sheet: @sheet_name)
        end
      end

      # Sets repeating print titles for this worksheet.
      #
      # @param opts [Elements::PrintTitles, Hash, nil] Options with :rows and/or :columns.
      # @return [Elements::PrintTitles, Hash, nil]
      # @api public
      #: (Elements::PrintTitles | Hash[Symbol, untyped] | nil opts) -> (Elements::PrintTitles | Hash[Symbol, untyped] | nil)
      def print_titles=(opts)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        if opts.nil?
          @writer.print_titles = nil
        else
          h = opts.respond_to?(:to_h) ? opts.to_h : opts
          @writer.print_titles(**h, sheet: @sheet_name)
        end
        opts
      end

      # Sets repeating title rows to appear at the top of each printed page.
      #
      # @param first [Integer, String] First row index (0-based) or range string (e.g. "1:2").
      # @param last [Integer, String, nil] Optional last row index (0-based).
      # @return [void]
      # @api public
      #: (Integer | String first, ?(Integer | String)? last) -> void
      def repeat_rows(first, last = nil)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.repeat_rows(first, last, sheet: @sheet_name)
      end

      # Sets repeating title columns to appear at the left of each printed page.
      #
      # @param first [Integer, String] First column index (0-based) or column name/range.
      # @param last [Integer, String, nil] Optional last column index/name.
      # @return [void]
      # @api public
      #: (Integer | String first, ?(Integer | String)? last) -> void
      def repeat_columns(first, last = nil)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.repeat_columns(first, last, sheet: @sheet_name)
      end
      # simplecov:enable

      # Splits sheet view into panes or queries split pane settings (<pane>).
      #
      # When called with no arguments, returns the {Elements::SplitPane} value object or nil.
      # When called with arguments, configures the split pane and returns nil.
      #
      # @param x_split [Integer, nil] Horizontal split position.
      # @param y_split [Integer, nil] Vertical split position.
      # @param top_left_cell [String, nil] Top-left visible cell in bottom-right pane.
      # @param active_pane [String, Symbol, nil] Active pane identifier.
      # @param state [String, Symbol, nil] Split state.
      # @return [Elements::SplitPane, nil, void]
      # @api public
      #: () -> Elements::SplitPane?
      #: (?x_split: Integer | nil, ?y_split: Integer | nil, ?top_left_cell: String?, ?active_pane: (String | Symbol)?) -> void
      def split_pane(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.split_pane(...)
      end

      # Sets or clears split pane configuration.
      #
      # @param opts [Elements::SplitPane, Hash, nil] Split pane options.
      # @return [void]
      # @api public
      #: (Elements::SplitPane | Hash[Symbol, untyped] | nil opts) -> void
      def split_pane=(opts)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.split_pane = opts
      end

      # Sets or queries the active/selected cell on the sheet (<selection>).
      #
      # When called with no arguments, returns the selection ({Elements::Selection}, Array, or nil).
      # When called with arguments, updates the selection and returns nil.
      #
      # @param active_cell [String, nil] Cell reference (e.g. "A1") or range (e.g. "A1:D10").
      # @param sqref [String, nil] Selection range.
      # @param pane [String, Symbol, nil] Pane identifier.
      # @return [Elements::Selection, Array<Elements::Selection>, nil, void]
      # @api public
      #: () -> (Elements::Selection | Array[Elements::Selection])?
      #: (?String? active_cell, ?sqref: String?, ?pane: (String | Symbol)?) -> void
      def selection(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.selection(...)
      end

      # Sets or clears the selection configuration.
      #
      # @param opts [Elements::Selection, Array<Elements::Selection, Hash>, Hash, nil]
      # @return [void]
      # @api public
      #: (Elements::Selection | Array[Elements::Selection | Hash[Symbol, untyped]] | Hash[Symbol, untyped] | nil opts) -> void
      def selection=(opts)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.selection = opts
      end

      # Configures or queries page margins for printing (<pageMargins>).
      #
      # When called with no arguments, returns the {Elements::PageMargins} value object or nil.
      # When called with arguments, configures page margins and returns nil.
      #
      # @param left [Float, nil] Left margin in inches.
      # @param right [Float, nil] Right margin in inches.
      # @param top [Float, nil] Top margin in inches.
      # @param bottom [Float, nil] Bottom margin in inches.
      # @param header [Float, nil] Header margin in inches.
      # @param footer [Float, nil] Footer margin in inches.
      # @return [Elements::PageMargins, nil, void]
      # @api public
      #: () -> Elements::PageMargins?
      #: (?left: Float | nil, ?right: Float | nil, ?top: Float | nil, ?bottom: Float | nil, ?header: Float | nil, ?footer: Float | nil) -> void
      def page_margins(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.page_margins(...)
      end

      # Sets or clears page margins configuration.
      #
      # @param opts [Hash, Elements::PageMargins, nil]
      # @return [void]
      # @api public
      #: (Hash[Symbol, untyped] | Elements::PageMargins | nil opts) -> void
      def page_margins=(opts)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.page_margins = opts
      end

      # Configures or queries page orientation, paper size, and print setup (<pageSetup>).
      #
      # When called with no arguments, returns the {Elements::PageSetup} value object.
      # When called with keyword arguments, updates page setup and returns nil.
      #
      # @param orientation [String, Symbol, nil] "portrait" or "landscape" (or :portrait, :landscape).
      # @param paper_size [Integer, nil] Paper size index (e.g. 9 for A4, 1 for Letter).
      # @param opts [Hash] Additional options (scale, fit_to_width, fit_to_height).
      # @return [Elements::PageSetup, void]
      # @api public
      #: () -> Elements::PageSetup
      #: (**untyped opts) -> void
      def page_setup(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.page_setup(...)
      end

      # Sets page setup configuration.
      #
      # @param opts [Hash, Elements::PageSetup]
      # @return [void]
      # @api public
      #: (Hash[Symbol, untyped] | Elements::PageSetup opts) -> void
      def page_setup=(opts)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.page_setup = opts
      end

      # Configures or queries header and footer text for printing (<headerFooter>).
      #
      # When called with no arguments, returns the {Elements::HeaderFooter} value object.
      # When called with keyword arguments, updates headers and footers and returns nil.
      #
      # @param opts [Hash] Header and footer specifications.
      # @return [Elements::HeaderFooter, void]
      # @api public
      #: () -> Elements::HeaderFooter
      #: (**untyped opts) -> void
      def header_footer(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.header_footer(...)
      end

      # Sets header and footer configuration.
      #
      # @param opts [Hash, Elements::HeaderFooter]
      # @return [void]
      # @api public
      #: (Hash[Symbol, untyped] | Elements::HeaderFooter opts) -> void
      def header_footer=(opts)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.header_footer = opts
      end

      # Configures or gets print options (e.g. gridlines, headings).
      # Configures or queries print options.
      #
      # @param opts [Hash] Additional options.
      # @return [Elements::PrintOptions, void]
      # @api public
      #: () -> Elements::PrintOptions
      #: (**untyped opts) -> void
      def print_options(**opts)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.print_options(**opts)
      end

      # Replaces print options.
      # @param opts [Hash, Elements::PrintOptions]
      # @return [void]
      # @api public
      #: (Hash[Symbol, untyped] | Elements::PrintOptions opts) -> void
      def print_options=(opts)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.print_options = opts
      end

      # Sets or queries sheet-level protection with optional password hashing (<sheetProtection>).
      #
      # When called with no arguments, returns the {Elements::SheetProtection} value object or nil.
      # When called with keyword arguments, sets sheet protection and returns nil.
      #
      # @param opts [Hash] Protection options.
      # @return [Elements::SheetProtection, nil, void]
      # @api public
      #: () -> Elements::SheetProtection?
      #: (**untyped opts) -> void
      def sheet_protection(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.sheet_protection(...)
      end

      # Sets or clears sheet protection configuration.
      #
      # @param opts [Elements::SheetProtection, Hash, Boolean, nil]
      # @return [void]
      # @api public
      #: (Elements::SheetProtection | Hash[Symbol, untyped] | bool | nil opts) -> void
      def sheet_protection=(opts)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.sheet_protection = opts
      end

      # Returns column configurations defined for this worksheet.
      #
      # @return [Array<Elements::Column>]
      # @api public
      #: () -> Array[Elements::Column]
      def columns
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.columns
      end

      # Inserts an image into the sheet.
      #
      # @param file_data [String] Binary image data or file content.
      # @param ext [String] Image extension ("png", "jpeg", etc.).
      # @param from_column [Integer, nil] Starting column index (0-based).
      # @param from_row [Integer] Starting row index (0-based).
      # @param to_column [Integer, nil] Ending column index (0-based).
      # @param to_row [Integer] Ending row index (0-based).
      # @param opts [Hash] Additional anchor and sizing options.
      # @return [void]
      # @api public
      #: (String file_data, ?ext: String, ?from_column: Integer?, ?from_row: Integer, ?to_column: Integer?, ?to_row: Integer, **untyped opts) -> void
      def image(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.image(...)
      end

      # Gets the worksheet background image (<picture>).
      #
      # @return [Elements::BackgroundImage, nil]
      # @api public
      #: () -> Elements::BackgroundImage?
      def background_image
        @writer.background_image
      end

      # Sets the worksheet background image (<picture>).
      #
      # @param val [String, Hash, Elements::Image, nil] Image binary data, file path, or image options hash.
      # @return [untyped]
      # @api public
      #: (untyped val) -> untyped
      def background_image=(val)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.background_image = val
      end

      # Configures or gets sheet view settings (zoom scale, grid lines visibility).
      #
      # @param opts [Hash] Additional options.
      # @return [Elements::SheetView, void]
      # @api public
      #: () -> Elements::SheetView
      #: (**untyped opts) -> void
      def sheet_view(**opts)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.sheet_view(**opts)
      end

      # Replaces sheet view properties.
      # @param opts [Hash, Elements::SheetView]
      # @return [void]
      # @api public
      #: (Hash[Symbol, untyped] | Elements::SheetView opts) -> void
      def sheet_view=(opts)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.sheet_view = opts
      end

      # Configures or gets sheet format properties (e.g. default_row_height, default_column_width).
      #
      # @param opts [Hash] Keyword arguments for sheet format properties.
      # @return [Elements::SheetFormat, void]
      # @api public
      #: () -> Elements::SheetFormat
      #: (**untyped opts) -> void
      def sheet_format(**opts)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.sheet_format(**opts)
      end

      # Replaces sheet format properties.
      # @param opts [Hash, Elements::SheetFormat]
      # @return [void]
      # @api public
      #: (Hash[Symbol, untyped] | Elements::SheetFormat opts) -> void
      def sheet_format=(opts)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.sheet_format = opts
      end

      # simplecov:disable
      # Edge case / untested delegation block
      # Adds a horizontal page break after the given row index.
      #
      # @param row_index [Integer] 0-based row index.
      # @return [void]
      # @api public
      #: (Integer row_index) -> void
      def row_break(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.row_break(...)
      end
      # simplecov:enable

      # simplecov:disable
      # Edge case / untested delegation block
      # Adds a vertical page break after the given column.
      #
      # @param col [Integer, String, Symbol] 0-based column index or column letter.
      # @return [void]
      # @api public
      #: (Integer | String | Symbol col) -> void
      def column_break(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.column_break(...)
      end
      # simplecov:enable

      # Adds a protected range to the sheet.
      #
      # @param name [String] Name of the protected range.
      # @param sqref [String] Cell range (e.g. "A1:B10").
      # @param password [String, nil] Optional password or hash.
      # @param opts [Hash] Additional options.
      # @return [void]
      # @api public
      #: (String name, String sqref, ?password: String?, **untyped opts) -> void
      def protected_range(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.protected_range(...)
      end

      # Adds a cell watch to the sheet.
      #
      # @param ref [String] Cell reference (e.g. "A1").
      # @return [void]
      # @api public
      #: (String ref) -> void
      def cell_watch(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.cell_watch(...)
      end

      # Adds an ignored error entry to the sheet.
      #
      # @param sqref [String] Cell range reference (e.g. "A1" or "A1:B10").
      # @param opts [Hash] Error check flags to ignore.
      # @return [void]
      # @api public
      #: (String sqref, ?eval_error: bool, ?two_digit_text_year: bool, ?number_stored_as_text: bool, ?formula: bool, ?formula_range: bool, ?unlocked_formula: bool, ?empty_cell_reference: bool, ?list_data_validation: bool, ?calculated_column: bool) -> void
      def ignored_error(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.ignored_error(...)
      end

      # Returns the ignored error entries configured for this sheet.
      #
      # @return [Array<Elements::IgnoredError>]
      # @api public
      #: () -> Array[Elements::IgnoredError]
      def ignored_errors
        @writer.ignored_errors
      end

      # Gets phonetic (furigana) guide properties for the sheet.
      #
      # @return [Elements::PhoneticProperties, nil]
      # @api public
      #: () -> Elements::PhoneticProperties?
      def phonetic_properties
        @writer.phonetic_properties
      end

      # Sets phonetic (furigana) guide properties for the sheet.
      #
      # @param spec [Hash, Integer, nil] Options hash or font ID integer.
      # @return [Hash[Symbol, untyped], Integer, nil]
      # @api public
      #: (Hash[Symbol, untyped] | Integer | nil spec) -> (Hash[Symbol, untyped] | Integer | nil)
      def phonetic_properties=(spec)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.phonetic_properties = spec
      end

      # Gets the internal code name of the worksheet.
      #
      # @return [String, nil]
      # @api public
      #: () -> String?
      def code_name
        @writer.code_name
      end

      # Sets the internal code name of the worksheet.
      #
      # @param val [String, Symbol, nil] The code name.
      # @return [String, nil]
      # @api public
      #: (String | Symbol | nil val) -> String?
      def code_name=(val)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.code_name = val
      end

      # Returns whether page setup fits content to page width/height.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def fit_to_page?
        @writer.fit_to_page?
      end

      # Sets whether page setup fits content to page width/height.
      #
      # @param enabled [Boolean, nil] Whether to fit to page.
      # @return [Boolean]
      # @api public
      #: (bool? enabled) -> bool??
      def fit_to_page=(enabled)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.fit_to_page = enabled
      end

      # Gets worksheet calculation properties (<sheetCalcPr>).
      #
      # @return [Elements::SheetCalcProperties, nil]
      # @api public
      #: () -> Elements::SheetCalcProperties?
      def sheet_calc_properties
        @writer.sheet_calc_properties
      end

      # Sets worksheet calculation properties (<sheetCalcPr>).
      #
      # @param spec [Hash, nil] Calculation properties hash.
      # @return [Hash, nil]
      # @api public
      #: (Hash[Symbol, untyped]? spec) -> Hash[Symbol, untyped]?
      def sheet_calc_properties=(spec)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.sheet_calc_properties = spec
      end
    end

    # Returns the list of worksheet names created in this workbook so far.
    #
    # @return [Array<String>] Worksheet names.
    # @api public
    #: () -> Array[String]
    def sheet_names
      names = @sheets.map { |s| s.respond_to?(:name) ? s.name : s[:name].to_s }
      names << @current_sheet if @current_sheet
      names
    end

    # Adds a new sheet to the workbook and starts streaming rows into it.
    #
    # @param name [String, nil] Sheet name (max 31 characters).
    # @param state [Symbol, String, nil] Sheet visibility state (:visible, :hidden, :very_hidden).
    # @param tab_color [String, Hash, Symbol, nil] Tab color hex string (e.g. "FF0000") or hash options.
    # @param code_name [String, Symbol, nil] Internal VBA code name of the sheet.
    # @param opts [Hash] Sheet-level configuration.
    # @yield [sheet_proxy]
    # @yieldparam sheet_proxy [WorksheetProxy] The streaming sheet proxy.
    # @return [String] The sheet name.
    # @api public
    #: (?String? name, ?state: Symbol | String | nil, ?tab_color: String | Hash[Symbol, untyped] | Symbol | nil, ?code_name: String | Symbol | nil, **untyped opts) ?{ (WorksheetProxy) -> void } -> String
    def sheet(name = nil, state: nil, tab_color: nil, code_name: nil, **opts)
      sheet_count = @sheets.size + (@current_sheet ? 1 : 0)
      name ||= "Sheet#{sheet_count + 1}"
      raise ArgumentError, "Sheet name '#{name}' must be <= 31 characters (Excel limitation)" if @strict_excel_mode && name.length > 31
      raise ArgumentError, "Sheet name '#{name}' contains invalid characters (Excel limitation)" if name.match?(Elements::Worksheet.forbidden_name_chars)
      raise ArgumentError, "Sheet name '#{name}' cannot start or end with an apostrophe (Excel limitation)" if @strict_excel_mode && (name.start_with?("'") || name.end_with?("'"))
      raise ArgumentError, "Sheet name '#{name}' is reserved by Excel (Excel limitation)" if @strict_excel_mode && name.casecmp?("History")

      existing_names = @sheets.map { |s| s.respond_to?(:name) ? s.name.downcase : s[:name].to_s.downcase }
      existing_names << @current_sheet.downcase if @current_sheet
      raise ArgumentError, "Sheet name '#{name}' is already used. Excel requires unique sheet names." if @strict_excel_mode && existing_names.include?(name.downcase)

      active = opts.delete(:active)
      state = opts.delete(:state) if opts.key?(:state)
      hidden = opts.delete(:hidden)
      code_name = opts.delete(:code_name) if opts.key?(:code_name)
      tab_color = opts.delete(:tab_color) if opts.key?(:tab_color)
      state ||= :hidden if hidden
      state ||= :visible
      state = state.to_sym
      raise ArgumentError, "Invalid sheet state ':#{state}'. Allowed states are :visible, :hidden, :very_hidden" unless VALID_SHEET_STATES.include?(state)

      internal_sheet_setup(name, state: state)
      if active
        @active_sheet = name
        @current_sheet_view[:tab_selected] = true
      end
      self.tab_color = tab_color if tab_color
      self.code_name = code_name if code_name
      sheet_properties(**opts) unless opts.empty?

      yield WorksheetProxy.new(self, @current_sheet) if block_given?
      @current_sheet
    end

    # Returns a WorksheetProxy for the current or named sheet.
    #
    # @param name_or_index [String, Symbol, Integer, nil] Sheet name or 0-based sheet index.
    # @yield [proxy]
    # @yieldparam proxy [WorksheetProxy]
    # @return [WorksheetProxy]
    # @api public
    #: (?String | Symbol | Integer | nil name_or_index) ?{ (WorksheetProxy) -> void } -> WorksheetProxy
    def [](name_or_index = nil)
      target_name = if name_or_index.nil?
                      @current_sheet || sheet
                    elsif name_or_index.is_a?(Integer)
                      all_names = sheet_names
                      all_names[name_or_index] || raise(ArgumentError, "Sheet index #{name_or_index} out of bounds")
                    else
                      name_or_index.to_s
                    end
      proxy = WorksheetProxy.new(self, target_name)
      yield proxy if block_given?
      proxy
    end

    # Internal: Start or switch to a named sheet (internal helper).
    # @api private
    #: (?String? name, ?state: Symbol) ?{ (WorksheetProxy) -> void } -> (WorksheetProxy | nil)
    def internal_sheet_setup(name = nil, state: :visible)
      flush_current_sheet
      name ||= "Sheet#{@sheets.size + 1}"
      @current_sheet = name
      @current_sheet_state = state
      @current_sheet_index = @sheets.size + 1
      @current_row_index = 0
      @sheet_entry_started = false
      @current_row_buffer = String.new(capacity: 65_536)
      @current_sheet_io = StringIO.new(@current_row_buffer)
      date1904 = @workbook_properties&.[](:date1904) ? true : false
      @current_row_writer = Ooxml::WorksheetWriter.new(@current_sheet_io, date1904: date1904, escape_formulas: @escape_formulas)
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
      @current_threaded_comments = []
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
      @current_sheet_format = {}
      @current_sheet_view = {}
      @current_row_breaks = []
      @current_col_breaks = []
      @current_protected_ranges = []
      @current_cell_watches = []
      @current_ignored_errors = []
      @current_phonetic_properties = nil
      @current_sheet_calc_pr = nil
      @current_background_image = nil
      @current_print_area = nil
      @current_print_titles = nil
      @current_row_outline_level = nil
      @current_row_collapsed = false
      @current_row_hidden = false
      @current_cells = {}

      return unless block_given?

      # simplecov:disable
      # Edge case / untested delegation block
      yield self
      flush_current_sheet
      # simplecov:enable
    end

    # @api private
    #: () -> void
    def start_sheet_entry
      return if @sheet_entry_started

      @sheet_entry_started = true
      @zip.start_entry("xl/worksheets/sheet#{@current_sheet_index}.xml")
      zip_io = Ooxml::WorkbookWriter::ZipEntryIO.new(@zip)

      row_buf = @current_row_writer.instance_variable_get(:@row_buffer)
      if row_buf && !row_buf.empty?
        @current_sheet_io.write(row_buf)
        row_buf.clear
      end

      if @current_columns && !@current_columns.empty? && @style_name_to_id
        @current_columns.each do |col|
          s = col[:style] || col[:style_index]
          if s.is_a?(Symbol) || s.is_a?(String)
            resolved = @style_name_to_id[s] || @style_name_to_id[s.to_s]
            col[:style] = resolved if resolved
          end
        end
      end

      date1904 = @workbook_properties&.[](:date1904) ? true : false
      @current_row_writer = Ooxml::WorksheetWriter.new(zip_io, date1904: date1904, escape_formulas: @escape_formulas)
      @current_row_writer.start(
        columns: @current_columns,
        sheet_properties: @current_sheet_properties.empty? ? nil : @current_sheet_properties,
        freeze_pane: @current_freeze_pane,
        split_pane: @current_split_pane,
        selection: @current_selection,
        sheet_view: @current_sheet_view.empty? ? nil : @current_sheet_view,
        sheet_format: @current_sheet_format.empty? ? nil : @current_sheet_format
      )
      @zip.write_data(@current_row_buffer) unless @current_row_buffer.empty?
      @current_row_buffer.clear
    end

    # Appends a row of values to the active sheet.
    #
    # @param values [Array<Object>, Hash] The cell values.
    # @param cell_styles [String, Symbol, Array, Hash, nil] Style names or inline style definitions to apply to cells.
    # @param row_style [Integer, Symbol, String, Hash, nil] Row style format or index.
    # @param height [Float, Integer, nil] The row height in points (0 - 409).
    # @param hidden [Boolean] Whether the row is hidden.
    # @param custom_height [Boolean] Whether custom row height is set.
    # @param outline_level [Integer, nil] Grouping/outline level (0 - 7).
    # @param collapsed [Boolean] Whether the row is collapsed.
    # @return [void]
    # @api public
    #: (Array[untyped] | Hash[untyped, untyped] values, ?cell_styles: untyped, ?row_style: untyped, ?height: Float | Integer | nil, ?hidden: bool, ?custom_height: bool, ?collapsed: bool, ?outline_level: Integer | nil) -> void
    def row(values, cell_styles: nil, row_style: nil, height: nil, hidden: false, custom_height: false, collapsed: false, outline_level: nil)
      sheet if @current_sheet.nil?

      row_index = @current_row_index
      DslHelpers.validate_row_bounds!(row_index, height, strict_excel_mode: @strict_excel_mode)
      @current_row_index += 1

      values = DslHelpers.normalize_row_values(values)
      cell_styles = DslHelpers.normalize_row_styles(cell_styles)
      cell_styles = with_date_style_variants(values, cell_styles) if cell_styles

      if @current_charts && !@current_charts.empty?
        @current_cells ||= {}
        row_num = row_index + 1
        values.each_with_index do |val, col_idx|
          next if val.nil?

          addr = "#{Elements::Cell._column_letter(col_idx)}#{row_num}"
          @current_cells[addr] = val
        end
      end

      raise ArgumentError, "Row contains #{values.length} columns, exceeding Excel limit of 16_384 columns" if @strict_excel_mode && values.length > 16_384

      row_style_id = if row_style.is_a?(Hash)
                       inline_name = "__inline_row_#{row_style.hash}"
                       style(inline_name, **row_style) unless @style_name_to_id.key?(inline_name)
                       @style_name_to_id[inline_name]
                     else
                       row_style
                     end

      effective_outline = outline_level || @current_row_outline_level
      effective_collapsed = collapsed || @current_row_collapsed
      effective_hidden = hidden || @current_row_hidden

      attrs = nil
      if height || effective_hidden || effective_outline || effective_collapsed || row_style_id
        attrs = {}
        attrs[:height] = height if height
        attrs[:hidden] = true if effective_hidden
        attrs[:custom_height] = custom_height || !height.nil?
        attrs[:collapsed] = true if effective_collapsed
        attrs[:outline_level] = effective_outline if effective_outline
        attrs[:style_index] = row_style_id if row_style_id
      end

      @current_row_writer.write_row_values(row_index, values, styles: cell_styles, style_map: @style_name_to_id, sst: @sst, sst_index: @sst_index, attrs: attrs)
      start_sheet_entry if !@sheet_entry_started && @current_row_buffer.bytesize >= 65_536
    end

    # Appends a row to the worksheet and returns self for chaining.
    #
    # @example
    #   writer << ["A", 1] << ["B", 2]
    #
    # @param values [Array, Hash] Row cell values.
    # @return [self]
    # @api public
    #: (Array[untyped] | Hash[untyped, untyped] values) -> self
    def <<(values)
      row(values)
      self
    end

    # Advances the row cursor by count without writing empty row elements.
    #
    # @example Skip multiple rows
    #   w.skip_rows(3)
    #
    # @param count [Integer] Number of rows to skip (must be >= 1).
    # @return [void]
    # @api public
    #: (?Integer count) -> void
    def skip_rows(count = 1)
      raise ArgumentError, "count must be positive: #{count.inspect}" if !count.is_a?(Integer) || count < 1

      sheet if @current_sheet.nil?
      @current_row_index += count
    end

    # Sets column formatting and properties for one or multiple columns.
    #
    # @param index [Integer, String, Symbol, Range, Array] Column index (0-based), letter ("A", :A), range, or array.
    # @param width [Float, Integer, nil] Column width in character units (0 - 255).
    # @param hidden [Boolean] Whether the column is hidden.
    # @param custom_width [Boolean] Whether custom width flag is set.
    # @param outline_level [Integer, nil] Grouping/outline level.
    # @param collapsed [Boolean] Whether the column is collapsed.
    # @param style [Integer, Symbol, String, Hash, nil] Column style format or index.
    # @param best_fit [Boolean] Whether the column automatically fits content.
    # @return [void]
    # @api public
    #: (Integer | String | Symbol | Range[Integer | String | Symbol] | Array[Integer | String | Symbol] index, ?width: Float | Integer | nil, ?hidden: bool, ?custom_width: bool, ?outline_level: Integer | nil, ?collapsed: bool, ?style: (Integer | Symbol | String | Hash[Symbol, untyped])?, ?best_fit: bool) -> void
    def column(index, width: nil, hidden: false, custom_width: false, outline_level: nil, collapsed: false, style: nil, best_fit: false)
      raise ArgumentError, "Column width #{width} must be between 0 and 255 characters (Excel limitation)" if @strict_excel_mode && width && (width.negative? || width > 255)

      sheet if @current_sheet.nil?

      style_id = if style.is_a?(Hash)
                   inline_name = "__inline_col_#{style.hash}"
                   style(inline_name, **style) unless @style_name_to_id.key?(inline_name)
                   @style_name_to_id[inline_name]
                 elsif @style_name_to_id && style
                   @style_name_to_id[style] || @style_name_to_id[style.to_s] || style
                 else
                   style
                 end
      DslHelpers.normalize_column_indices(index).each do |idx|
        existing = @current_columns.find { |c| c[:index] == idx }
        if existing
          existing[:width] = width unless width.nil?
          existing[:custom_width] = true if custom_width || !width.nil?
          existing[:hidden] = hidden if hidden
          existing[:outline_level] = outline_level unless outline_level.nil?
          existing[:collapsed] = true if collapsed
          existing[:best_fit] = true if best_fit
          if style_id
            existing[:style] = style_id
            @current_row_writer&.set_column_style(idx, style_id)
          end
        else
          entry = { index: idx, width: width, hidden: hidden, custom_width: custom_width || !width.nil?, outline_level: outline_level }
          entry[:collapsed] = true if collapsed
          entry[:best_fit] = true if best_fit
          if style_id
            entry[:style] = style_id
            @current_row_writer&.set_column_style(idx, style_id)
          end
          @current_columns << entry
        end
      end
      @current_columns.sort_by! { |c| c[:index] }
      nil
    end

    # Returns column configurations defined for the active worksheet.
    #
    # @return [Array<Elements::Column>]
    # @api public
    #: () -> Array[Elements::Column]
    def columns
      ensure_sheet_started!
      @current_columns.map do |c|
        if c.is_a?(Elements::Column)
          c
        else
          idx = c[:index]
          style_idx = c[:style] || c[:style_index]
          unmapped = c.except(:index, :width, :hidden, :custom_width, :outline_level, :style, :style_index, :collapsed, :best_fit)
          Elements::Column.new(
            index: idx,
            width: c[:width],
            hidden: c[:hidden] ? true : false,
            custom_width: c[:custom_width] ? true : false,
            outline_level: c[:outline_level],
            style_index: style_idx,
            collapsed: c[:collapsed] ? true : false,
            best_fit: c[:best_fit] ? true : false,
            unmapped_data: unmapped
          )
        end
      end
    end

    # Groups multiple columns together with outline hierarchy.
    #
    # @example Group columns B through D
    #   sheet.group_columns("B".."D", level: 1)
    #
    # @example Group columns 1 through 3 and collapse
    #   sheet.group_columns(1..3, level: 1, collapsed: true)
    #
    # @param index [Integer, String, Range, Array] Column index (0-based), letter, range, or array.
    # @param level [Integer] Grouping outline level (1 to 7).
    # @param collapsed [Boolean] Whether the group is collapsed.
    # @param hidden [Boolean] Whether the columns are hidden.
    # @param width [Float, Integer, nil] Optional column width.
    # @param style [Integer, Symbol, String, nil] Optional column style.
    # @return [void]
    # @api public
    #: (Integer | String | Range[Integer | String] | Array[Integer | String] index, ?level: Integer, ?collapsed: bool, ?hidden: bool, ?width: Float | Integer | nil, ?style: untyped) -> void
    def group_columns(index, level: 1, collapsed: false, hidden: false, width: nil, style: nil)
      column(index, outline_level: level, collapsed: collapsed, hidden: hidden, width: width, style: style)
    end

    # Groups multiple rows with outline hierarchy with a block.
    #
    # @example Group rows with a block
    #   sheet.group_rows(level: 1, collapsed: true) do
    #     sheet.row(["Subtotal Item 1"])
    #     sheet.row(["Subtotal Item 2"])
    #   end
    #
    # @param rows [Range, Array, Integer, nil] Unused in streaming write (streaming rows are written on the fly).
    # @param level [Integer] Grouping outline level (1 to 7).
    # @param collapsed [Boolean] Whether the group is collapsed.
    # @param hidden [Boolean] Whether the rows are hidden.
    # @yield Optional block executed with active group outline settings.
    # @return [void]
    # @api public
    #: (?untyped rows, ?level: Integer, ?collapsed: bool, ?hidden: bool) ?{ (self) -> void } -> void
    def group_rows(rows = nil, level: 1, collapsed: false, hidden: false)
      _ = rows
      return unless block_given?

      prev_level = @current_row_outline_level
      prev_collapsed = @current_row_collapsed
      prev_hidden = @current_row_hidden
      @current_row_outline_level = level
      @current_row_collapsed = collapsed
      @current_row_hidden = hidden
      begin
        yield(self)
      ensure
        @current_row_outline_level = prev_level
        @current_row_collapsed = prev_collapsed
        @current_row_hidden = prev_hidden
      end
    end

    # Configures or queries worksheet-level outline and grouping view properties.
    #
    # When called with no arguments, returns the outline properties Hash.
    # When called with arguments, updates outline properties and returns nil.
    #
    # @example Show summary row below detail rows
    #   sheet.outline_properties(summary_below: true, summary_right: true)
    #
    # @param summary_below [Boolean, nil] Flag indicating if summary rows appear below detail rows.
    # @param summary_right [Boolean, nil] Flag indicating if summary columns appear to the right of detail columns.
    # @param show_outline_symbols [Boolean, nil] Flag indicating whether outline symbols are visible.
    # @param apply_styles [Boolean, nil] Flag indicating whether automatic styles apply to outlines.
    # @param opts [Hash] Additional options.
    # @return [Hash<Symbol, untyped>, void]
    # @api public
    #: () -> Hash[Symbol, untyped]
    #: (?summary_below: bool?, ?summary_right: bool?, ?show_outline_symbols: bool?, ?apply_styles: bool?, **untyped opts) -> void
    def outline_properties(summary_below: nil, summary_right: nil, show_outline_symbols: nil, apply_styles: nil, **opts)
      sheet if @current_sheet.nil?

      return @current_sheet_properties.slice(:summary_below, :summary_right, :show_outline_symbols, :apply_styles) if summary_below.nil? && summary_right.nil? && show_outline_symbols.nil? && apply_styles.nil? && opts.empty?

      props = (@current_sheet_properties[:outline_properties] || {}).dup
      props[:summary_below] = summary_below unless summary_below.nil?
      props[:summary_right] = summary_right unless summary_right.nil?
      props[:show_outline_symbols] = show_outline_symbols unless show_outline_symbols.nil?
      props[:apply_styles] = apply_styles unless apply_styles.nil?
      props.merge!(opts)
      @current_sheet_properties[:outline_properties] = props
      @current_sheet_properties[:summary_below] = props[:summary_below] unless props[:summary_below].nil?
      @current_sheet_properties[:summary_right] = props[:summary_right] unless props[:summary_right].nil?
      nil
    end

    # Sets worksheet outline and grouping view properties.
    #
    # @param opts [Hash] Outline properties.
    # @return [void]
    # @api public
    #: (Hash[Symbol, untyped] opts) -> void
    def outline_properties=(opts)
      sheet if @current_sheet.nil?
      h = opts.respond_to?(:to_h) ? opts.to_h : (opts || {})
      outline_properties(**h)
      nil
    end

    # Adds a chart to the current worksheet.
    #
    # @param type [Symbol, String, nil] Optional chart type (e.g. :bar, :line).
    # @param options [Hash] Chart options.
    # @yield [builder]
    # @yieldparam builder [Xlsxrb::ChartBuilder]
    # @return [void]
    # @api public
    #: (?Symbol | String? type, **untyped options) ?{ (ChartBuilder) -> void } -> void
    def chart(type = nil, **options)
      sheet if @current_sheet.nil?

      options = options.merge(type: type) if type
      if block_given?
        builder = ChartBuilder.new
        builder.type = type if type
        yield builder
        options = builder.options.merge(options)
      end

      @current_charts << options
      nil
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
    #: (String | Integer cell, ?(String | Elements::Hyperlink)? url, ?display: String?, ?tooltip: String?, ?location: String?) -> void
    def hyperlink(cell, url = nil, display: nil, tooltip: nil, location: nil)
      sheet if @current_sheet.nil?
      if url.is_a?(Elements::Hyperlink)
        link = { cell: cell }
        link[:url] = url.url if url.url
        link[:display] = url.display || display if url.display || display
        link[:tooltip] = url.tooltip || tooltip if url.tooltip || tooltip
        link[:location] = url.location || location if url.location || location
        @current_hyperlinks << link
        return
      end
      link = { cell: cell }
      link[:url] = url if url
      link[:display] = display if display
      link[:tooltip] = tooltip if tooltip
      link[:location] = location if location
      @current_hyperlinks << link
    end

    # Returns the auto-filter range on the active sheet, or nil if none.
    #
    # @return [String, nil]
    # @api public
    #: () -> String?
    def auto_filter
      ensure_sheet_started!
      @current_auto_filter
    end

    # Sets the auto-filter range for the current worksheet.
    #
    # @param range [String, nil] Cell range (e.g. "A1:E100").
    # @return [String, nil]
    # @api public
    #: (String? range) -> String?
    def auto_filter=(range)
      ensure_sheet_started!
      @current_auto_filter = range&.to_s
    end

    # Sets filter criteria on a column in the auto-filter.
    #
    # @param col_id [Integer] 0-based column index.
    # @param filter [Hash] Filter criteria.
    # @return [void]
    # @api public
    # Sets filter criteria on an auto-filter column.
    #
    # @param col_id [Integer, String, Symbol] 0-based column index relative to auto-filter or column letter ("A", :B).
    # @param filter [Hash] Filter definition.
    # @return [void]
    # @api public
    #: (Integer | String | Symbol col_id, (Hash[untyped, untyped] | Array[untyped]) filter) -> void
    def filter_column(col_id, filter)
      sheet if @current_sheet.nil?
      col_id = Elements::Cell._column_index(col_id) if col_id.is_a?(String) || col_id.is_a?(Symbol)
      @current_filter_columns[col_id] = filter
      nil
    end

    # Configures or queries column sort state on the active sheet.
    #
    # When called with no arguments, returns the {Elements::SortState} value object.
    # When called with arguments, updates sort state and returns nil.
    #
    # @param ref [String, nil] The sorted range.
    # @param sort_conditions [Array<Hash>, nil] Sort conditions.
    # @param case_sensitive [Boolean] Whether sorting is case sensitive.
    # @param sort_method [String, nil] Sort method.
    # @param opts [Hash] Additional options.
    # @return [Elements::SortState, nil, void]
    # @api public
    #: () -> Elements::SortState?
    #: (?String? ref, ?Array[Hash[untyped, untyped]]? sort_conditions, ?case_sensitive: bool, ?sort_method: String?, **untyped opts) -> void
    def sort_state(ref = nil, sort_conditions = nil, case_sensitive: false, sort_method: nil, **opts)
      sheet if @current_sheet.nil?
      if ref.nil? && sort_conditions.nil?
        return @current_sort_state ? Elements::SortState.new(**@current_sort_state) : nil
      end

      entry = { ref: ref, sort_conditions: sort_conditions, case_sensitive: case_sensitive }
      entry[:sort_method] = sort_method if sort_method
      @current_sort_state = entry.merge(opts)
      nil
    end

    # Sets column sort state on the active sheet.
    #
    # @param opts [Elements::SortState, Hash, nil] Sort state options.
    # @return [void]
    # @api public
    #: (Elements::SortState | Hash[Symbol, untyped] | nil opts) -> void
    def sort_state=(opts)
      sheet if @current_sheet.nil?
      @current_sort_state = opts&.respond_to?(:to_h) ? opts.to_h : opts
      nil
    end

    # Adds a data validation rule.
    #
    # @param sqref [String, nil] Cell range.
    # @param opts [Hash] Validation options (or coordinate keywords: `columns:`, `rows:`, `except_rows:`, `list:`, `in:`).
    # @yield [data_validation_builder] Optional configuration block.
    # @return [void]
    # @api public
    #: (?String? sqref, **untyped opts) ?{ (DataValidationBuilder) -> void } -> void
    def data_validation(sqref = nil, **opts, &)
      sheet if @current_sheet.nil?
      normalized_list = DslHelpers.normalize_data_validation(sqref, **opts, &)
      @current_data_validations.concat(normalized_list)
    end

    # Adds a conditional formatting rule.
    # Supports block configuration via {ConditionalFormatBuilder} and coordinate targeting (:columns, :rows).
    #
    # @param sqref [String, nil] Cell range (e.g. "A1:A10").
    # @param opts [Hash] Rule options (e.g. `type:`, `operator:`, `formula:`, `dxf:`, `format_id:`, `font_color:`, `fill_color:`, `bold:`, `italic:`).
    # @option opts [Hash] :dxf Hierarchical differential format specification (e.g. `dxf: { font: { bold: false, color: { auto: true } }, fill: { pattern: "none", bg_color: "FFFF00" } }`).
    # @option opts [Integer] :format_id Pre-registered differential format index (dxfId).
    # @yield [conditional_format_builder] Optional block for configuration.
    # @return [void]
    # @api public
    #: (?String? sqref, **untyped opts) ?{ (ConditionalFormatBuilder) -> void } -> void
    def conditional_format(sqref = nil, **opts, &)
      sheet if @current_sheet.nil?
      normalized_list = DslHelpers.normalize_conditional_format(sqref, **opts, &)
      @current_conditional_formats.concat(normalized_list)
    end

    # Adds an Excel Table (ListObject).
    #
    # @param ref [String, nil] Table range.
    # @param columns [Array<String>, nil] Column names.
    # @param name [String, nil] Table name.
    # @param display_name [String, nil] Display name.
    # @param style [String, nil] Table style.
    # @param opts [Hash] Additional options.
    # @yield [table_builder] Optional configuration block.
    # @return [void]
    # @api public
    #: (?String? ref, ?columns: untyped, ?name: untyped, ?display_name: untyped, ?style: untyped, **untyped opts) ?{ (TableBuilder) -> void } -> void
    def table(ref = nil, columns: nil, name: nil, display_name: nil, style: nil, **opts)
      sheet if @current_sheet.nil?
      opts = opts.dup
      if block_given?
        builder = TableBuilder.new(ref, columns: columns, name: name, display_name: display_name, style: style, **opts)
        yield builder
        opts = builder.to_h
        ref = opts.delete(:ref) || ref
        columns = opts.delete(:columns) || columns
        name = opts.delete(:name) || name
        display_name = opts.delete(:display_name) || display_name
        style = opts.delete(:style) || style
      end
      raise ArgumentError, "table ref must be provided" unless ref
      raise ArgumentError, "table columns must be provided" if columns.nil?

      @current_tables << DslHelpers.normalize_table_options(ref, columns: columns, name: name, display_name: display_name, style: style, **opts)
    end

    # Adds a Pivot Table.
    #
    # @param source_ref [String] Source data range.
    # @param row_fields [Array<Integer>] 0-based field indices for rows.
    # @param data_fields [Array<Hash>] Data fields.
    # @param column_fields [Array<Integer>, nil] 0-based field indices for columns.
    # @param ref [String, nil] Top-left destination cell (default: "E1").
    # @param name [String, nil] Pivot table name.
    # @param field_names [Array<String>, nil] Override field names.
    # @param items [Array, nil] Items configuration.
    # @param opts [Hash] Additional options.
    # @return [void]
    # @api public
    #: (untyped source_ref, row_fields: untyped, data_fields: untyped, ?column_fields: untyped, ?ref: untyped, ?name: untyped, ?field_names: untyped, ?items: untyped, **untyped opts) -> void
    def pivot_table(source_ref, row_fields:, data_fields:, column_fields: nil, ref: nil, name: nil, field_names: nil, items: nil, **opts)
      sheet if @current_sheet.nil?
      c_fields = column_fields || []
      target_ref = ref || "E1"
      @current_pivot_tables ||= []
      entry = {
        source_ref: source_ref, row_fields: row_fields,
        data_fields: data_fields, column_fields: c_fields,
        ref: target_ref, destination: target_ref, name: name,
        field_names: field_names, items: items
      }.merge(opts)
      @current_pivot_tables << entry
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
      ref_str = cell.to_s
      @current_comments << { ref: ref_str, text: text, author: author }
      nil
    end

    # Adds a threaded comment to a cell.
    #
    # @param cell [String, Integer] Cell coordinate (e.g. "A1").
    # @param text [String] Comment text.
    # @param author [String] Author name.
    # @param timestamp [String, Time, nil] Optional ISO 8601 timestamp.
    # @param id [String, nil] Optional GUID for this comment.
    # @param parent_id [String, nil] Optional parent comment GUID for replies.
    # @param done [Boolean, nil] Optional resolved flag.
    # @param user_id [String, nil] Optional user ID.
    # @param provider_id [String, nil] Optional provider ID.
    # @return [void]
    # @api public
    #: (String | Integer cell, String text, ?author: untyped, ?timestamp: untyped, ?id: String?, ?parent_id: String?, ?done: bool?, ?user_id: String?, ?provider_id: String?) -> void
    def threaded_comment(cell, text, author: "Author", timestamp: nil, id: nil, parent_id: nil, done: false, user_id: nil, provider_id: nil)
      sheet if @current_sheet.nil?
      ref_str = cell.to_s
      @current_threaded_comments << {
        ref: ref_str,
        text: text,
        author: author,
        timestamp: timestamp,
        created_at: timestamp,
        id: id,
        parent_id: parent_id,
        done: done ? true : false,
        resolved: done ? true : false,
        user_id: user_id,
        author_id: user_id,
        provider_id: provider_id
      }.compact
      nil
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
    # @example
    #   sheet.merge("A1:C1")
    #
    # @param range [String] Cell range (e.g. "A1:B2").
    # @return [void]
    # @api public
    #: (String range) -> void
    def merge(range)
      sheet if @current_sheet.nil?
      canonical = DslHelpers.normalize_merge_range(range, strict_excel_mode: @strict_excel_mode)
      @current_merge_cells << canonical unless @current_merge_cells.include?(canonical)
    end

    # Configures or queries freeze panes (<pane state="frozen">).
    #
    # When called with no arguments, returns the {Elements::FreezePane} value object or nil.
    # When called with arguments, configures the freeze pane and returns nil.
    #
    # @param row [Integer, nil] Number of rows to freeze.
    # @param column [Integer, String, nil] Number of columns to freeze.
    # @param top_left_cell [String, nil] Top-left cell reference in bottom-right pane.
    # @param active_pane [String, Symbol, nil] Active pane identifier.
    # @return [Elements::FreezePane, nil, void]
    # @api public
    #: () -> Elements::FreezePane?
    #: (?row: Integer | nil, ?column: (Integer | String)?, ?top_left_cell: String?, ?active_pane: (String | Symbol)?) -> void
    def freeze_pane(row: nil, column: nil, top_left_cell: nil, active_pane: nil)
      sheet if @current_sheet.nil?
      if row.nil? && column.nil? && top_left_cell.nil? && active_pane.nil?
        return @current_freeze_pane ? Elements::FreezePane.new(**@current_freeze_pane) : nil
      end

      col_val = column ? Elements::Cell._column_index(column) : 0
      entry = { row: row || 0, column: col_val }
      entry[:top_left_cell] = top_left_cell if top_left_cell
      entry[:active_pane] = active_pane.to_s if active_pane
      @current_freeze_pane = entry
      nil
    end

    # Sets or clears freeze pane configuration.
    #
    # @param opts [Elements::FreezePane, Hash, nil] Freeze pane options.
    # @return [void]
    # @api public
    #: (Elements::FreezePane | Hash[Symbol, untyped] | nil opts) -> void
    def freeze_pane=(opts)
      sheet if @current_sheet.nil?
      if opts.nil?
        @current_freeze_pane = nil
      else
        h = opts.respond_to?(:to_h) ? opts.to_h : opts
        c = h[:column] || h[:col]
        c = Elements::Cell._column_index(c) if c
        entry = { row: h[:row] || 0, column: c || 0 }
        entry[:top_left_cell] = h[:top_left_cell] if h[:top_left_cell]
        entry[:active_pane] = h[:active_pane].to_s if h[:active_pane]
        @current_freeze_pane = entry
      end
      nil
    end

    # Splits window panes without freezing or queries split pane settings (<pane>).
    #
    # When called with no arguments, returns the {Elements::SplitPane} value object or nil.
    # When called with arguments, configures the split pane and returns nil.
    #
    # @param x_split [Integer, nil] Horizontal split in points.
    # @param y_split [Integer, nil] Vertical split in points.
    # @param top_left_cell [String, nil] Top-left cell reference in bottom-right pane.
    # @param active_pane [String, Symbol, nil] Active pane identifier.
    # @param state [String, Symbol, nil] Split state.
    # @return [Elements::SplitPane, nil, void]
    # @api public
    #: () -> Elements::SplitPane?
    #: (?x_split: Integer | nil, ?y_split: Integer | nil, ?top_left_cell: String?, ?active_pane: (String | Symbol)?) -> void
    def split_pane(x_split: nil, y_split: nil, top_left_cell: nil, active_pane: nil)
      sheet if @current_sheet.nil?
      if x_split.nil? && y_split.nil? && top_left_cell.nil? && active_pane.nil?
        return @current_split_pane ? Elements::SplitPane.new(**@current_split_pane) : nil
      end

      entry = { x_split: x_split || 0, y_split: y_split || 0, top_left_cell: top_left_cell }
      entry[:active_pane] = active_pane.to_s if active_pane
      @current_split_pane = entry
      nil
    end

    # Sets or clears split pane configuration.
    #
    # @param opts [Elements::SplitPane, Hash, nil] Split pane options.
    # @return [void]
    # @api public
    #: (Elements::SplitPane | Hash[Symbol, untyped] | nil opts) -> void
    def split_pane=(opts)
      sheet if @current_sheet.nil?
      if opts.nil?
        @current_split_pane = nil
      else
        h = opts.respond_to?(:to_h) ? opts.to_h : opts
        entry = { x_split: h[:x_split] || 0, y_split: h[:y_split] || 0, top_left_cell: h[:top_left_cell] }
        entry[:active_pane] = h[:active_pane].to_s if h[:active_pane]
        @current_split_pane = entry
      end
      nil
    end

    # Sets or queries the active cell selection on the current sheet (<selection>).
    #
    # When called with no arguments, returns the selection ({Elements::Selection}, Array, or nil).
    # When called with arguments, updates the selection and returns nil.
    #
    # @param active_cell [String, nil] Active cell reference (e.g. "A1") or range (e.g. "A1:D10").
    # @param sqref [String, nil] Selection range.
    # @param pane [String, Symbol, nil] Target pane.
    # @return [Elements::Selection, Array<Elements::Selection>, nil, void]
    # @api public
    #: () -> (Elements::Selection | Array[Elements::Selection])?
    #: (?String? active_cell, ?sqref: String?, ?pane: (String | Symbol)?) -> void
    def selection(active_cell = nil, sqref: nil, pane: nil)
      sheet if @current_sheet.nil?
      if active_cell.nil? && sqref.nil? && pane.nil?
        return nil if @current_selection.nil?
        return @current_selection.map { |s| Elements::Selection.new(**s) } if @current_selection.is_a?(Array)

        return Elements::Selection.new(**@current_selection)
      end

      if sqref.nil? && active_cell.is_a?(String) && active_cell.include?(":")
        sqref = active_cell
        active_cell = sqref.split(":").first
      else
        sqref ||= active_cell
      end
      sel = { active_cell: active_cell, sqref: sqref }
      sel[:pane] = pane if pane
      @current_selection = sel
      nil
    end

    # Sets or clears the selection configuration.
    #
    # @param opts [Elements::Selection, Array<Elements::Selection, Hash>, Hash, nil]
    # @return [void]
    # @api public
    #: (Elements::Selection | Array[Elements::Selection | Hash[Symbol, untyped]] | Hash[Symbol, untyped] | nil opts) -> void
    def selection=(opts)
      sheet if @current_sheet.nil?
      @current_selection = if opts.nil?
                             nil
                           elsif opts.is_a?(Array)
                             opts.map { |o| o.respond_to?(:to_h) ? o.to_h : o }
                           else
                             opts.respond_to?(:to_h) ? opts.to_h : opts
                           end
      nil
    end

    # Configures or queries page margins in inches for printing (<pageMargins>).
    #
    # When called with no arguments, returns the {Elements::PageMargins} value object or nil.
    # When called with arguments, configures page margins and returns nil.
    #
    # @param left [Float, nil] Left margin.
    # @param right [Float, nil] Right margin.
    # @param top [Float, nil] Top margin.
    # @param bottom [Float, nil] Bottom margin.
    # @param header [Float, nil] Header margin.
    # @param footer [Float, nil] Footer margin.
    # @return [Elements::PageMargins, nil, void]
    # @api public
    #: () -> Elements::PageMargins?
    #: (?left: Float?, ?right: Float?, ?top: Float?, ?bottom: Float?, ?header: Float?, ?footer: Float?) -> void
    def page_margins(left: nil, right: nil, top: nil, bottom: nil, header: nil, footer: nil)
      sheet if @current_sheet.nil?
      if left.nil? && right.nil? && top.nil? && bottom.nil? && header.nil? && footer.nil?
        return @current_page_margins ? Elements::PageMargins.new(**@current_page_margins) : nil
      end

      @current_page_margins = DslHelpers.normalize_page_margins(
        left: left, right: right, top: top, bottom: bottom, header: header, footer: footer
      )
      nil
    end

    # Sets or clears page margins configuration.
    #
    # @param opts [Hash, Elements::PageMargins, nil]
    # @return [void]
    # @api public
    #: (Hash[Symbol, untyped] | Elements::PageMargins | nil opts) -> void
    def page_margins=(opts)
      sheet if @current_sheet.nil?
      if opts.nil?
        @current_page_margins = nil
      else
        h = opts.respond_to?(:to_h) ? opts.to_h : opts
        @current_page_margins = DslHelpers.normalize_page_margins(**h)
      end
      nil
    end

    # Configures or queries page setup configuration (<pageSetup>).
    #
    # When called with no arguments, returns the {Elements::PageSetup} value object.
    # When called with keyword arguments, updates page setup and returns nil.
    #
    # @param opts [Hash] Page setup options (e.g. orientation: :landscape).
    # @return [Elements::PageSetup, void]
    # @api public
    #: () -> Elements::PageSetup
    #: (**untyped opts) -> void
    def page_setup(**opts)
      sheet if @current_sheet.nil?
      return Elements::PageSetup.new(**@current_page_setup) if opts.empty?

      @current_page_setup.merge!(opts)
      @current_sheet_properties[:fit_to_page] = true if (opts.key?(:fit_to_width) || opts.key?(:fit_to_height)) && !@current_sheet_properties.key?(:fit_to_page)
      nil
    end

    # Sets page setup configuration.
    #
    # @param opts [Hash, Elements::PageSetup]
    # @return [void]
    # @api public
    #: (Hash[Symbol, untyped] | Elements::PageSetup opts) -> void
    def page_setup=(opts)
      sheet if @current_sheet.nil?
      h = opts.respond_to?(:to_h) ? opts.to_h : (opts || {})
      @current_page_setup = h.transform_keys(&:to_sym)
      nil
    end

    # Configures or queries headers and footers for printing (<headerFooter>).
    #
    # When called with no arguments, returns the {Elements::HeaderFooter} value object.
    # When called with keyword arguments, updates headers and footers and returns nil.
    #
    # @param opts [Hash] Header/footer options.
    # @return [Elements::HeaderFooter, void]
    # @api public
    #: () -> Elements::HeaderFooter
    #: (**untyped opts) -> void
    def header_footer(**opts)
      sheet if @current_sheet.nil?
      return Elements::HeaderFooter.new(**@current_header_footer) if opts.empty?

      @current_header_footer.merge!(opts)
      nil
    end

    # Sets header and footer configuration.
    #
    # @param opts [Hash, Elements::HeaderFooter]
    # @return [void]
    # @api public
    #: (Hash[Symbol, untyped] | Elements::HeaderFooter opts) -> void
    def header_footer=(opts)
      sheet if @current_sheet.nil?
      h = opts.respond_to?(:to_h) ? opts.to_h : (opts || {})
      @current_header_footer = h.transform_keys(&:to_sym)
      nil
    end

    # Configures or queries print options (<printOptions>).
    #
    # When called with no arguments, returns the {Elements::PrintOptions} value object.
    # When called with keyword options, configures print options and returns nil.
    #
    # @param opts [Hash] Additional print options.
    # @return [Elements::PrintOptions, void]
    # @api public
    #: () -> Elements::PrintOptions
    #: (**untyped opts) -> void
    def print_options(**opts)
      sheet if @current_sheet.nil?
      return Elements::PrintOptions.new(**@current_print_options) if opts.empty?

      opts.each { |k, v| @current_print_options[k.to_sym] = v }
      nil
    end

    # Replaces print options with the given hash or {Elements::PrintOptions} object.
    #
    # @param opts [Hash, Elements::PrintOptions]
    # @return [void]
    # @api public
    #: (Hash[Symbol, untyped] | Elements::PrintOptions opts) -> void
    def print_options=(opts)
      sheet if @current_sheet.nil?
      h = opts.respond_to?(:to_h) ? opts.to_h : (opts || {})
      @current_print_options = h.transform_keys(&:to_sym)
      nil
    end

    # Sets or queries sheet-level protection with optional password hashing (<sheetProtection>).
    #
    # When called with no arguments, returns the {Elements::SheetProtection} value object or nil.
    # When called with keyword arguments, sets sheet protection and returns nil.
    #
    # @param opts [Hash] Protection options.
    # @return [Elements::SheetProtection, nil, void]
    # @api public
    #: () -> Elements::SheetProtection?
    #: (**untyped opts) -> void
    def sheet_protection(**opts)
      sheet if @current_sheet.nil?
      if opts.empty?
        @current_sheet_protection ? Elements::SheetProtection.new(**@current_sheet_protection) : nil
      else
        @current_sheet_protection = DslHelpers.normalize_protection_options(opts)
        nil
      end
    end

    # Sets or clears sheet protection configuration.
    #
    # @param opts [Elements::SheetProtection, Hash, Boolean, nil]
    # @return [void]
    # @api public
    #: (Elements::SheetProtection | Hash[Symbol, untyped] | bool | nil opts) -> void
    def sheet_protection=(opts)
      sheet if @current_sheet.nil?
      if opts.nil? || opts == false
        @current_sheet_protection = nil
      elsif opts == true
        @current_sheet_protection = DslHelpers.normalize_protection_options(sheet: true)
      else
        h = opts.respond_to?(:to_h) ? opts.to_h : opts
        @current_sheet_protection = DslHelpers.normalize_protection_options(h)
      end
      nil
    end

    # Inserts an embedded image into the active sheet.
    #
    # @param file_data [String] Binary image data.
    # @param ext [String] File extension ("png", "jpeg", etc.).
    # @param from_column [Integer, nil] Top-left starting column.
    # @param from_row [Integer] Top-left starting row.
    # @param to_column [Integer, nil] Bottom-right ending column.
    # @param to_row [Integer] Bottom-right ending row.
    # @param opts [Hash] Additional options.
    # @return [void]
    # @api public
    #: (String file_data, ?ext: ::String, ?from_column: ::Integer?, ?from_row: ::Integer, ?to_column: ::Integer?, ?to_row: ::Integer, **untyped opts) -> void
    def image(file_data, ext: "png", from_column: nil, from_row: 0, to_column: nil, to_row: 10, **opts)
      fc = from_column || 0
      tc = to_column || 5
      sheet if @current_sheet.nil?
      img = { file_data: file_data, ext: ext, from_column: fc, from_row: from_row, to_column: tc, to_row: to_row }
      img.merge!(opts)
      @current_images << img
    end

    # Worksheet background image (<picture>).
    #
    # @return [Elements::BackgroundImage, nil]
    # @api public
    #: () -> Elements::BackgroundImage?
    def background_image
      sheet if @current_sheet.nil?
      return nil unless @current_background_image

      @current_background_image.is_a?(Elements::BackgroundImage) ? @current_background_image : Elements::BackgroundImage.new(**@current_background_image)
    end

    # Sets the worksheet background image (<picture>).
    #
    # @param val [String, Hash, Elements::Image, nil] Image binary data, file path, or image options hash.
    # @return [untyped]
    # @api public
    #: (untyped val) -> untyped
    def background_image=(val)
      sheet if @current_sheet.nil?
      if val.nil?
        @current_background_image = nil
        return
      end

      data = val
      inferred_ext = nil
      if data.is_a?(String) && !data.include?("\x00") && File.file?(data)
        inferred_ext = File.extname(data).delete_prefix(".").downcase
        data = File.binread(data)
      elsif data.is_a?(Hash)
        inferred_ext = data[:ext] || (data[:filename] && File.extname(data[:filename]).delete_prefix("."))
        data = data[:file_data] || data[:data]
      elsif data.is_a?(Elements::Image)
        inferred_ext = File.extname(data.filename).delete_prefix(".").downcase
        data = data.data
      end

      final_ext = (inferred_ext || (val.is_a?(Hash) ? val[:ext] : nil) || "png").to_s.delete_prefix(".").downcase
      mime = case final_ext
             when "png" then "image/png"
             when "jpg", "jpeg" then "image/jpeg"
             when "gif" then "image/gif"
             when "bmp" then "image/bmp"
             when "emf" then "image/x-emf"
             when "wmf" then "image/x-wmf"
             else "application/octet-stream"
             end

      opts = val.is_a?(Hash) ? val : {}
      @current_background_image = {
        data: data&.b,
        file_data: data&.b,
        extension: final_ext,
        ext: final_ext,
        path: "background.#{final_ext}",
        filename: "background.#{final_ext}",
        content_type: mime
      }.merge(opts).freeze
      val
    end

    # Inserts a drawing shape.
    #
    # @param preset [String] Preset shape name.
    # @param text [String, nil] Label text.
    # @param from_column [Integer, nil] Top-left column.
    # @param from_row [Integer] Top-left row.
    # @param to_column [Integer, nil] Bottom-right column.
    # @param to_row [Integer] Bottom-right row.
    # @param opts [Hash] Additional options.
    # @return [void]
    # @api public
    #: (**untyped opts) -> void
    def shape(preset: "rect", text: nil, from_column: nil, from_row: 0, to_column: nil, to_row: 5, **opts)
      fc = from_column || 0
      tc = to_column || 5
      sheet if @current_sheet.nil?
      shape_data = { preset: preset, text: text, from_column: fc, from_row: from_row, to_column: tc, to_row: to_row }
      shape_data[:name] = opts.delete(:name) || "Shape #{@current_shapes.size + 1}"
      shape_data.merge!(opts)
      @current_shapes << shape_data
    end

    # Gets the sheet visibility state (:visible, :hidden, or :very_hidden).
    #
    # @return [Symbol]
    # @api public
    #: () -> Symbol
    def state
      sheet if @current_sheet.nil?
      @current_sheet_state || :visible
    end

    # Sets the sheet visibility state (:visible, :hidden, or :very_hidden).
    #
    # @param new_state [Symbol, String] The new state.
    # @return [Symbol]
    # @api public
    #: (Symbol | String new_state) -> Symbol
    def state=(new_state)
      sheet if @current_sheet.nil?
      sheet_properties(state: new_state)
      @current_sheet_state
    end

    # Gets the tab color of the current worksheet.
    #
    # @return [String, Hash, nil]
    # @api public
    #: () -> (String | Hash[Symbol, untyped] | nil)
    def tab_color
      sheet if @current_sheet.nil?
      @current_sheet_properties[:tab_color]
    end

    # Sets the tab color of the current worksheet.
    #
    # @param spec [String, Hash, Symbol, nil] Tab color hex string or hash options.
    # @return [String, Hash, nil]
    # @api public
    #: (String | Symbol | Hash[Symbol, untyped] | nil spec) -> (String | Hash[Symbol, untyped] | nil)
    def tab_color=(spec)
      sheet if @current_sheet.nil?
      @current_sheet_properties[:tab_color] = spec
    end

    # Configures or queries sheet properties (<sheetPr>).
    #
    # When called with no arguments, returns the {Elements::SheetProperties} value object.
    # When called with keyword options, configures properties and returns nil.
    #
    # @param opts [Hash] Additional sheet properties.
    # @return [Elements::SheetProperties, void]
    # @api public
    #: () -> Elements::SheetProperties
    #: (**untyped opts) -> void
    def sheet_properties(**opts)
      sheet if @current_sheet.nil?
      return Elements::SheetProperties.new(**@current_sheet_properties) if opts.empty?

      opts.each do |k, v|
        sym = k.to_sym
        if sym == :state
          val = v.to_sym
          raise ArgumentError, "Invalid sheet state ':#{val}'. Allowed states are :visible, :hidden, :very_hidden" unless VALID_SHEET_STATES.include?(val)

          @current_sheet_state = val
        elsif sym == :hidden
          @current_sheet_state = v ? :hidden : :visible
        else
          @current_sheet_properties[sym] = v
        end
      end
      nil
    end

    # Replaces worksheet properties with the given hash or {Elements::SheetProperties} object.
    #
    # @param opts [Hash, Elements::SheetProperties]
    # @return [void]
    # @api public
    #: (Hash[Symbol, untyped] | Elements::SheetProperties opts) -> void
    def sheet_properties=(opts)
      sheet if @current_sheet.nil?
      h = opts.respond_to?(:to_h) ? opts.to_h : (opts || {})
      @current_sheet_properties = {}
      sheet_properties(**h)
      nil
    end

    # Configures or queries sheet view display properties (<sheetView>).
    #
    # When called with no arguments, returns the {Elements::SheetView} value object.
    # When called with keyword options, configures view properties and returns nil.
    #
    # @param opts [Hash] Additional view properties.
    # @return [Elements::SheetView, void]
    # @api public
    #: () -> Elements::SheetView
    #: (**untyped opts) -> void
    def sheet_view(**opts)
      sheet if @current_sheet.nil?
      return Elements::SheetView.new(**@current_sheet_view) if opts.empty?

      opts.each { |k, v| @current_sheet_view[k.to_sym] = v }
      nil
    end

    # Replaces sheet view properties with the given hash or {Elements::SheetView} object.
    #
    # @param opts [Hash, Elements::SheetView]
    # @return [void]
    # @api public
    #: (Hash[Symbol, untyped] | Elements::SheetView opts) -> void
    def sheet_view=(opts)
      sheet if @current_sheet.nil?
      h = opts.respond_to?(:to_h) ? opts.to_h : (opts || {})
      @current_sheet_view = h.transform_keys(&:to_sym)
      nil
    end

    # Configures or queries sheet format properties (<sheetFormatPr>).
    #
    # When called with no arguments, returns the {Elements::SheetFormat} value object.
    # When called with keyword options, configures format properties and returns nil.
    #
    # @param opts [Hash] Keyword arguments for sheet format properties.
    # @return [Elements::SheetFormat, void]
    # @api public
    #: () -> Elements::SheetFormat
    #: (**untyped opts) -> void
    def sheet_format(**opts)
      sheet if @current_sheet.nil?
      return Elements::SheetFormat.new(**@current_sheet_format) if opts.empty?

      opts.each { |k, v| @current_sheet_format[k.to_sym] = v }
      nil
    end

    # Replaces sheet format properties with the given hash or {Elements::SheetFormat} object.
    #
    # @param opts [Hash, Elements::SheetFormat]
    # @return [void]
    # @api public
    #: (Hash[Symbol, untyped] | Elements::SheetFormat opts) -> void
    def sheet_format=(opts)
      sheet if @current_sheet.nil?
      h = opts.respond_to?(:to_h) ? opts.to_h : (opts || {})
      @current_sheet_format = h.transform_keys(&:to_sym)
      nil
    end

    # Inserts a horizontal page break after the given 0-based row.
    #
    # @param row_index [Integer] 0-based row index.
    # @return [void]
    # @api public
    #: (Integer row_index) -> void
    def row_break(row_index)
      sheet if @current_sheet.nil?
      @current_row_breaks << { id: row_index, max: 16_383, man: 1 }
    end

    # Inserts a vertical page break after the given column.
    #
    # @param col [Integer, String, Symbol] 0-based column index or letter ("B", :B).
    # @return [void]
    # @api public
    #: (Integer | String | Symbol col) -> void
    def column_break(col)
      col_index = Elements::Cell._column_index(col)
      sheet if @current_sheet.nil?
      @current_col_breaks << { id: col_index, max: 1_048_575, man: 1 }
    end

    # Adds a protected range to the current sheet.
    #
    # @param name [String] Name of the protected range.
    # @param sqref [String] Cell range (e.g. "A1:B10").
    # @param password [String, nil] Optional password or hash.
    # @param opts [Hash] Additional options.
    # @return [void]
    # @api public
    #: (String name, String sqref, ?password: String?, **untyped opts) -> void
    def protected_range(name, sqref, password: nil, **opts)
      sheet if @current_sheet.nil?
      entry = { name: name.to_s, sqref: sqref.to_s }
      entry[:password] = password if password
      entry.merge!(opts) unless opts.empty?
      @current_protected_ranges << DslHelpers.normalize_protection_options(entry)
    end

    # Adds a cell watch to the current sheet.
    #
    # @param ref [String] Cell reference (e.g. "A1").
    # @return [void]
    # @api public
    #: (String ref) -> void
    def cell_watch(ref)
      sheet if @current_sheet.nil?
      @current_cell_watches << ref.to_s
    end

    # Adds an ignored error entry to the current sheet.
    #
    # @param sqref [String] Cell range reference (e.g. "A1" or "A1:B10").
    # @param opts [Hash] Error check flags to ignore.
    # @return [void]
    # @api public
    #: (String sqref, ?eval_error: bool, ?two_digit_text_year: bool, ?number_stored_as_text: bool, ?formula: bool, ?formula_range: bool, ?unlocked_formula: bool, ?empty_cell_reference: bool, ?list_data_validation: bool, ?calculated_column: bool) -> void
    def ignored_error(sqref, **opts)
      sheet if @current_sheet.nil?
      @current_ignored_errors << { sqref: sqref.to_s }.merge(opts)
    end

    # Returns the ignored error entries configured for the current sheet.
    #
    # @return [Array<Elements::IgnoredError>]
    # @api public
    #: () -> Array[Elements::IgnoredError]
    def ignored_errors
      @current_ignored_errors.map { |ie| ie.is_a?(Elements::IgnoredError) ? ie : Elements::IgnoredError.new(**ie) }.freeze
    end

    # Phonetic (furigana) guide properties for the current sheet.
    #
    # @return [Elements::PhoneticProperties, nil]
    # @api public
    #: () -> Elements::PhoneticProperties?
    def phonetic_properties
      sheet if @current_sheet.nil?
      return nil unless @current_phonetic_properties

      @current_phonetic_properties.is_a?(Elements::PhoneticProperties) ? @current_phonetic_properties : Elements::PhoneticProperties.new(**@current_phonetic_properties)
    end

    # Sets phonetic (furigana) guide properties for the current sheet.
    #
    # @param spec [Hash, Integer, nil]
    # @return [Hash[Symbol, untyped], Integer, nil]
    # @api public
    #: (Hash[Symbol, untyped] | Integer | nil spec) -> (Hash[Symbol, untyped] | Integer | nil)
    def phonetic_properties=(spec)
      sheet if @current_sheet.nil?
      if spec.nil?
        @current_phonetic_properties = nil
      else
        props = if spec.is_a?(Hash)
                  spec.dup
                elsif spec.is_a?(Integer)
                  { font_id: spec }
                else
                  {}
                end
        @current_phonetic_properties = WorksheetBuilder.normalize_phonetic_properties(props)
      end
      spec
    end

    # Gets the internal code name of the current worksheet (or workbook if no sheet is open).
    #
    # @return [String, nil]
    # @api public
    #: () -> String?
    def code_name
      if @current_sheet.nil?
        @workbook_properties[:code_name]
      else
        @current_sheet_properties[:code_name]
      end
    end

    # Sets the internal code name of the current worksheet (or workbook if no sheet is open).
    #
    # @param val [String, Symbol, nil] The code name.
    # @return [String, nil]
    # @api public
    #: (String | Symbol | nil val) -> String?
    def code_name=(val)
      if @current_sheet.nil?
        @workbook_properties[:code_name] = val&.to_s
      else
        @current_sheet_properties[:code_name] = val&.to_s
      end
    end

    # Returns whether page setup fits content to page width/height for the current sheet.
    #
    # @return [Boolean]
    # @api public
    #: () -> bool
    def fit_to_page?
      sheet if @current_sheet.nil?
      @current_sheet_properties[:fit_to_page] ? true : false
    end

    # Sets whether page setup fits content to page width/height for the current sheet.
    #
    # @param enabled [Boolean, nil] Whether to fit to page.
    # @return [Boolean]
    # @api public
    #: (bool? enabled) -> bool?
    def fit_to_page=(enabled)
      sheet if @current_sheet.nil?
      @current_sheet_properties[:fit_to_page] = enabled ? true : false
    end

    # Worksheet calculation properties (<sheetCalcPr>) for the current sheet.
    #
    # @example Instruct Excel to recalculate formulas when loading sheet
    #   w.sheet_calc_properties = { full_calc_on_load: true }
    #
    # @return [Elements::SheetCalcProperties, nil]
    # @api public
    #: () -> Elements::SheetCalcProperties?
    def sheet_calc_properties
      sheet if @current_sheet.nil?
      calc = @current_sheet_calc_pr || (@current_sheet_properties[:full_calc_on_load] ? { full_calc_on_load: true } : nil)
      return nil unless calc

      calc.is_a?(Elements::SheetCalcProperties) ? calc : Elements::SheetCalcProperties.new(**calc)
    end

    # Sets worksheet calculation properties (<sheetCalcPr>) for the current sheet.
    #
    # @param spec [Hash, nil] Calculation properties hash.
    # @return [Hash, nil]
    # @api public
    #: (Hash[Symbol, untyped]? spec) -> Hash[Symbol, untyped]?
    def sheet_calc_properties=(spec)
      sheet if @current_sheet.nil?
      if spec.nil?
        @current_sheet_properties.delete(:full_calc_on_load)
        @current_sheet_calc_pr = nil
      else
        pr = spec.is_a?(Hash) ? spec.dup : {}
        @current_sheet_properties[:full_calc_on_load] = pr[:full_calc_on_load] unless pr[:full_calc_on_load].nil?
        @current_sheet_calc_pr = pr
      end
      spec
    end

    # --- Workbook-Level Methods ---

    # Adds a defined name.
    #
    # @param name [String] The defined name.
    # @param value [Object, nil] The formula or value expression.
    # @param sheet [String, nil] Reference sheet name.
    # @param scope [String, Symbol, nil] Local scope sheet name or :workbook.
    # @param hidden [Boolean] Whether the defined name is hidden.
    # @param columns [Integer, String, nil] Column index or letter for coordinate-based reference.
    # @param rows [Integer, Range, nil] Row or row range (0-based) for coordinate-based reference.
    # @param range [String, nil] Cell range reference.
    # @return [void]
    # @api public
    #: (String name, ?untyped value, ?sheet: String?, ?scope: (String | Symbol)?, ?hidden: bool, ?columns: untyped, ?rows: untyped, ?range: untyped) -> void
    def defined_name(name, value = nil, sheet: nil, scope: nil, hidden: false, columns: nil, rows: nil, range: nil)
      normalized_val = DslHelpers.normalize_defined_name_value(value, sheet: sheet, columns: columns, rows: rows, range: range)
      entry = { name: name, value: normalized_val, hidden: hidden }
      local_sheet = if scope
                      scope == :workbook ? nil : scope.to_s
                    end
      if local_sheet
        # local_sheet_id will be resolved at close time
        entry[:local_sheet_name] = local_sheet
      end
      @defined_names << entry
    end

    # Returns workbook-level defined names.
    #
    # @return [Array<Elements::DefinedName>]
    # @api public
    #: () -> Array[Elements::DefinedName]
    def defined_names
      @defined_names.map { |dn| dn.is_a?(Elements::DefinedName) ? dn : Elements::DefinedName.new(**dn) }
    end

    # Returns the print area range for the current or named sheet, or nil if none.
    #
    # @param sheet [String, nil] Target sheet name.
    # @return [String, nil]
    # @api public
    #: (?sheet: String?) -> String?
    def print_area(sheet: nil)
      sheet_name = sheet || @current_sheet
      if sheet_name.nil? || sheet_name == @current_sheet
        @current_print_area
      else
        dn = @defined_names.find { |d| d[:name] == "_xlnm.Print_Area" && d[:local_sheet_name] == sheet_name }
        dn ? dn[:value].split("!", 2).last.tr("$", "") : nil
      end
    end

    # Sets the print area for the current worksheet.
    #
    # @param range [String, nil] Cell range (e.g. "A1:G50").
    # @return [String, nil]
    # @api public
    #: (String? range) -> String?
    def print_area=(range)
      set_print_area(range, sheet: @current_sheet)
    end

    # Sets or gets repeating print titles (rows and/or columns) for pagination.
    #
    # When called with no row/column arguments, returns the {Elements::PrintTitles} value object or nil.
    # When called with row/column arguments, sets repeating print titles and returns nil.
    #
    # @param rows [String, Integer, Range, Object, nil] Repeating row range (e.g. "1:2" or 0-based 0..1).
    # @param columns [String, Integer, Symbol, Range, Object, nil] Repeating column range (e.g. "A:B" or 0..1).
    # @param sheet [String, nil] Target sheet name.
    # @return [Elements::PrintTitles, nil, void]
    # @api public
    #: () -> Elements::PrintTitles?
    #: (?rows: (String | Integer | Range[Integer])?, ?columns: (String | Integer | Symbol | Range[untyped])?, ?sheet: String?) -> void
    def print_titles(rows: nil, columns: nil, sheet: nil)
      sheet_name = sheet || @current_sheet || "Sheet1"
      if rows.nil? && columns.nil?
        return Elements::PrintTitles.new(rows: @current_print_titles[:rows], columns: @current_print_titles[:columns]) if sheet_name == @current_sheet && @current_print_titles

        existing = @defined_names.find { |dn| dn[:name] == "_xlnm.Print_Titles" && dn[:local_sheet_name] == sheet_name }
        return nil unless existing

        c = Regexp.last_match(1).delete("$") if existing[:value] =~ /'[^']*'!\$([A-Za-z]+:\$[A-Za-z]+)/
        r = Regexp.last_match(1).delete("$") if existing[:value] =~ /'[^']*'!\$([0-9]+:\$[0-9]+)/
        return Elements::PrintTitles.new(rows: r, columns: c)
      end

      @current_print_titles = { rows: rows, columns: columns }.compact if sheet_name == @current_sheet
      c = DslHelpers.normalize_print_titles_columns(columns)
      r = DslHelpers.normalize_print_titles_rows(rows)
      existing = @defined_names.find { |dn| dn[:name] == "_xlnm.Print_Titles" && dn[:local_sheet_name] == sheet_name }
      if existing && (r.nil? || c.nil?)
        c = Regexp.last_match(1).delete("$") if c.nil? && existing[:value] =~ /'[^']*'!\$([A-Za-z]+:\$[A-Za-z]+)/
        r = Regexp.last_match(1).delete("$") if r.nil? && existing[:value] =~ /'[^']*'!\$([0-9]+:\$[0-9]+)/
      end
      parts = []
      parts << "'#{sheet_name}'!$#{c.sub(":", ":$")}" if c
      parts << "'#{sheet_name}'!$#{r.sub(":", ":$")}" if r
      value = parts.join(",")
      @defined_names.reject! { |dn| dn[:name] == "_xlnm.Print_Titles" && dn[:local_sheet_name] == sheet_name }
      defined_name("_xlnm.Print_Titles", value, sheet: sheet_name, scope: sheet_name) unless value.empty?
      nil
    end

    # Sets repeating print titles for the current worksheet.
    #
    # @param opts [Elements::PrintTitles, Hash, nil] Options with :rows and/or :columns.
    # @return [Elements::PrintTitles, Hash, nil]
    # @api public
    #: (Elements::PrintTitles | Hash[Symbol, untyped] | nil opts) -> (Elements::PrintTitles | Hash[Symbol, untyped] | nil)
    def print_titles=(opts)
      if opts.nil?
        @current_print_titles = nil
        sheet_name = @current_sheet || "Sheet1"
        @defined_names.reject! { |dn| dn[:name] == "_xlnm.Print_Titles" && dn[:local_sheet_name] == sheet_name }
      else
        h = opts.respond_to?(:to_h) ? opts.to_h : opts
        print_titles(**h)
      end
      opts
    end

    # Sets repeating title rows for pagination on the current or named sheet.
    #
    # @example Repeat first row
    #   stream.repeat_rows(0)
    #
    # @param first [Integer, String] First row index (0-based) or range string (e.g. "1:2").
    # @param last [Integer, String, nil] Optional last row index (0-based).
    # @param sheet [String, nil] Target sheet name.
    # @return [void]
    # @api public
    #: (Integer | String first, ?(Integer | String)? last, ?sheet: String?) -> void
    def repeat_rows(first, last = nil, sheet: nil)
      rows_str = if first.is_a?(String)
                   first.include?(":") ? first : "#{first}:#{last || first}"
                 else
                   last ||= first
                   "#{first + 1}:#{last + 1}"
                 end
      print_titles(rows: rows_str, sheet: sheet)
    end

    # Sets repeating title columns for pagination on the current or named sheet.
    #
    # @example Repeat columns A through C
    #   stream.repeat_columns("A", "C")
    #
    # @param first [Integer, String] First column index (0-based) or column name/range.
    # @param last [Integer, String, nil] Optional last column index/name.
    # @param sheet [String, nil] Target sheet name.
    # @return [void]
    # @api public
    #: (Integer | String first, ?(Integer | String)? last, ?sheet: String?) -> void
    def repeat_columns(first, last = nil, sheet: nil)
      cols_str = if first.is_a?(String)
                   first.include?(":") ? first : "#{first}:#{last || first}"
                 else
                   first_name = Elements::Cell._column_letter(first)
                   last_name = if last.is_a?(Integer)
                                 Elements::Cell._column_letter(last)
                               elsif last
                                 last.to_s
                               else
                                 first_name
                               end
                   "#{first_name}:#{last_name}"
                 end
      print_titles(columns: cols_str, sheet: sheet)
    end

    # Sets or returns workbook protection settings.
    #
    # @param opts [Hash] Protection options.
    # @return [Elements::WorkbookProtection, nil, void]
    # @api public
    #: (**String | Integer | bool | nil opts) -> (Elements::WorkbookProtection | nil)
    def workbook_protection(**opts)
      if opts.empty?
        @workbook_protection ? Elements::WorkbookProtection.new(**@workbook_protection) : nil
      else
        @workbook_protection = opts
        nil
      end
    end

    # Sets workbook protection settings.
    #
    # @param opts [Hash, Elements::MetadataRecord, nil]
    # @return [void]
    # @api public
    #: (untyped opts) -> void
    def workbook_protection=(opts)
      @workbook_protection = opts.is_a?(Elements::MetadataRecord) ? opts.to_h : opts
    end

    # Configures or queries core document properties.
    #
    # When called with no arguments, returns the {Elements::CoreProperties} value object.
    # When called with keyword options, configures properties and returns nil.
    #
    # @param opts [Hash] Core properties map.
    # @return [Elements::CoreProperties, void]
    # @api public
    #: () -> Elements::CoreProperties
    #: (**untyped opts) -> void
    def core_properties(**opts)
      return Elements::CoreProperties.new(**@core_properties) if opts.empty?

      @core_properties.merge!(opts)
      nil
    end

    # Replaces core document properties.
    #
    # @param props [Hash, Elements::CoreProperties]
    # @return [void]
    # @api public
    #: (Hash[Symbol, untyped] | Elements::CoreProperties props) -> void
    def core_properties=(props)
      h = props.respond_to?(:to_h) ? props.to_h : (props || {})
      @core_properties = h.transform_keys(&:to_sym)
      nil
    end

    # Configures or queries application document properties.
    #
    # When called with no arguments, returns the {Elements::AppProperties} value object.
    # When called with keyword options, configures properties and returns nil.
    #
    # @param opts [Hash] Application properties map.
    # @return [Elements::AppProperties, void]
    # @api public
    #: () -> Elements::AppProperties
    #: (**untyped opts) -> void
    def app_properties(**opts)
      return Elements::AppProperties.new(**@app_properties) if opts.empty?

      opts.each do |k, v|
        norm_name = k.to_sym == :application ? :app_name : k.to_sym
        @app_properties[norm_name] = v
      end
      nil
    end

    # Replaces application document properties.
    #
    # @param props [Hash, Elements::AppProperties]
    # @return [void]
    # @api public
    #: (Hash[Symbol, untyped] | Elements::AppProperties props) -> void
    def app_properties=(props)
      h = props.respond_to?(:to_h) ? props.to_h : (props || {})
      @app_properties = {}
      app_properties(**h)
      nil
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
      core_properties(**core) if core
      app_properties(**app) if app
      custom&.each { |k, v| custom_property(k.to_s, v) }
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

    # Returns custom document properties.
    #
    # @return [Hash<String, Elements::CustomProperty>]
    # @api public
    #: () -> Hash[String, Elements::CustomProperty]
    def custom_properties
      @custom_properties.to_h do |cp|
        prop = cp.is_a?(Elements::CustomProperty) ? cp : Elements::CustomProperty.new(name: (cp[:name] || cp["name"]).to_s, value: cp[:value] || cp["value"], type: cp[:type] || cp["type"] || :string)
        [prop.name.to_s, prop]
      end.freeze
    end

    # Finalizes and writes all streaming sheet contents to the destination target.
    #
    # @return [void]
    # @api public
    #: () -> untyped
    def close
      raise ArgumentError, "Workbook must contain at least one sheet (Excel limitation)" if @strict_excel_mode && @sheets.empty? && @current_sheet.nil?

      flush_current_sheet

      raise ArgumentError, "Workbook must contain at least one visible sheet (Excel limitation)" if @strict_excel_mode && @sheets.none? { |s| (s.is_a?(Hash) ? s[:state] : s.state) == :visible }

      styles_definition = {
        fonts: @style_writer.fonts.dup,
        fills: @style_writer.fills.dup,
        borders: @style_writer.borders.dup,
        xf_entries: @style_writer.xf_entries.dup,
        num_fmts: @style_writer.num_fmts.dup,
        dxfs: @dxfs || []
      }

      resolved_names = resolve_defined_names(@defined_names, @sheets)

      @workbook_properties[:active_sheet] = @active_sheet unless @active_sheet.nil?
      @workbook_properties[:active_sheet_index] = @active_sheet unless @active_sheet.nil?
      @workbook_properties[:first_sheet] = @first_sheet unless @first_sheet.nil?
      @workbook_properties[:show_sheet_tabs] = @show_sheet_tabs unless @show_sheet_tabs.nil?
      @workbook_properties[:tab_ratio] = @tab_ratio unless @tab_ratio.nil?

      wb_writer = Ooxml::WorkbookWriter.new(
        sheets: @sheets,
        shared_strings: @sst,
        shared_strings_index: @sst_index,
        styles: styles_definition,
        defined_names: resolved_names.empty? ? nil : resolved_names,
        core_properties: @core_properties.empty? ? nil : @core_properties,
        app_properties: @app_properties.empty? ? nil : @app_properties,
        custom_properties: @custom_properties.empty? ? nil : @custom_properties,
        workbook_protection: @workbook_protection,
        workbook_properties: @workbook_properties
      )
      wb_writer.write_package_parts(@zip)

      @zip.close
      @io.close if @owns_io && !@io.closed?
    ensure
      cleanup!
    end

    # Explicitly cleans up any resources if not already closed.
    #
    # @return [void]
    # @api private
    #: () -> void
    def cleanup!
      if @zip && !@zip.instance_variable_get(:@closed)
        begin
          @zip.close
        rescue StandardError
          nil
        end
      end
      return unless @owns_io && @io && !@io.closed?

      begin
        @io.close
      rescue StandardError
        nil
      end
    end

    private

    # Gives Date and Time values whose named cell style has no number format a variant of
    # that style with the default date or time format, registered on first use.
    #: (Array[untyped] values, untyped cell_styles) -> untyped
    def with_date_style_variants(values, cell_styles)
      styles_for_cells = cell_styles.is_a?(Array) ? cell_styles : nil
      values.each_with_index do |value, idx|
        next unless value.is_a?(Date) || value.is_a?(Time)

        style_name = styles_for_cells ? styles_for_cells[idx] : cell_styles
        next unless style_name.is_a?(String) || style_name.is_a?(Symbol)

        variant = DslHelpers.date_style_variant(value, style_name.to_s, @styles[style_name.to_s])
        next unless variant

        variant_name, number_format = variant
        unless @style_name_to_id.key?(variant_name)
          variant_builder = @styles[style_name.to_s].with_number_format(variant_name, number_format)
          @styles[variant_name] = variant_builder
          @style_name_to_id[variant_name] = variant_builder.register_with(@style_writer)
        end
        styles_for_cells ||= Array.new(values.size, cell_styles)
        styles_for_cells = styles_for_cells.dup if styles_for_cells.equal?(cell_styles)
        styles_for_cells[idx] = variant_name
      end
      styles_for_cells || cell_styles
    end

    def ensure_sheet_started!
      sheet if @current_sheet.nil?
    end

    #: (String? range, ?sheet: String?) -> String?
    def set_print_area(range, sheet: nil)
      sheet_name = sheet || @current_sheet || "Sheet1"
      @current_print_area = range&.to_s if sheet_name == @current_sheet
      @defined_names.reject! { |dn| dn[:name] == "_xlnm.Print_Area" && dn[:local_sheet_name] == sheet_name }
      if range
        value = "'#{sheet_name}'!#{absolute_range(range)}"
        defined_name("_xlnm.Print_Area", value, sheet: sheet_name, scope: sheet_name)
      end
      range&.to_s
    end

    #: (untyped range) -> untyped
    def absolute_range(range)
      DslHelpers.absolute_range(range)
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

      # Normalize conditional formatting rules and record dxfs in @dxfs
      unless @current_conditional_formats.empty?
        @dxfs ||= []
        @current_conditional_formats = @current_conditional_formats.map do |rule|
          next rule if rule[:dxf_id] || rule[:format_id]

          normalized = rule.dup
          dxf = normalized.delete(:dxf) || DslHelpers.extract_dxf_from_rule!(normalized)
          next normalized if dxf.nil? || dxf.empty?

          dxf_id = @dxfs.index(dxf)
          unless dxf_id
            @dxfs << dxf
            dxf_id = @dxfs.size - 1
          end
          normalized[:dxf_id] = dxf_id
          normalized
        end
      end

      start_sheet_entry unless @sheet_entry_started

      sheet_data = {
        name: @current_sheet,
        state: @current_sheet_state || :visible,
        columns: @current_columns,
        charts: @current_charts,
        hyperlinks: @current_hyperlinks,
        auto_filter: @current_auto_filter,
        filter_columns: @current_filter_columns,
        sort_state: @current_sort_state,
        data_validations: @current_data_validations,
        conditional_formats: @current_conditional_formats,
        tables: @current_tables,
        pivot_tables: @current_pivot_tables,
        sparkline_groups: @current_sparkline_groups,
        comments: @current_comments,
        threaded_comments: @current_threaded_comments,
        merge_cells: @current_merge_cells,
        freeze_pane: @current_freeze_pane,
        split_pane: @current_split_pane,
        selection: @current_selection,
        page_margins: @current_page_margins,
        page_setup: @current_page_setup,
        header_footer: @current_header_footer,
        print_options: @current_print_options,
        sheet_protection: @current_sheet_protection,
        images: @current_images,
        shapes: @current_shapes,
        sheet_properties: @current_sheet_properties,
        sheet_format: @current_sheet_format,
        sheet_view: @current_sheet_view,
        row_breaks: @current_row_breaks,
        column_breaks: @current_col_breaks,
        protected_ranges: @current_protected_ranges,
        cell_watches: @current_cell_watches,
        ignored_errors: @current_ignored_errors,
        phonetic_properties: @current_phonetic_properties,
        sheet_calc_pr: @current_sheet_calc_pr,
        background_image: @current_background_image
      }
      sheet_data[:cells] = @current_cells if @current_cells && !@current_cells.empty?
      @current_cells = nil

      temp_wb_writer = Ooxml::WorkbookWriter.new(sheets: [sheet_data])
      info = temp_wb_writer.prepare_sheet_auxiliary(sheet_data, @current_sheet_index - 1)

      @current_row_writer.finish(
        drawing_rid: info[:drawing_rid],
        picture_rid: info[:picture_rid],
        sheet_protection: @current_sheet_protection,
        protected_ranges: @current_protected_ranges.empty? ? nil : @current_protected_ranges,
        auto_filter: @current_auto_filter,
        filter_columns: @current_filter_columns.empty? ? nil : @current_filter_columns,
        sort_state: @current_sort_state,
        merge_cells: @current_merge_cells.empty? ? nil : @current_merge_cells,
        conditional_formats: @current_conditional_formats.empty? ? nil : @current_conditional_formats,
        data_validations: @current_data_validations.empty? ? nil : @current_data_validations,
        hyperlinks: info[:enriched_hyperlinks].empty? ? nil : info[:enriched_hyperlinks],
        print_options: @current_print_options.empty? ? nil : @current_print_options,
        page_margins: @current_page_margins,
        page_setup: @current_page_setup.empty? ? nil : @current_page_setup,
        header_footer: @current_header_footer.empty? ? nil : @current_header_footer,
        row_breaks: @current_row_breaks.empty? ? nil : @current_row_breaks,
        column_breaks: @current_col_breaks.empty? ? nil : @current_col_breaks,
        cell_watches: @current_cell_watches.empty? ? nil : @current_cell_watches,
        ignored_errors: @current_ignored_errors.empty? ? nil : @current_ignored_errors,
        phonetic_pr: @current_phonetic_properties,
        sheet_calc_pr: @current_sheet_calc_pr,
        tables: @current_tables.empty? ? nil : @current_tables,
        table_start_rid: info[:table_start_rid],
        legacy_drawing_rid: info[:vml_rid],
        sparkline_groups: @current_sparkline_groups.empty? ? nil : @current_sparkline_groups
      )

      @zip.finish_entry
      @sheets << sheet_data

      @current_sheet = nil
      @current_row_writer = nil
      @sheet_entry_started = false
    end
  end
end
