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
    #: String?
    attr_reader :current_sheet

    # @return [Boolean] Whether to treat strings starting with formula prefixes as plain text.
    attr_accessor :escape_formulas

    # @return [Boolean] Whether the shared strings table is enabled.
    attr_reader :shared_strings
    alias shared_strings? shared_strings

    # @return [Integer] The ZIP deflate compression level (0..9).
    attr_reader :compression_level

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
    # @param compression_level [Integer, nil] ZIP deflate compression level (0..9, or Zlib::DEFAULT_COMPRESSION).
    #   Level 6 is standard. Level 4 accelerates compression by 20-40% with an ~10% size tradeoff,
    #   recommended for high-throughput streaming exports. Level 0 disables compression (store only).
    #: (untyped target, ?strict_excel_mode: bool, ?escape_formulas: bool, ?default_font: untyped?, ?shared_strings: bool, ?compression_level: Integer | nil) -> void
    def initialize(target, strict_excel_mode: true, escape_formulas: false, default_font: nil, shared_strings: true, compression_level: Zlib::DEFAULT_COMPRESSION)
      @target = target
      @strict_excel_mode = strict_excel_mode
      @escape_formulas = escape_formulas ? true : false
      @shared_strings = shared_strings ? true : false
      @compression_level = compression_level || Zlib::DEFAULT_COMPRESSION
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
      @current_row_outline_level = nil
      @current_row_collapsed = false
      @current_row_hidden = false
      @current_cells = {}
      @styles = {} # { style_name => StyleBuilder }
      @style_writer = Ooxml::Writer.new(default_font: default_font)
      @style_name_to_id = {}
      style("__xlsxrb_date", number_format: "yyyy-mm-dd")
      style("__xlsxrb_time", number_format: "yyyy-mm-dd hh:mm:ss")
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
    # Gets or sets the active worksheet tab in the workbook window.
    #
    # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#active_sheet=} instead.
    # @param name_or_index [String, Symbol, Integer, nil] Sheet name or 0-based sheet index.
    # @return [String, Symbol, Integer, nil]
    # @api public
    #: (?String | Symbol | Integer | nil name_or_index) -> (String | Symbol | Integer | nil)
    def active_sheet(name_or_index = OMITTED)
      if name_or_index.equal?(OMITTED)
        @active_sheet
      else
        @active_sheet = name_or_index
      end
    end
    # @deprecated Use {#active_sheet} instead. Will be removed in 0.2.0.
    alias active_tab active_sheet

    # Sets the active worksheet tab in the workbook window.
    #
    # @param name_or_index [String, Symbol, Integer, nil] Sheet name or 0-based sheet index.
    # @return [String, Symbol, Integer, nil]
    # @api public
    attr_writer :active_sheet #: (String | Symbol | Integer | nil)
    # @deprecated Use {#active_sheet=} instead. Will be removed in 0.2.0.
    alias active_tab= active_sheet=

    # Gets or sets the first sheet tab visible in the sheet tab scroll area.
    #
    # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#first_sheet=} instead.
    # @param name_or_index [String, Symbol, Integer, nil] Sheet name or 0-based sheet index.
    # @return [String, Symbol, Integer, nil]
    # @api public
    #: (?String | Symbol | Integer | nil name_or_index) -> (String | Symbol | Integer | nil)
    def first_sheet(name_or_index = OMITTED)
      if name_or_index.equal?(OMITTED)
        @first_sheet
      else
        @first_sheet = name_or_index
      end
    end

    # Sets the first sheet tab visible in the sheet tab scroll area.
    #
    # @param name_or_index [String, Symbol, Integer, nil] Sheet name or 0-based sheet index.
    # @return [String, Symbol, Integer, nil]
    # @api public
    attr_writer :first_sheet #: (String | Symbol | Integer | nil)

    # Gets or sets whether sheet tabs are visible in the workbook window.
    #
    # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#show_sheet_tabs=} instead.
    # @param visible [Boolean, nil]
    # @return [Boolean, nil]
    # @api public
    #: (?bool? visible) -> bool?
    def show_sheet_tabs(visible = OMITTED)
      if visible.equal?(OMITTED)
        @show_sheet_tabs
      else
        @show_sheet_tabs = visible
      end
    end

    # Sets whether sheet tabs are visible in the workbook window.
    #
    # @param visible [Boolean, nil]
    # @return [Boolean, nil]
    # @api public
    attr_writer :show_sheet_tabs #: bool?

    # Gets or sets the ratio between the sheet tabs bar and the horizontal scroll bar (0-1000).
    #
    # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#tab_ratio=} instead.
    # @param ratio [Integer, nil]
    # @return [Integer, nil]
    # @api public
    #: (?Integer? ratio) -> Integer?
    def tab_ratio(ratio = OMITTED)
      if ratio.equal?(OMITTED)
        @tab_ratio
      else
        @tab_ratio = ratio
      end
    end

    # Sets the ratio between the sheet tabs bar and the horizontal scroll bar (0-1000).
    #
    # @param ratio [Integer, nil]
    # @return [Integer, nil]
    # @api public
    attr_writer :tab_ratio #: Integer?

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
    # @deprecated Use {#default_font} instead. Will be removed in 0.2.0.
    alias set_default_font default_font

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

      # Returns the worksheet name.
      #
      # @return [String]
      # @api public
      #: () -> String
      def name
        @sheet_name
      end

      # Returns column configurations defined for this worksheet.
      #
      # @return [Array<Elements::Column>]
      # @api public
      #: () -> Array[Elements::Column]
      def columns
        @writer.columns
      end

      # Sets or configures the workbook-wide default font.
      #
      # @param name [String, Hash, nil] Font name or font options Hash.
      # @param size [Float, Integer, nil] Font size in points.
      # @param opts [Hash] Additional font options.
      # @return [void]
      # @api public
      #: (?String | Hash[Symbol, untyped]? name, ?Float | Integer | nil size, **untyped opts) -> void
      def default_font(...)
        @writer.default_font(...)
      end
      # @deprecated Use {#default_font} instead. Will be removed in 0.2.0.
      alias set_default_font default_font

      # Returns the list of worksheet names created in this workbook so far.
      #
      # @return [Array<String>] Worksheet names.
      # @api public
      #: () -> Array[String]
      def sheet_names
        @writer.sheet_names
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
      # @deprecated Passing coordinate keyword arguments (row:, col_start:, col_end:, row_start:, row_end:) is deprecated and will be removed in 0.2.0. Use a range string (e.g. merge("A1:C1")) instead.
      # @param range [String, Hash, nil] The cell range (e.g. "A1:B2").
      # @param row [Integer, nil] 0-based row index.
      # @param col_start [Integer, String, nil] 0-based start column index or letter.
      # @param col_end [Integer, String, nil] 0-based end column index or letter.
      # @param row_start [Integer, nil] 0-based start row index.
      # @param row_end [Integer, nil] 0-based end row index.
      # @return [void]
      # @api public
      #: (?(String | Hash[Symbol, Integer | String])? range, ?row: Integer | nil, ?col_start: (Integer | String)?, ?col_end: (Integer | String)?, ?row_start: Integer | nil, ?row_end: Integer | nil) -> void
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
      # @param from_column [Integer, nil] Starting column index (0-based).
      # @param from_col [Integer, nil] Deprecated alias for from_column.
      # @param from_row [Integer] Starting row index (0-based).
      # @param to_column [Integer, nil] Ending column index (0-based).
      # @param to_col [Integer, nil] Deprecated alias for to_column.
      # @param to_row [Integer] Ending row index (0-based).
      # @param opts [Hash] Additional shape formatting options.
      # @return [void]
      # @api public
      #: (?preset: String, ?text: String?, ?from_column: Integer?, ?from_col: Integer?, ?from_row: Integer, ?to_column: Integer?, ?to_col: Integer?, ?to_row: Integer, **untyped opts) -> void
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
      # @param styles [String, Symbol, Array, Hash, nil] Deprecated. Use cell_styles: instead.
      # @param style [Integer, Symbol, String, Hash, nil] Deprecated. Use row_style: instead.
      # @return [void]
      # @api public
      #: (Array[untyped] | Hash[untyped, untyped] values, ?cell_styles: untyped, ?row_style: untyped, ?styles: untyped, ?style: untyped, ?height: Float | Integer | nil, ?hidden: bool, ?custom_height: bool, ?collapsed: bool, ?outline_level: Integer | nil) -> void
      def row(values, cell_styles: nil, row_style: nil, styles: nil, style: nil, height: nil, hidden: false, custom_height: false, collapsed: false, outline_level: nil)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        if styles
          warn "[DEPRECATION] WorksheetProxy#row keyword :styles is deprecated; use :cell_styles instead.", uplevel: 1
          cell_styles ||= styles
        end
        if style
          warn "[DEPRECATION] WorksheetProxy#row keyword :style is deprecated; use :row_style instead.", uplevel: 1
          row_style ||= style
        end

        @writer.row(values, cell_styles: cell_styles, row_style: row_style, height: height, hidden: hidden, custom_height: custom_height, collapsed: collapsed, outline_level: outline_level)
      end
      # @deprecated Use {#row} instead. Will be removed in 0.2.0.
      alias add_row row

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
      # @example Skip one row
      #   s.skip_row
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
      # @deprecated Use {#skip_rows} instead. Will be removed in 0.2.0.
      alias skip_row skip_rows

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
      # @param style [Integer, Symbol, String, Hash, nil] Column style format or index.
      # @return [void]
      # @api public
      #: (Integer | String | Symbol | Range[Integer | String] | Array[Integer | String] col_index, ?width: Float | Integer | nil, ?hidden: bool, ?best_fit: bool, ?custom_width: bool, ?outline_level: Integer | nil, ?collapsed: bool, ?style: (Integer | Symbol | String | Hash[Symbol, untyped])?) -> void
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
      # @deprecated Use {#group_columns} instead. Will be removed in 0.2.0.
      alias group_column group_columns
      # @deprecated Use {#group_columns} instead. Will be removed in 0.2.0.
      alias group_cols group_columns
      # @deprecated Use {#group_columns} instead. Will be removed in 0.2.0.
      alias group_col group_columns

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
      # @deprecated Use {#group_rows} instead. Will be removed in 0.2.0.
      alias group_row group_rows

      # Configures worksheet-level outline and grouping view properties.
      #
      # @param summary_below [Boolean, nil] Summary rows below detail rows.
      # @param summary_right [Boolean, nil] Summary columns to the right of detail columns.
      # @param show_outline_symbols [Boolean, nil] Outline symbols visible.
      # @param show_symbols [Boolean, nil] Alias for show_outline_symbols.
      # @param apply_styles [Boolean, nil] Automatic styles applied.
      # @param opts [Hash] Additional options.
      # @return [Hash]
      # @api public
      #: (?summary_below: bool?, ?summary_right: bool?, ?show_outline_symbols: bool?, ?show_symbols: bool?, ?apply_styles: bool?, **untyped opts) -> Hash[Symbol, untyped]
      def outline_properties(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.outline_properties(...)
      end
      # @deprecated Use {#outline_properties} instead. Will be removed in 0.2.0.
      alias outline_pr outline_properties

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
      #: (String | Integer cell, ?(String | Elements::Hyperlink)? url, ?display: String?, ?tooltip: String?, ?location: String?) -> void
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
      # @param col_id [Integer, String, Symbol] 0-based column index relative to auto-filter range or column letter.
      # @param filter_values [Array<String>, Hash] Values or filter specification hash.
      # @return [void]
      # @api public
      #: (Integer | String | Symbol col_id, Array[String] | Hash[Symbol, untyped] filter_values) -> void
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
      #   s.data_validation("B2:B100", type: "list", formula1: '"High,Medium,Low"')
      #
      # @example Block syntax
      #   s.data_validation("B2:B100") do |dv|
      #     dv.type :list
      #     dv.formula1 '"High,Medium,Low"'
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

      # @deprecated Use {#data_validation} instead.
      #: (?String? range, ?type: (String | Symbol)?, **untyped opts) ?{ (DataValidationBuilder) -> void } -> void
      def validate_data(...)
        warn "[DEPRECATION] `validate_data` is deprecated; use `data_validation` instead.", uplevel: 1
        data_validation(...)
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
      #     t.columns ["ID", "Name", "Total"]
      #     t.style "TableStyleMedium9"
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
      # @deprecated Use {#threaded_comment} instead. Will be removed in 0.2.0.
      alias add_threaded_comment threaded_comment

      # Adds a Pivot Table to the sheet.
      #
      # @param source_ref [String] Source data range reference (e.g. "Sheet1!A1:D100").
      # @param row_fields [Array<String>] Field names for rows.
      # @param data_fields [Array<String>] Field names for data values.
      # @param column_fields [Array<String>] Field names for columns.
      # @param col_fields [Array<String>] Deprecated alias for column_fields.
      # @param ref [String] Target top-left cell reference (default: "E1").
      # @param dest_ref [String] Deprecated alias for ref.
      # @param name [String, nil] Pivot table name.
      # @param field_names [Array<String>, nil] Override field names.
      # @param items [Array, nil] Items configuration.
      # @param opts [Hash] Additional options.
      # @return [void]
      # @api public
      #: (String source_ref, row_fields: untyped, data_fields: untyped, ?column_fields: untyped, ?col_fields: untyped, ?ref: String?, ?dest_ref: String?, ?name: String?, ?field_names: untyped, ?items: untyped, **untyped opts) -> void
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

      # Adds a single sparkline to the sheet.
      #
      # @example Add a line sparkline
      #   s.add_sparkline(location: "F1", range: "A1:E1")
      #
      # @example Add a column sparkline with high/low markers
      #   s.add_sparkline(location: "F2", range: "A2:E2", type: :column, high: true, low: true)
      #
      # @example Add a win/loss sparkline
      #   s.add_sparkline(location: "F3", range: "A3:E3", type: :win_loss)
      #
      # @param location [String, nil] Target cell for the sparkline (e.g. "F1").
      # @param range [String, nil] Data range for the sparkline (e.g. "A1:E1").
      # @param type [Symbol, String, nil] Sparkline type (:line, :column, :win_loss / :stacked).
      # @param opts [Hash] Additional sparkline options.
      # @deprecated Use {#sparkline_group} instead. Will be removed in 0.2.0.
      # @return [void]
      # @api public
      #: (?location: String?, ?range: String?, ?type: (String | Symbol)?, **untyped opts) -> void
      def add_sparkline(...)
        warn "[DEPRECATION] #add_sparkline is deprecated; use #sparkline_group instead.", uplevel: 1
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.add_sparkline(...)
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

      # Sets or gets the active sheet in the workbook window.
      # If called without arguments on the proxy, marks this sheet as the active sheet.
      #
      # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#active_sheet=} instead.
      # @param name_or_index [String, Symbol, Integer, nil]
      # @return [String, Symbol, Integer, nil]
      # @api public
      #: (?String | Symbol | Integer | nil name_or_index) -> (String | Symbol | Integer | nil)
      def active_sheet(name_or_index = @sheet_name)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.active_sheet(name_or_index)
      end
      # @deprecated Use {#active_sheet} instead. Will be removed in 0.2.0.
      alias active_tab active_sheet

      # Sets or gets the first visible sheet in the tab bar.
      #
      # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#first_sheet=} instead.
      # @param name_or_index [String, Symbol, Integer, nil]
      # @return [String, Symbol, Integer, nil]
      # @api public
      #: (?String | Symbol | Integer | nil name_or_index) -> (String | Symbol | Integer | nil)
      def first_sheet(name_or_index = @sheet_name)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.first_sheet(name_or_index)
      end

      # Sets or gets whether sheet tabs are visible in the workbook window.
      #
      # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#show_sheet_tabs=} instead.
      # @param visible [Boolean, nil]
      # @return [Boolean, nil]
      # @api public
      #: (?bool? visible) -> bool?
      def show_sheet_tabs(visible = nil)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.show_sheet_tabs(visible)
      end

      # Sets or gets the tab ratio between tab bar and horizontal scrollbar (0-1000).
      #
      # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#tab_ratio=} instead.
      # @param ratio [Integer, nil]
      # @return [Integer, nil]
      # @api public
      #: (?Integer? ratio) -> Integer?
      def tab_ratio(ratio = nil)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.tab_ratio(ratio)
      end

      # Gets or sets the sheet visibility state (:visible, :hidden, or :very_hidden).
      #
      # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#state=} instead.
      # @param new_state [Symbol, nil] The new state.
      # @return [Symbol]
      # @api public
      #: (?Symbol? new_state) -> Symbol
      def state(new_state = nil)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.state(new_state)
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

      # Gets or sets the tab color of this worksheet.
      #
      # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#tab_color=} instead.
      # @param spec [String, Hash, Symbol, nil] Tab color hex string or hash options.
      # @param opts [Hash] Color keyword arguments (e.g. theme: 1, tint: 0.5).
      # @return [String, Hash, nil]
      # @api public
      #: (?String | Symbol | Hash[Symbol, untyped] | nil spec, **untyped opts) -> (String | Hash[Symbol, untyped] | nil)
      def tab_color(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.tab_color(...)
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

      # Sets a single sheet property.
      #
      # @param name [Symbol, String] Property name.
      # @param value [Object] Property value.
      # @return [void]
      # @api public
      #: (Symbol | String name, untyped value) -> void
      def sheet_property(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.sheet_property(...)
      end

      # Sets sheet properties (e.g. tab color, page setup flags). Accepts kwargs or name/value pair.
      #
      # @example
      #   s.sheet_properties(tab_color: "FF0000")
      #
      # @param name [Symbol, String, Hash, nil] Property name or hash.
      # @param value [Object, nil] Property value.
      # @param opts [Hash] Additional options.
      # @return [void]
      # @api public
      #: (?Symbol | String | Hash[Symbol, untyped] | nil name, ?untyped value, **untyped opts) -> void
      def sheet_properties(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.sheet_properties(...)
      end

      # Adds a defined name (named range or formula) scoped to this sheet or workbook.
      # Bare cell references and coordinate keywords are automatically formatted with sheet quotes and absolute $ references.
      #
      # @param name [String] Defined name identifier.
      # @param value [Object, nil] Value or formula expression.
      # @param hidden [Boolean] Whether the defined name is hidden.
      # @param column [Integer, String, nil] Column index or letter for coordinate-based reference.
      # @param columns [Integer, String, nil] Alias for column.
      # @param col [Integer, String, nil] Deprecated alias for column.
      # @param row [Integer, Range, nil] Row or row range (1-based) for coordinate-based reference.
      # @param rows [Integer, Range, nil] Alias for row.
      # @param range [String, nil] Cell range reference.
      # @param sheet [String, nil] Reference sheet name (defaults to current sheet).
      # @param scope [String, Symbol, nil] Scope identifier (:workbook or sheet name).
      # @param local [Boolean] Deprecated alias for scope: :current.
      # @return [void]
      # @api public
      #: (String name, ?untyped value, ?hidden: bool, ?column: untyped, ?columns: untyped, ?col: untyped, ?row: untyped, ?rows: untyped, ?range: untyped, ?sheet: String?, ?scope: (String | Symbol)?, ?local: bool) -> void
      def defined_name(name, value = nil, hidden: false, column: nil, columns: nil, col: nil, row: nil, rows: nil, range: nil, sheet: nil, scope: nil, local: false)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        if col
          warn "[DEPRECATION] `col:` is deprecated; use `column:` (or `columns:`) instead.", uplevel: 1
        end
        if local
          warn "[DEPRECATION] `local:` is deprecated; use `scope:` instead.", uplevel: 1
        end
        c_arg = column || columns || col
        r_arg = row || rows
        ref_sheet = sheet || @sheet_name
        local_sheet = if scope
                        scope == :workbook ? nil : scope.to_s
                      elsif local
                        @sheet_name || sheet
                      else
                        sheet
                      end
        @writer.defined_name(name, value, sheet: local_sheet, hidden: hidden, column: c_arg, row: r_arg, range: range, ref_sheet: ref_sheet)
      end

      # Freezes rows and/or columns for scrolling.
      #
      # @example Freeze top row
      #   s.freeze_pane(row: 1)
      #
      # @example Freeze first column and top 2 rows
      #   s.freeze_pane(row: 2, column: 1)
      #
      # @param row [Integer, nil] Number of rows to freeze.
      # @param column [Integer, String, nil] Number of columns to freeze.
      # @param col [Integer, String, nil] Deprecated alias for column.
      # @return [void]
      # @api public
      #: (?row: Integer | nil, ?column: (Integer | String)?, ?col: (Integer | String)?) -> void
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
      # @param rows [String, Integer, Range, nil] Repeating row range (e.g. "1:2" or 0-based 0..1).
      # @param columns [String, Integer, Symbol, Range, nil] Repeating column range (e.g. "A:B" or 0..1).
      # @param cols [String, Integer, Symbol, Range, nil] Deprecated alias for columns.
      # @return [void]
      # @api public
      #: (?rows: (String | Integer | Range[Integer])?, ?columns: (String | Integer | Symbol | Range[untyped])?, ?cols: (String | Integer | Symbol | Range[untyped])?) -> void
      def print_titles(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.print_titles(...)
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
      # Sets workbook structure and window protection.
      #
      # @param opts [Hash] Protection options.
      # @return [void]
      # @api public
      #: (**untyped opts) -> void
      def workbook_protection(...)
        @writer.workbook_protection(...)
      end

      # @deprecated Use {#workbook_protection} instead.
      #: (**untyped opts) -> void
      def protect_workbook(...)
        warn "[DEPRECATION] `protect_workbook` is deprecated; use `workbook_protection` instead.", uplevel: 1
        workbook_protection(...)
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
      # @param active_cell [String] Cell reference (e.g. "A1") or range (e.g. "A1:D10").
      # @param sqref [String, nil] Selection range.
      # @param pane [String, Symbol, nil] Pane identifier.
      # @return [void]
      # @api public
      #: (String active_cell, ?sqref: String?, ?pane: (String | Symbol)?) -> void
      def selection(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.selection(...)
      end

      # @deprecated Use {#selection} instead.
      #: (String active_cell, ?sqref: String?, ?pane: (String | Symbol)?) -> void
      def select_cell(...)
        warn "[DEPRECATION] `select_cell` is deprecated; use `selection` instead.", uplevel: 1
        selection(...)
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
      # @param name [Symbol, String, Hash, nil] Print option name or hash.
      # @param value [Object, nil] Print option value.
      # @param opts [Hash] Additional options.
      # @return [void]
      # @api public
      #: (?Symbol | String | Hash[Symbol, untyped] | nil name, ?untyped value, **untyped opts) -> void
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

      # Sets sheet-level protection with optional password hashing.
      #
      # @param opts [Hash] Protection options.
      # @return [void]
      # @api public
      #: (**untyped opts) -> void
      def sheet_protection(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.sheet_protection(...)
      end

      # @deprecated Use {#sheet_protection} instead.
      #: (**untyped opts) -> void
      def protect_sheet(...)
        warn "[DEPRECATION] `protect_sheet` is deprecated; use `sheet_protection` instead.", uplevel: 1
        sheet_protection(...)
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
      # @param from_column [Integer, nil] Starting column index (0-based).
      # @param from_col [Integer, nil] Deprecated alias for from_column.
      # @param from_row [Integer] Starting row index (0-based).
      # @param to_column [Integer, nil] Ending column index (0-based).
      # @param to_col [Integer, nil] Deprecated alias for to_column.
      # @param to_row [Integer] Ending row index (0-based).
      # @param opts [Hash] Additional anchor and sizing options.
      # @return [void]
      # @api public
      #: (String file_data, ?ext: String, ?from_column: Integer?, ?from_col: Integer?, ?from_row: Integer, ?to_column: Integer?, ?to_col: Integer?, ?to_row: Integer, **untyped opts) -> void
      def image(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.image(...)
      end

      # Configures the worksheet background image (<picture>).
      #
      # @example Set background image from binary data
      #   sheet.background_image(png_bytes, ext: "png")
      #
      # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#background_image=} instead.
      # @param file_data [String, Hash, Elements::Image, nil] Image binary data, file path, or image options hash.
      # @param ext [String, nil] Image extension.
      # @param opts [Hash] Additional options.
      # @return [Hash, nil]
      # @api public
      #: (?untyped file_data, ?ext: String?, **untyped opts) -> Hash[Symbol, untyped]?
      def background_image(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.background_image(...)
      end
      # @deprecated Use {#background_image} instead. Will be removed in 0.2.0.
      alias background_picture background_image
      # @deprecated Use {#background_image} instead. Will be removed in 0.2.0.
      alias picture background_image

      #: (untyped val) -> untyped
      def background_image=(val)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.background_image = val
      end
      # @deprecated Use {#background_image=} instead. Will be removed in 0.2.0.
      alias background_picture= background_image=
      # @deprecated Use {#background_image=} instead. Will be removed in 0.2.0.
      alias picture= background_image=

      # Configures sheet view settings (zoom scale, grid lines visibility).
      #
      # @param name [Symbol, String, Hash, nil] View setting name or hash.
      # @param value [Object, nil] View setting value.
      # @param opts [Hash] Additional options.
      # @return [void]
      # @api public
      #: (?Symbol | String | Hash[Symbol, untyped] | nil name, ?untyped value, **untyped opts) -> void
      def sheet_view(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.sheet_view(...)
      end

      # Configures sheet format properties (e.g. default_row_height, default_col_width).
      #
      # @param name [Symbol, String, Hash, nil] Property name or hash of properties.
      # @param value [Object, nil] Property value when name is provided.
      # @param opts [Hash] Keyword arguments for sheet format properties.
      # @return [Hash[Symbol, untyped]]
      # @api public
      #: (?Symbol | String | Hash[Symbol, untyped] | nil name, ?untyped value, **untyped opts) -> Hash[Symbol, untyped]
      def sheet_format(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.sheet_format(...)
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

      # @deprecated Use {#row_break} instead.
      #: (Integer row_num) -> void
      def page_break_row(...)
        warn "[DEPRECATION] `page_break_row` is deprecated; use `row_break` instead.", uplevel: 1
        row_break(...)
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

      # @deprecated Use {#column_break} instead.
      #: (Integer | String | Symbol col_index) -> void
      def page_break_col(...)
        warn "[DEPRECATION] `page_break_col` is deprecated; use `column_break` instead.", uplevel: 1
        column_break(...)
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
      # @deprecated Use {#protected_range} instead. Will be removed in 0.2.0.
      alias add_protected_range protected_range

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
      # @deprecated Use {#cell_watch} instead. Will be removed in 0.2.0.
      alias add_cell_watch cell_watch

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

      # @deprecated Use {#ignored_error} instead. Will be removed in 0.2.0.
      #: (String sqref, ?eval_error: bool, ?two_digit_text_year: bool, ?number_stored_as_text: bool, ?formula: bool, ?formula_range: bool, ?unlocked_formula: bool, ?empty_cell_reference: bool, ?list_data_validation: bool, ?calculated_column: bool) -> void
      def ignore_error(...)
        warn "[DEPRECATION] #ignore_error is deprecated; use #ignored_error instead.", uplevel: 1
        ignored_error(...)
      end

      # @deprecated Use {#ignored_error} instead. Will be removed in 0.2.0.
      #: (String sqref, ?eval_error: bool, ?two_digit_text_year: bool, ?number_stored_as_text: bool, ?formula: bool, ?formula_range: bool, ?unlocked_formula: bool, ?empty_cell_reference: bool, ?list_data_validation: bool, ?calculated_column: bool) -> void
      def add_ignored_error(...)
        warn "[DEPRECATION] #add_ignored_error is deprecated; use #ignored_error instead.", uplevel: 1
        ignored_error(...)
      end

      # Returns the ignored error entries configured for this sheet.
      #
      # @return [Array<Hash[Symbol, untyped]>]
      # @api public
      #: () -> Array[Hash[Symbol, untyped]]
      def ignored_errors
        @writer.ignored_errors
      end

      # Gets or sets phonetic (furigana) guide properties for the sheet.
      #
      # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#phonetic_properties=} instead.
      # @api public
      #: (?Hash[Symbol, untyped] | Integer | nil spec, ?font_id: Integer?, ?type: (Symbol | String)?, ?alignment: (Symbol | String)?, **untyped opts) -> Hash[Symbol, untyped]?
      def phonetic_properties(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.phonetic_properties(...)
      end
      # @deprecated Use {#phonetic_properties} instead. Will be removed in 0.2.0.
      alias phonetic_pr phonetic_properties

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
      # @deprecated Use {#phonetic_properties=} instead. Will be removed in 0.2.0.
      alias phonetic_pr= phonetic_properties=

      # Sets or gets the internal code name of the worksheet.
      #
      # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#code_name=} instead.
      # @param val [String, Symbol, nil] The code name.
      # @return [String, nil]
      # @api public
      #: (?String | Symbol | nil val) -> String?
      def code_name(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.code_name(...)
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

      # Configures whether page setup fits content to page width/height.
      # When enabled, optionally sets fit_to_width and/or fit_to_height on page_setup.
      #
      # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#fit_to_page=} or {#page_setup} with :fit_to_width/:fit_to_height instead.
      # @param enabled [Boolean, nil] Whether to fit to page.
      # @param width [Integer, nil] Pages wide to fit to.
      # @param height [Integer, nil] Pages tall to fit to.
      # @return [Boolean, void]
      # @api public
      #: (?bool? enabled, ?width: Integer?, ?height: Integer?) -> (bool | void)
      def fit_to_page(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.fit_to_page(...)
      end
      alias fit_to_page? fit_to_page

      # Sets whether page setup fits content to page width/height.
      #
      # @param enabled [Boolean, nil] Whether to fit to page.
      # @return [Boolean]
      # @api public
      #: (bool? enabled) -> bool
      def fit_to_page=(enabled)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.fit_to_page = enabled
      end

      # Configures worksheet calculation properties (<sheetCalcPr>).
      #
      # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#sheet_calc_properties=} instead.
      # @param spec [Hash, nil] Calculation properties hash.
      # @param full_calc_on_load [Boolean, nil] Whether to recalculate when loading.
      # @return [Hash, nil]
      # @api public
      #: (?Hash[Symbol, untyped]? spec, ?full_calc_on_load: bool?, **untyped opts) -> Hash[Symbol, untyped]?
      def sheet_calc_properties(...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.sheet_calc_properties(...)
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

      # @deprecated Use {#sheet_calc_properties} instead. Will be removed in 0.2.0.
      #: (?Hash[Symbol, untyped]? spec, ?full_calc_on_load: bool?, **untyped opts) -> Hash[Symbol, untyped]?
      def sheet_calc_pr(...)
        warn "[DEPRECATION] #sheet_calc_pr is deprecated; use #sheet_calc_properties instead.", uplevel: 1
        sheet_calc_properties(...)
      end

      # @deprecated Use {#sheet_calc_properties} instead. Will be removed in 0.2.0.
      #: (?Hash[Symbol, untyped]? spec, ?full_calc_on_load: bool?, **untyped opts) -> Hash[Symbol, untyped]?
      def calc_pr(...)
        warn "[DEPRECATION] #calc_pr is deprecated; use #sheet_calc_properties instead.", uplevel: 1
        sheet_calc_properties(...)
      end

      # @deprecated Use {#sheet_calc_properties=} instead. Will be removed in 0.2.0.
      #: (Hash[Symbol, untyped]? spec) -> Hash[Symbol, untyped]?
      def sheet_calc_pr=(spec)
        warn "[DEPRECATION] #sheet_calc_pr= is deprecated; use #sheet_calc_properties= instead.", uplevel: 1
        self.sheet_calc_properties = spec
      end

      # @deprecated Use {#sheet_calc_properties=} instead. Will be removed in 0.2.0.
      #: (Hash[Symbol, untyped]? spec) -> Hash[Symbol, untyped]?
      def calc_pr=(spec)
        warn "[DEPRECATION] #calc_pr= is deprecated; use #sheet_calc_properties= instead.", uplevel: 1
        self.sheet_calc_properties = spec
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
    # @param opts [Hash] Sheet-level configuration.
    # @yield [sheet_proxy]
    # @yieldparam sheet_proxy [WorksheetProxy] The streaming sheet proxy.
    # @return [String] The sheet name.
    # @api public
    #: (?String? name, **untyped opts) ?{ (WorksheetProxy) -> void } -> untyped
    def sheet(name = nil, **opts)
      sheet_count = @sheets.size + (@current_sheet ? 1 : 0)
      name ||= "Sheet#{sheet_count + 1}"
      raise ArgumentError, "Sheet name '#{name}' must be <= 31 characters (Excel limitation)" if @strict_excel_mode && name.length > 31
      raise ArgumentError, "Sheet name '#{name}' contains invalid characters (Excel limitation)" if name.match?(Elements::Worksheet::FORBIDDEN_NAME_CHARS)
      raise ArgumentError, "Sheet name '#{name}' cannot start or end with an apostrophe (Excel limitation)" if @strict_excel_mode && (name.start_with?("'") || name.end_with?("'"))
      raise ArgumentError, "Sheet name '#{name}' is reserved by Excel (Excel limitation)" if @strict_excel_mode && name.casecmp?("History")

      existing_names = @sheets.map { |s| s.respond_to?(:name) ? s.name.downcase : s[:name].to_s.downcase }
      existing_names << @current_sheet.downcase if @current_sheet
      raise ArgumentError, "Sheet name '#{name}' is already used. Excel requires unique sheet names." if @strict_excel_mode && existing_names.include?(name.downcase)

      active = opts.delete(:active)
      state = opts.delete(:state)
      hidden = opts.delete(:hidden)
      state ||= :hidden if hidden
      state ||= :visible
      state = state.to_sym
      raise ArgumentError, "Invalid sheet state ':#{state}'. Allowed states are :visible, :hidden, :very_hidden" unless VALID_SHEET_STATES.include?(state)

      internal_sheet_setup(name, state: state)
      if active
        @active_sheet = name
        @current_sheet_view[:tab_selected] = true
      end
      opts.each { |k, v| sheet_properties(k, v) }

      yield WorksheetProxy.new(self, @current_sheet) if block_given?
      @current_sheet
    end

    # Internal: Start or switch to a named sheet (internal helper).
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
    # @param styles [String, Symbol, Array, Hash, nil] Deprecated. Use cell_styles: instead.
    # @param style [Integer, Symbol, String, Hash, nil] Deprecated. Use row_style: instead.
    # @return [void]
    # @api public
    #: (Array[untyped] | Hash[untyped, untyped] values, ?cell_styles: untyped, ?row_style: untyped, ?styles: untyped, ?style: untyped, ?height: Float | Integer | nil, ?hidden: bool, ?custom_height: bool, ?collapsed: bool, ?outline_level: Integer | nil) -> void
    def row(values, cell_styles: nil, row_style: nil, styles: nil, style: nil, height: nil, hidden: false, custom_height: false, collapsed: false, outline_level: nil)
      sheet if @current_sheet.nil?

      if styles
        warn "[DEPRECATION] StreamWriter#row keyword :styles is deprecated; use :cell_styles instead.", uplevel: 1
        cell_styles ||= styles
      end
      if style
        warn "[DEPRECATION] StreamWriter#row keyword :style is deprecated; use :row_style instead.", uplevel: 1
        row_style ||= style
      end

      row_index = @current_row_index
      DslHelpers.validate_row_bounds!(row_index, height, strict_excel_mode: @strict_excel_mode)
      @current_row_index += 1

      values = DslHelpers.normalize_row_values(values)
      cell_styles = DslHelpers.normalize_row_styles(cell_styles)

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
                       self.style(inline_name, **row_style) unless @style_name_to_id.key?(inline_name)
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
    # @deprecated Use {#row} instead. Will be removed in 0.2.0.
    alias add_row row

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
    # @example Skip one row
    #   w.skip_row
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
    # @deprecated Use {#skip_rows} instead. Will be removed in 0.2.0.
    alias skip_row skip_rows

    # Sets column formatting and properties for one or multiple columns.
    #
    # @param index [Integer, String, Range, Array] Column index (0-based) or letter ("A".."D").
    # @param width [Float, Integer, nil] Column width in character units (0 - 255).
    # @param hidden [Boolean] Whether the column is hidden.
    # @param collapsed [Boolean] Whether the column is collapsed.
    # @param style [Integer, Symbol, String, Hash, nil] Column style format or index.
    # @return [void]
    # @api public
    #: (Integer | String | Range[Integer | String] | Array[Integer | String] index, ?width: Float | Integer | nil, ?hidden: bool, ?custom_width: bool, ?outline_level: Integer | nil, ?collapsed: bool, ?style: (Integer | Symbol | String | Hash[Symbol, untyped])?) -> void
    def column(index, width: nil, hidden: false, custom_width: false, outline_level: nil, collapsed: false, style: nil)
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
          if style_id
            existing[:style] = style_id
            @current_row_writer&.set_column_style(idx, style_id)
          end
        else
          entry = { index: idx, width: width, hidden: hidden, custom_width: custom_width || !width.nil?, outline_level: outline_level }
          entry[:collapsed] = true if collapsed
          if style_id
            entry[:style] = style_id
            @current_row_writer&.set_column_style(idx, style_id)
          end
          @current_columns << entry
        end
      end
      @current_columns.sort_by! { |c| c[:index] }
    end

    # Returns column configurations defined for the current worksheet.
    #
    # @return [Array<Elements::Column>]
    # @api public
    #: () -> Array[Elements::Column]
    def columns
      sheet if @current_sheet.nil?
      @current_columns.map do |c|
        if c.is_a?(Elements::Column)
          c
        else
          idx = c[:index]
          style_idx = c[:style] || c[:style_index]
          unmapped = c.except(:index, :width, :hidden, :custom_width, :outline_level, :style, :style_index)
          unmapped[:collapsed] = true if c[:collapsed]
          unmapped[:best_fit] = true if c[:best_fit]
          Elements::Column.new(
            index: idx,
            width: c[:width],
            hidden: c[:hidden] ? true : false,
            custom_width: c[:custom_width] ? true : false,
            outline_level: c[:outline_level],
            style_index: style_idx,
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
    # @deprecated Use {#group_columns} instead. Will be removed in 0.2.0.
    alias group_column group_columns
    # @deprecated Use {#group_columns} instead. Will be removed in 0.2.0.
    alias group_cols group_columns
    # @deprecated Use {#group_columns} instead. Will be removed in 0.2.0.
    alias group_col group_columns

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
    # @deprecated Use {#group_rows} instead. Will be removed in 0.2.0.
    alias group_row group_rows

    # Configures worksheet-level outline and grouping view properties.
    #
    # @example Show summary row below detail rows
    #   sheet.outline_properties(summary_below: true, summary_right: true)
    #
    # @param summary_below [Boolean, nil] Flag indicating if summary rows appear below detail rows.
    # @param summary_right [Boolean, nil] Flag indicating if summary columns appear to the right of detail columns.
    # @param show_outline_symbols [Boolean, nil] Flag indicating whether outline symbols are visible.
    # @param show_symbols [Boolean, nil] Alias for show_outline_symbols.
    # @param apply_styles [Boolean, nil] Flag indicating whether automatic styles apply to outlines.
    # @param opts [Hash] Additional options.
    # @return [Hash]
    # @api public
    #: (?summary_below: bool?, ?summary_right: bool?, ?show_outline_symbols: bool?, ?show_symbols: bool?, ?apply_styles: bool?, **untyped opts) -> Hash[Symbol, untyped]
    def outline_properties(summary_below: nil, summary_right: nil, show_outline_symbols: nil, show_symbols: nil, apply_styles: nil, **opts)
      sheet if @current_sheet.nil?

      props = (@current_sheet_properties[:outline_properties] || {}).dup
      props[:summary_below] = summary_below unless summary_below.nil?
      props[:summary_right] = summary_right unless summary_right.nil?
      symbols = show_outline_symbols.nil? ? show_symbols : show_outline_symbols
      props[:show_outline_symbols] = symbols unless symbols.nil?
      props[:apply_styles] = apply_styles unless apply_styles.nil?
      props.merge!(opts)
      @current_sheet_properties[:outline_properties] = props
      @current_sheet_properties[:summary_below] = props[:summary_below] unless props[:summary_below].nil?
      @current_sheet_properties[:summary_right] = props[:summary_right] unless props[:summary_right].nil?
      props
    end
    # @deprecated Use {#outline_properties} instead. Will be removed in 0.2.0.
    alias outline_pr outline_properties

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
    # Sets filter criteria on an auto-filter column.
    #
    # @param col_id [Integer, String, Symbol] 0-based column index relative to auto-filter or column letter ("A", :B).
    # @param filter [Hash] Filter definition.
    # @return [void]
    # @api public
    #: (Integer | String | Symbol col_id, Hash[untyped, untyped] filter) -> void
    def filter_column(col_id, filter)
      sheet if @current_sheet.nil?
      col_id = Elements::Cell._column_index(col_id) if col_id.is_a?(String) || col_id.is_a?(Symbol)
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

    # @deprecated Use {#data_validation} instead.
    #: (?String? sqref, **untyped opts) ?{ (DataValidationBuilder) -> void } -> void
    def validate_data(...)
      warn "[DEPRECATION] `validate_data` is deprecated; use `data_validation` instead.", uplevel: 1
      data_validation(...)
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
    # @param column_fields [Array<Integer>] 0-based field indices for columns.
    # @param col_fields [Array<Integer>] Deprecated alias for column_fields.
    # @param ref [String] Top-left destination cell (default: "E1").
    # @param dest_ref [String] Deprecated alias for ref.
    # @param name [String, nil] Pivot table name.
    # @param field_names [Array<String>, nil] Override field names.
    # @param items [Array, nil] Items configuration.
    # @param opts [Hash] Additional options.
    # @return [void]
    # @api public
    #: (untyped source_ref, row_fields: untyped, data_fields: untyped, ?column_fields: untyped, ?col_fields: untyped, ?ref: untyped, ?dest_ref: untyped, ?name: untyped, ?field_names: untyped, ?items: untyped, **untyped opts) -> void
    def pivot_table(source_ref, row_fields:, data_fields:, column_fields: nil, col_fields: nil, ref: nil, dest_ref: nil, name: nil, field_names: nil, items: nil, **opts)
      sheet if @current_sheet.nil?
      if col_fields && column_fields.nil?
        warn "[DEPRECATION] `col_fields:` is deprecated; use `column_fields:` instead.", uplevel: 1
      end
      if dest_ref && ref.nil?
        warn "[DEPRECATION] `dest_ref:` is deprecated; use `ref:` instead.", uplevel: 1
      end
      c_fields = column_fields || col_fields || []
      target_ref = ref || dest_ref || "E1"
      @current_pivot_tables ||= []
      entry = {
        source_ref: source_ref, row_fields: row_fields,
        data_fields: data_fields, column_fields: c_fields,
        col_fields: c_fields, ref: target_ref,
        dest_ref: target_ref, name: name,
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
      @current_comments << { cell: cell, text: text, author: author }
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
      @current_threaded_comments << {
        cell: cell,
        text: text,
        author: author,
        timestamp: timestamp,
        id: id,
        parent_id: parent_id,
        done: done ? true : false,
        user_id: user_id,
        provider_id: provider_id
      }.compact
    end
    # @deprecated Use {#threaded_comment} instead. Will be removed in 0.2.0.
    alias add_threaded_comment threaded_comment

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

    # Adds a single sparkline to the sheet.
    #
    # @example Add a line sparkline
    #   writer.add_sparkline(location: "F1", range: "A1:E1")
    #
    # @example Add a column sparkline with high/low markers
    #   writer.add_sparkline(location: "F2", range: "A2:E2", type: :column, high: true, low: true)
    #
    # @example Add a win/loss sparkline
    #   writer.add_sparkline(location: "F3", range: "A3:E3", type: :win_loss)
    #
    # @param location [String, nil] Target cell for the sparkline (e.g. "F1").
    # @param range [String, nil] Data range for the sparkline (e.g. "A1:E1").
    # @param type [Symbol, String, nil] Sparkline type (:line, :column, :win_loss / :stacked).
    # @param opts [Hash] Additional sparkline options.
    # @deprecated Use {#sparkline_group} instead. Will be removed in 0.2.0.
    # @return [void]
    # @api public
    #: (?location: String?, ?range: String?, ?type: (String | Symbol)?, **untyped opts) -> void
    def add_sparkline(location: nil, range: nil, type: nil, **opts)
      warn "[DEPRECATION] #add_sparkline is deprecated; use #sparkline_group instead.", uplevel: 1
      loc = location || opts.delete(:location_ref)
      rng = range || opts.delete(:data_ref) || opts.delete(:data)
      raise ArgumentError, "Missing required location: parameter for sparkline" unless loc
      raise ArgumentError, "Missing required range: parameter for sparkline" unless rng

      mapped_type = case type&.to_s
                    when "win_loss" then "stacked"
                    when nil, "" then nil
                    else type.to_s
                    end
      sparkline_group(sparklines: [{ data_ref: rng.to_s, location_ref: loc.to_s }], type: mapped_type, **opts)
    end

    # Merges a range of cells into a single cell.
    #
    # @example
    #   sheet.merge("A1:C1")
    #
    # @deprecated Passing coordinate keyword arguments (row:, col_start:, col_end:, row_start:, row_end:) is deprecated and will be removed in 0.2.0. Use a range string (e.g. merge("A1:C1")) instead.
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
      canonical = DslHelpers.normalize_merge_range(
        range,
        row: row, col_start: col_start, col_end: col_end,
        row_start: row_start, row_end: row_end,
        strict_excel_mode: @strict_excel_mode
      )
      @current_merge_cells << canonical unless @current_merge_cells.include?(canonical)
    end

    # Freezes window panes at the given row and column.
    #
    # @param row [Integer] Number of rows to freeze.
    # @param column [Integer, String, nil] Number of columns to freeze.
    # @param col [Integer, String, nil] Deprecated alias for column.
    # @return [void]
    # @api public
    #: (?row: Integer, ?column: (Integer | String)?, ?col: (Integer | String)?) -> void
    def freeze_pane(row: 0, column: nil, col: nil)
      if col && column.nil?
        warn "[DEPRECATION] `col:` is deprecated; use `column:` instead.", uplevel: 1
      end
      col_val = Elements::Cell._column_index(column || col || 0)
      sheet if @current_sheet.nil?
      @current_freeze_pane = { row: row, col: col_val, column: col_val }
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
    # @param active_cell [String] Active cell reference (e.g. "A1") or range (e.g. "A1:D10").
    # @param sqref [String, nil] Selection range.
    # @param pane [String, Symbol, nil] Target pane.
    # @return [void]
    # @api public
    #: (String active_cell, ?sqref: String?, ?pane: (String | Symbol)?) -> void
    def selection(active_cell, sqref: nil, pane: nil)
      sheet if @current_sheet.nil?
      if sqref.nil? && active_cell.is_a?(String) && active_cell.include?(":")
        sqref = active_cell
        active_cell = sqref.split(":").first
      else
        sqref ||= active_cell
      end
      @current_selection = { active_cell: active_cell, sqref: sqref }
      @current_selection[:pane] = pane if pane
    end

    # @deprecated Use {#selection} instead.
    #: (String active_cell, ?sqref: String?, ?pane: (String | Symbol)?) -> void
    def select_cell(...)
      warn "[DEPRECATION] `select_cell` is deprecated; use `selection` instead.", uplevel: 1
      selection(...)
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
      @current_page_margins = DslHelpers.normalize_page_margins(
        left: left, right: right, top: top, bottom: bottom, header: header, footer: footer
      )
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
      return unless (opts.key?(:fit_to_width) || opts.key?(:fit_to_height)) && !@current_sheet_properties.key?(:fit_to_page)

      @current_sheet_properties[:fit_to_page] = true
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

    # Configures print options (e.g. grid_lines: true).
    #
    # @param name [Symbol, String, Hash, nil] Option name or hash of options.
    # @param value [Object, nil] Option value.
    # @param opts [Hash] Additional print options.
    # @return [void]
    # @api public
    #: (?Symbol | String | Hash[Symbol, untyped] | nil name, ?untyped value, **untyped opts) -> void
    def print_options(name = nil, value = nil, **opts)
      sheet if @current_sheet.nil?
      if name.is_a?(Hash)
        opts = name.merge(opts)
        name = nil
      end
      @current_print_options[name.to_sym] = value if name && !value.nil?
      opts.each { |k, v| @current_print_options[k.to_sym] = v }
    end

    # Sets sheet-level protection with optional password hashing.
    #
    # @param opts [Hash] Protection options.
    # @return [void]
    # @api public
    #: (**untyped opts) -> void
    def sheet_protection(**opts)
      sheet if @current_sheet.nil?
      @current_sheet_protection = DslHelpers.normalize_protection_options(opts)
    end

    # @deprecated Use {#sheet_protection} instead.
    #: (**untyped opts) -> void
    def protect_sheet(...)
      warn "[DEPRECATION] `protect_sheet` is deprecated; use `sheet_protection` instead.", uplevel: 1
      sheet_protection(...)
    end

    # Inserts an embedded image into the active sheet.
    #
    # @param file_data [String] Binary image data.
    # @param ext [String] File extension ("png", "jpeg", etc.).
    # @param from_column [Integer, nil] Top-left starting column.
    # @param from_col [Integer, nil] Deprecated alias for from_column.
    # @param from_row [Integer] Top-left starting row.
    # @param to_column [Integer, nil] Bottom-right ending column.
    # @param to_col [Integer, nil] Deprecated alias for to_column.
    # @param to_row [Integer] Bottom-right ending row.
    # @param opts [Hash] Additional options.
    # @return [void]
    # @api public
    #: (String file_data, ?ext: ::String, ?from_column: ::Integer?, ?from_col: ::Integer?, ?from_row: ::Integer, ?to_column: ::Integer?, ?to_col: ::Integer?, ?to_row: ::Integer, **untyped opts) -> void
    def image(file_data, ext: "png", from_column: nil, from_col: nil, from_row: 0, to_column: nil, to_col: nil, to_row: 10, **opts)
      if from_col && from_column.nil?
        warn "[DEPRECATION] `from_col:` is deprecated; use `from_column:` instead.", uplevel: 1
      end
      if to_col && to_column.nil?
        warn "[DEPRECATION] `to_col:` is deprecated; use `to_column:` instead.", uplevel: 1
      end
      fc = from_column || from_col || 0
      tc = to_column || to_col || 5
      sheet if @current_sheet.nil?
      img = { file_data: file_data, ext: ext, from_column: fc, from_col: fc, from_row: from_row, to_column: tc, to_col: tc, to_row: to_row }
      img.merge!(opts)
      @current_images << img
    end

    # Configures the worksheet background image (<picture>).
    #
    # @example Set background image from binary data
    #   sheet.background_image(png_bytes, ext: "png")
    #
    # @example Set background image from file path
    #   sheet.background_image("path/to/background.png")
    #
    # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#background_image=} instead.
    # @param file_data [String, Hash, Elements::Image, nil] Image binary data, file path, or image options hash.
    # @param ext [String, nil] Image extension ("png", "jpeg", "jpg", "gif", etc.).
    # @param opts [Hash] Additional image options.
    # @return [Hash, nil]
    # @api public
    #: (?untyped file_data, ?ext: ::String?, **untyped opts) -> Hash[Symbol, untyped]?
    def background_image(file_data = nil, ext: nil, **opts)
      return @current_background_image if file_data.nil? && ext.nil? && opts.empty?

      sheet if @current_sheet.nil?

      data = file_data
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

      final_ext = (ext || inferred_ext || opts[:ext] || "png").to_s.delete_prefix(".").downcase
      mime = case final_ext
             when "png" then "image/png"
             when "jpg", "jpeg" then "image/jpeg"
             when "gif" then "image/gif"
             when "bmp" then "image/bmp"
             when "emf" then "image/x-emf"
             when "wmf" then "image/x-wmf"
             else "application/octet-stream"
             end

      @current_background_image = {
        data: data&.b,
        file_data: data&.b,
        ext: final_ext,
        filename: "background.#{final_ext}",
        content_type: mime
      }.merge(opts).freeze
    end
    # @deprecated Use {#background_image} instead. Will be removed in 0.2.0.
    alias background_picture background_image
    # @deprecated Use {#background_image} instead. Will be removed in 0.2.0.
    alias picture background_image

    #: (untyped val) -> untyped
    def background_image=(val)
      sheet if @current_sheet.nil?
      if val.nil?
        @current_background_image = nil
      else
        background_image(val)
      end
      val
    end
    # @deprecated Use {#background_image=} instead. Will be removed in 0.2.0.
    alias background_picture= background_image=
    # @deprecated Use {#background_image=} instead. Will be removed in 0.2.0.
    alias picture= background_image=

    # Inserts a drawing shape.
    #
    # @param preset [String] Preset shape name.
    # @param text [String, nil] Label text.
    # @param from_column [Integer, nil] Top-left column.
    # @param from_col [Integer, nil] Deprecated alias for from_column.
    # @param from_row [Integer] Top-left row.
    # @param to_column [Integer, nil] Bottom-right column.
    # @param to_col [Integer, nil] Deprecated alias for to_column.
    # @param to_row [Integer] Bottom-right row.
    # @param opts [Hash] Additional options.
    # @return [void]
    # @api public
    #: (**untyped opts) -> void
    def shape(preset: "rect", text: nil, from_column: nil, from_col: nil, from_row: 0, to_column: nil, to_col: nil, to_row: 5, **opts)
      if from_col && from_column.nil?
        warn "[DEPRECATION] `from_col:` is deprecated; use `from_column:` instead.", uplevel: 1
      end
      if to_col && to_column.nil?
        warn "[DEPRECATION] `to_col:` is deprecated; use `to_column:` instead.", uplevel: 1
      end
      fc = from_column || from_col || 0
      tc = to_column || to_col || 5
      sheet if @current_sheet.nil?
      shape_data = { preset: preset, text: text, from_column: fc, from_col: fc, from_row: from_row, to_column: tc, to_col: tc, to_row: to_row }
      shape_data[:name] = opts.delete(:name) || "Shape #{@current_shapes.size + 1}"
      shape_data.merge!(opts)
      @current_shapes << shape_data
    end

    # Gets or sets the sheet visibility state (:visible, :hidden, or :very_hidden).
    #
    # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#state=} instead.
    # @param new_state [Symbol, nil] The new state.
    # @return [Symbol]
    # @api public
    #: (?Symbol? new_state) -> Symbol
    def state(new_state = nil)
      sheet if @current_sheet.nil?
      if new_state
        sheet_properties(:state, new_state)
        @current_sheet_state
      else
        @current_sheet_state || :visible
      end
    end

    # Sets the sheet visibility state (:visible, :hidden, or :very_hidden).
    #
    # @param new_state [Symbol, String] The new state.
    # @return [Symbol]
    # @api public
    #: (Symbol | String new_state) -> Symbol
    def state=(new_state)
      sheet if @current_sheet.nil?
      sheet_properties(:state, new_state)
      @current_sheet_state
    end

    # Gets or sets the tab color of the current worksheet.
    #
    # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#tab_color=} instead.
    # @param spec [String, Hash, Symbol, nil] Tab color hex string or hash options.
    # @param opts [Hash] Color keyword arguments (e.g. theme: 1, tint: 0.5).
    # @return [String, Hash, nil]
    # @api public
    #: (?String | Symbol | Hash[Symbol, untyped] | nil spec, **untyped opts) -> (String | Hash[Symbol, untyped] | nil)
    def tab_color(spec = OMITTED, **opts)
      sheet if @current_sheet.nil?
      if spec.equal?(OMITTED) && opts.empty?
        @current_sheet_properties[:tab_color]
      else
        color = if opts.empty?
                  spec
                else
                  spec.is_a?(Hash) ? spec.merge(opts) : opts
                end
        @current_sheet_properties[:tab_color] = color
      end
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

    # Sets a single sheet property.
    #
    # @param name [Symbol, String] Property name.
    # @param value [Object] Property value.
    # @return [void]
    # @api public
    #: (Symbol | String name, untyped value) -> void
    def sheet_property(name, value)
      sheet_properties(name, value)
    end

    # Sets worksheet properties. Accepts name/value pair or keyword arguments.
    #
    # @param name [Symbol, String, Hash, nil] Property name or hash of properties.
    # @param value [Object, nil] Property value when name is provided.
    # @param opts [Hash] Additional sheet properties.
    # @return [void]
    # @api public
    #: (?Symbol | String | Hash[Symbol, untyped] | nil name, ?untyped value, **untyped opts) -> void
    def sheet_properties(name = nil, value = nil, **opts)
      sheet if @current_sheet.nil?
      if name.is_a?(Hash)
        opts = name.merge(opts)
        name = nil
      end
      if name && !value.nil?
        sym = name.to_sym
        if sym == :state
          val = value.to_sym
          raise ArgumentError, "Invalid sheet state ':#{val}'. Allowed states are :visible, :hidden, :very_hidden" unless VALID_SHEET_STATES.include?(val)

          @current_sheet_state = val
          return
        elsif sym == :hidden
          @current_sheet_state = value ? :hidden : :visible
          return
        end
        @current_sheet_properties[name.to_sym] = value
      end
      opts.each { |k, v| sheet_properties(k, v) }
    end

    # Sets a sheet view property (e.g. zoom_scale: 120).
    #
    # @param name [Symbol, String, Hash, nil] View property name or hash of properties.
    # @param value [Object, nil] View property value when name is provided.
    # @param opts [Hash] Additional view properties.
    # @return [void]
    # @api public
    #: (?Symbol | String | Hash[Symbol, untyped] | nil name, ?untyped value, **untyped opts) -> void
    def sheet_view(name = nil, value = nil, **opts)
      sheet if @current_sheet.nil?
      if name.is_a?(Hash)
        opts = name.merge(opts)
        name = nil
      end
      @current_sheet_view[name.to_sym] = value if name && !value.nil?
      opts.each { |k, v| @current_sheet_view[k.to_sym] = v }
    end

    # Configures sheet format properties (e.g. default_row_height, default_column_width).
    #
    # @param name [Symbol, String, Hash, nil] Property name or hash of properties.
    # @param value [Object, nil] Property value when name is provided.
    # @param opts [Hash] Keyword arguments for sheet format properties.
    # @return [Hash[Symbol, untyped]]
    # @api public
    #: (?Symbol | String | Hash[Symbol, untyped] | nil name, ?untyped value, **untyped opts) -> Hash[Symbol, untyped]
    def sheet_format(name = nil, value = nil, **opts)
      sheet if @current_sheet.nil?
      if name.is_a?(Hash)
        @current_sheet_format.merge!(name)
      elsif name && value
        @current_sheet_format[name.to_sym] = value
      end
      if opts.key?(:default_col_width) && !opts.key?(:default_column_width)
        warn "[DEPRECATION] `default_col_width:` is deprecated; use `default_column_width:` instead.", uplevel: 1
        opts[:default_column_width] = opts[:default_col_width]
      end
      if opts.key?(:base_col_width) && !opts.key?(:base_column_width)
        warn "[DEPRECATION] `base_col_width:` is deprecated; use `base_column_width:` instead.", uplevel: 1
        opts[:base_column_width] = opts[:base_col_width]
      end
      if opts.key?(:outline_level_col) && !opts.key?(:outline_level_column)
        warn "[DEPRECATION] `outline_level_col:` is deprecated; use `outline_level_column:` instead.", uplevel: 1
        opts[:outline_level_column] = opts[:outline_level_col]
      end
      @current_sheet_format.merge!(opts) unless opts.empty?
      @current_sheet_format
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

    # @deprecated Use {#row_break} instead.
    #: (Integer row_num) -> void
    def page_break_row(row_num)
      warn "[DEPRECATION] `page_break_row` is deprecated; use `row_break` instead.", uplevel: 1
      row_break(row_num)
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

    # @deprecated Use {#column_break} instead.
    #: (Integer | String | Symbol col_index) -> void
    def page_break_col(col_index)
      warn "[DEPRECATION] `page_break_col` is deprecated; use `column_break` instead.", uplevel: 1
      column_break(col_index)
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
    # @deprecated Use {#protected_range} instead. Will be removed in 0.2.0.
    alias add_protected_range protected_range

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
    # @deprecated Use {#cell_watch} instead. Will be removed in 0.2.0.
    alias add_cell_watch cell_watch

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

    # @deprecated Use {#ignored_error} instead. Will be removed in 0.2.0.
    #: (String sqref, ?eval_error: bool, ?two_digit_text_year: bool, ?number_stored_as_text: bool, ?formula: bool, ?formula_range: bool, ?unlocked_formula: bool, ?empty_cell_reference: bool, ?list_data_validation: bool, ?calculated_column: bool) -> void
    def ignore_error(sqref, **opts)
      warn "[DEPRECATION] #ignore_error is deprecated; use #ignored_error instead.", uplevel: 1
      ignored_error(sqref, **opts)
    end

    # @deprecated Use {#ignored_error} instead. Will be removed in 0.2.0.
    #: (String sqref, ?eval_error: bool, ?two_digit_text_year: bool, ?number_stored_as_text: bool, ?formula: bool, ?formula_range: bool, ?unlocked_formula: bool, ?empty_cell_reference: bool, ?list_data_validation: bool, ?calculated_column: bool) -> void
    def add_ignored_error(sqref, **opts)
      warn "[DEPRECATION] #add_ignored_error is deprecated; use #ignored_error instead.", uplevel: 1
      ignored_error(sqref, **opts)
    end

    # Returns the ignored error entries configured for the current sheet.
    #
    # @return [Array<Hash[Symbol, untyped]>]
    # @api public
    #: () -> Array[Hash[Symbol, untyped]]
    def ignored_errors
      @current_ignored_errors.dup
    end

    # Gets or sets phonetic (furigana) guide properties for the current sheet.
    #
    # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#phonetic_properties=} instead.
    # @overload phonetic_properties()
    #   Returns the current phonetic properties.
    #   @return [Hash, nil]
    # @overload phonetic_properties(spec = nil, font_id: nil, type: nil, alignment: nil)
    #   Sets phonetic properties for the current sheet.
    #   @param spec [Hash, Integer, nil]
    #   @param font_id [Integer, nil]
    #   @param type [Symbol, String, nil]
    #   @param alignment [Symbol, String, nil]
    #   @return [Hash, nil]
    # @api public
    #: (?Hash[Symbol, untyped] | Integer | nil spec, ?font_id: Integer?, ?type: (Symbol | String)?, ?alignment: (Symbol | String)?, **untyped opts) -> Hash[Symbol, untyped]?
    def phonetic_properties(spec = OMITTED, font_id: nil, type: nil, alignment: nil, **opts)
      sheet if @current_sheet.nil?
      if spec.equal?(OMITTED) && font_id.nil? && type.nil? && alignment.nil? && opts.empty?
        @current_phonetic_properties
      elsif !spec.equal?(OMITTED) && spec.nil? && font_id.nil? && type.nil? && alignment.nil? && opts.empty?
        @current_phonetic_properties = nil
      else
        props = if spec.is_a?(Hash)
                  spec.merge(opts)
                elsif spec.is_a?(Integer)
                  { font_id: spec, type: type, alignment: alignment }.merge(opts)
                else
                  { font_id: font_id, type: type, alignment: alignment }.merge(opts)
                end
        props.compact!
        @current_phonetic_properties = WorksheetBuilder.normalize_phonetic_properties(props)
      end
    end
    # @deprecated Use {#phonetic_properties} instead. Will be removed in 0.2.0.
    alias phonetic_pr phonetic_properties

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
    # @deprecated Use {#phonetic_properties=} instead. Will be removed in 0.2.0.
    alias phonetic_pr= phonetic_properties=

    # Sets or gets the internal code name of the current worksheet (or workbook if no sheet is open).
    #
    # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#code_name=} instead.
    # @param val [String, Symbol, nil] The code name.
    # @return [String, nil]
    # @api public
    #: (?String | Symbol | nil val) -> String?
    def code_name(val = OMITTED)
      if @current_sheet.nil?
        if val.equal?(OMITTED)
          @workbook_properties[:code_name]
        else
          @workbook_properties[:code_name] = val&.to_s
        end
      elsif val.equal?(OMITTED)
        @current_sheet_properties[:code_name]
      else
        @current_sheet_properties[:code_name] = val&.to_s
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

    # Configures whether page setup fits content to page width/height for the current sheet.
    # When enabled, optionally sets fit_to_width and/or fit_to_height on page_setup.
    #
    # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#fit_to_page=} or {#page_setup} with :fit_to_width/:fit_to_height instead.
    # @param enabled [Boolean, nil] Whether to fit to page.
    # @param width [Integer, nil] Pages wide to fit to.
    # @param height [Integer, nil] Pages tall to fit to.
    # @return [Boolean, void]
    # @api public
    #: (?bool? enabled, ?width: Integer?, ?height: Integer?) -> (bool | void)
    def fit_to_page(enabled = OMITTED, width: nil, height: nil)
      sheet if @current_sheet.nil?
      if enabled.equal?(OMITTED) && width.nil? && height.nil?
        return @current_sheet_properties[:fit_to_page] ? true : false
      end

      @current_sheet_properties[:fit_to_page] = if enabled.equal?(OMITTED)
                                                  true
                                                else
                                                  enabled ? true : false
                                                end
      page_setup(fit_to_width: width) if width
      page_setup(fit_to_height: height) if height
    end
    alias fit_to_page? fit_to_page

    # Sets whether page setup fits content to page width/height for the current sheet.
    #
    # @param enabled [Boolean, nil] Whether to fit to page.
    # @return [Boolean]
    # @api public
    #: (bool? enabled) -> bool
    def fit_to_page=(enabled)
      sheet if @current_sheet.nil?
      @current_sheet_properties[:fit_to_page] = enabled ? true : false
    end

    # Configures worksheet calculation properties (<sheetCalcPr>) for the current sheet.
    #
    # @example Instruct Excel to recalculate formulas when loading sheet
    #   w.sheet_calc_properties = { full_calc_on_load: true }
    #
    # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#sheet_calc_properties=} instead.
    # @param spec [Hash, nil] Calculation properties hash.
    # @param full_calc_on_load [Boolean, nil] Whether to recalculate when loading.
    # @return [Hash, nil]
    # @api public
    #: (?Hash[Symbol, untyped]? spec, ?full_calc_on_load: bool?, **untyped opts) -> Hash[Symbol, untyped]?
    def sheet_calc_properties(spec = OMITTED, full_calc_on_load: nil, **opts)
      sheet if @current_sheet.nil?
      if spec.equal?(OMITTED) && full_calc_on_load.nil? && opts.empty?
        return @current_sheet_calc_pr || (@current_sheet_properties[:full_calc_on_load] ? { full_calc_on_load: true } : nil)
      end

      warn "[DEPRECATION] Calling #sheet_calc_properties with arguments is deprecated and will be removed in 0.2.0. Use #sheet_calc_properties= instead.", uplevel: 1
      if !spec.equal?(OMITTED) && spec.nil? && full_calc_on_load.nil? && opts.empty?
        @current_sheet_properties.delete(:full_calc_on_load)
        return @current_sheet_calc_pr = nil
      end

      pr = spec.is_a?(Hash) ? spec.dup : {}
      pr[:full_calc_on_load] = full_calc_on_load unless full_calc_on_load.nil?
      pr.merge!(opts) unless opts.empty?
      @current_sheet_properties[:full_calc_on_load] = pr[:full_calc_on_load] unless pr[:full_calc_on_load].nil?
      @current_sheet_calc_pr = pr
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

    # @deprecated Use {#sheet_calc_properties} instead. Will be removed in 0.2.0.
    #: (?Hash[Symbol, untyped]? spec, ?full_calc_on_load: bool?, **untyped opts) -> Hash[Symbol, untyped]?
    def sheet_calc_pr(spec = OMITTED, full_calc_on_load: nil, **opts)
      warn "[DEPRECATION] #sheet_calc_pr is deprecated; use #sheet_calc_properties instead.", uplevel: 1
      sheet_calc_properties(spec, full_calc_on_load: full_calc_on_load, **opts)
    end

    # @deprecated Use {#sheet_calc_properties} instead. Will be removed in 0.2.0.
    #: (?Hash[Symbol, untyped]? spec, ?full_calc_on_load: bool?, **untyped opts) -> Hash[Symbol, untyped]?
    def calc_pr(spec = OMITTED, full_calc_on_load: nil, **opts)
      warn "[DEPRECATION] #calc_pr is deprecated; use #sheet_calc_properties instead.", uplevel: 1
      sheet_calc_properties(spec, full_calc_on_load: full_calc_on_load, **opts)
    end

    # @deprecated Use {#sheet_calc_properties=} instead. Will be removed in 0.2.0.
    #: (Hash[Symbol, untyped]? spec) -> Hash[Symbol, untyped]?
    def sheet_calc_pr=(spec)
      warn "[DEPRECATION] #sheet_calc_pr= is deprecated; use #sheet_calc_properties= instead.", uplevel: 1
      self.sheet_calc_properties = spec
    end

    # @deprecated Use {#sheet_calc_properties=} instead. Will be removed in 0.2.0.
    #: (Hash[Symbol, untyped]? spec) -> Hash[Symbol, untyped]?
    def calc_pr=(spec)
      warn "[DEPRECATION] #calc_pr= is deprecated; use #sheet_calc_properties= instead.", uplevel: 1
      self.sheet_calc_properties = spec
    end

    # --- Workbook-Level Methods ---

    # Adds a defined name.
    #
    # @param name [String] The defined name.
    # @param value [Object, nil] The formula or value expression.
    # @param sheet [String, nil] Reference sheet name.
    # @param scope [String, Symbol, nil] Local scope sheet name or :workbook.
    # @param hidden [Boolean] Whether the defined name is hidden.
    # @param column [Integer, String, nil] Column index or letter for coordinate-based reference.
    # @param columns [Integer, String, nil] Alias for column.
    # @param col [Integer, String, nil] Deprecated alias for column.
    # @param row [Integer, Range, nil] Row or row range (1-based) for coordinate-based reference.
    # @param rows [Integer, Range, nil] Alias for row.
    # @param range [String, nil] Cell range reference.
    # @param ref_sheet [String, nil] Deprecated alias for sheet.
    # @param local [Boolean] Deprecated alias for scope.
    # @return [void]
    # @api public
    #: (String name, ?untyped value, ?sheet: String?, ?scope: (String | Symbol)?, ?hidden: bool, ?column: untyped, ?columns: untyped, ?col: untyped, ?row: untyped, ?rows: untyped, ?range: untyped, ?ref_sheet: String?, ?local: bool) -> void
    def defined_name(name, value = nil, sheet: nil, scope: nil, hidden: false, column: nil, columns: nil, col: nil, row: nil, rows: nil, range: nil, ref_sheet: nil, local: false)
      if col
        warn "[DEPRECATION] `col:` is deprecated; use `column:` (or `columns:`) instead.", uplevel: 1
      end
      if ref_sheet
        warn "[DEPRECATION] `ref_sheet:` is deprecated; use `sheet:` for reference sheet and `scope:` for scope instead.", uplevel: 1
      end
      if local
        warn "[DEPRECATION] `local:` is deprecated; use `scope:` instead.", uplevel: 1
      end
      target_ref_sheet = ref_sheet || sheet
      c_arg = column || columns || col
      r_arg = row || rows
      normalized_val = DslHelpers.normalize_defined_name_value(value, sheet: target_ref_sheet, col: c_arg, rows: r_arg, range: range)
      entry = { name: name, value: normalized_val, hidden: hidden }
      local_sheet = if scope
                      scope == :workbook ? nil : scope.to_s
                    elsif ref_sheet
                      sheet
                    elsif local
                      sheet
                    else
                      sheet
                    end
      if local_sheet
        # local_sheet_id will be resolved at close time
        entry[:local_sheet_name] = local_sheet
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
    # @param rows [String, Integer, Range, nil] Repeating row range (e.g. "1:2" or 0-based 0..1).
    # @param columns [String, Integer, Symbol, Range, nil] Repeating column range (e.g. "A:B" or 0..1).
    # @param cols [String, Integer, Symbol, Range, nil] Deprecated alias for columns.
    # @param sheet [String, nil] Target sheet name.
    # @return [void]
    # @api public
    #: (?rows: (String | Integer | Range[Integer])?, ?columns: (String | Integer | Symbol | Range[untyped])?, ?cols: (String | Integer | Symbol | Range[untyped])?, ?sheet: String?) -> void
    def print_titles(rows: nil, columns: nil, cols: nil, sheet: nil)
      if cols && columns.nil?
        warn "[DEPRECATION] `cols:` is deprecated; use `columns:` instead.", uplevel: 1
      end
      c = DslHelpers.normalize_print_titles_columns(columns || cols)
      r = DslHelpers.normalize_print_titles_rows(rows)
      sheet_name = sheet || @current_sheet || "Sheet1"
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
      defined_name("_xlnm.Print_Titles", value, sheet: sheet_name)
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
      print_titles(cols: cols_str, sheet: sheet)
    end

    # Sets workbook protection.
    #
    # @param opts [Hash] Protection options.
    # @return [void]
    # @api public
    #: (**String | Integer | bool | nil opts) -> void
    def workbook_protection(**opts)
      @workbook_protection = opts
    end

    # @deprecated Use {#workbook_protection} instead.
    #: (**String | Integer | bool | nil opts) -> void
    def protect_workbook(...)
      warn "[DEPRECATION] `protect_workbook` is deprecated; use `workbook_protection` instead.", uplevel: 1
      workbook_protection(...)
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

      @workbook_properties[:active_tab] = @active_sheet unless @active_sheet.nil?
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
    # @api public
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
          next rule if rule.respond_to?(:dxf_id) ? rule.dxf_id : (rule.is_a?(Hash) && (rule[:dxf_id] || rule[:format_id]))

          normalized = rule.dup
          dxf = normalized.delete(:dxf) || DslHelpers.extract_dxf_from_rule!(normalized)
          next normalized if dxf.nil? || dxf.empty?

          dxf_id = @dxfs.index(dxf)
          unless dxf_id
            @dxfs << dxf
            dxf_id = @dxfs.size - 1
          end
          normalized[:dxf_id] = dxf_id
          normalized[:format_id] = dxf_id
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
        col_breaks: @current_col_breaks,
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
        col_breaks: @current_col_breaks.empty? ? nil : @current_col_breaks,
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
