# frozen_string_literal: true

# rbs_inline: enabled

require_relative "dsl_helpers"

module Xlsxrb
  # DSL context for building a single in-memory worksheet in {Xlsxrb.build}.
  #
  # @api public
  class WorksheetBuilder
    # @param name [String] The worksheet name.
    # @param strict_excel_mode [Boolean] Whether to enforce Microsoft Excel limits.
    # @param state [Symbol] Sheet visibility state (:visible, :hidden, or :very_hidden).
    # @param style_log [StyleLog] When the workbook's styles are defined (internal).
    #: (String name, ?strict_excel_mode: bool, ?state: Symbol, ?style_log: StyleLog) -> void
    def initialize(name, strict_excel_mode: true, state: :visible, style_log: StyleLog.new)
      @name = name
      @style_log = style_log
      @column_style_clocks = {} # { column index => style log clock when its style was set }
      @strict_excel_mode = strict_excel_mode
      @state = state ? state.to_sym : :visible
      @next_row_index = 0
      @rows = []
      @columns = []
      @charts = []
      @styles = {} # { style_name => StyleBuilder }
      @style_index_map = {} # { style_name => xf_index } (populated at build time)
      @hyperlinks = []
      @auto_filter = nil
      @filter_columns = {}
      @sort_state = nil
      @data_validations = []
      @conditional_formats = []
      @tables = []
      @comments = []
      @threaded_comments = []
      @sparkline_groups = []
      @merge_cells_ranges = []
      @freeze_pane = nil
      @split_pane = nil
      @selection = nil
      @page_margins = nil
      @page_setup = {}
      @header_footer = {}
      @print_options = {}
      @sheet_protection = nil
      @images = []
      @shapes = []
      @sheet_properties = {}
      @sheet_view = {}
      @sheet_format = {}
      @row_breaks = []
      @col_breaks = []
      @protected_ranges = []
      @cell_watches = []
      @ignored_errors = []
      @phonetic_properties = nil
      @sheet_calc_pr = nil
      @background_image = nil
      @print_area = nil
      @print_titles = nil
      @defined_names = []
      @current_row_outline_level = nil
      @current_row_collapsed = false
      @current_row_hidden = false
      @pending_row_attributes = {}
    end

    #: String
    attr_reader :name

    #: Array[Hash[Symbol, untyped]]
    attr_reader :defined_names

    # Defines or configures a named cell style.
    #
    # @example Define a bold header style
    #   sheet.style(:header, bold: true, fill_color: "4F81BD", font_color: "FFFFFF")
    #
    # @param name [String, Symbol] The name of the style.
    # @param opts [Hash] Style options (e.g. bold: true, fill_color: "FF0000").
    # @yield [style_builder]
    # @yieldparam style_builder [Xlsxrb::StyleBuilder]
    # @return [StyleBuilder]
    #   Using the returned value is deprecated and warns; in version 0.3.0, this method
    #   will return nil. Configure the style with options or in the block.
    # @api public
    #: (String | Symbol name, **untyped opts) ?{ (StyleBuilder) -> void } -> StyleBuilder
    def style(name, **opts)
      style_name = name.to_s
      style_builder = StyleBuilder.new(style_name)
      style_builder.apply_options!(**opts) unless opts.empty?
      yield style_builder if block_given?
      @styles[style_name] = style_builder
      @style_log.record(style_name, style_builder)
      ReturnedStyleBuilder.new(style_builder)
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

    # Appends a row of cells to the worksheet.
    #
    # @example Add a row with values and array styles
    #   sheet.row(["ID", "Name", "Total"], cell_styles: [:header, :header, :header])
    #
    # @example Add a row with a hash of column keys
    #   sheet.row({ "A" => "Invoice", "C" => 12345 })
    #
    # @param values [Array<Object>, Hash{String, Integer => Object}] The cell values.
    # @param cell_styles [String, Symbol, Array<String | Symbol>, Hash, nil] Style names or inline styles to apply to each cell.
    # @param row_style [Integer, Symbol, String, Hash, nil] Row style format or index.
    # @param height [Float, Integer, nil] The row height in points (0 to 409).
    # @param hidden [Boolean] Whether the row is hidden.
    # @param custom_height [Boolean] Whether custom row height is enforced.
    # @param outline_level [Integer, nil] Grouping/outline hierarchy level.
    # @param collapsed [Boolean] Whether the row is collapsed.
    # @return [void]
    # @raise [ArgumentError] If limits are exceeded when strict_excel_mode is enabled.
    # @api public
    #: (Array[untyped] | Hash[untyped, untyped] values, ?cell_styles: untyped, ?row_style: (Integer | Symbol | String | Hash[Symbol, untyped])?, ?height: Float | Integer | nil, ?hidden: bool, ?custom_height: bool, ?collapsed: bool, ?outline_level: Integer | nil) -> void
    def row(values, cell_styles: nil, row_style: nil, height: nil, hidden: false, custom_height: false, collapsed: false, outline_level: nil)
      check_open!
      row_index = @next_row_index
      DslHelpers.validate_row_bounds!(row_index, height, strict_excel_mode: @strict_excel_mode)
      @next_row_index += 1

      pending = @pending_row_attributes.delete(row_index) || {}
      outline_level ||= @current_row_outline_level || pending[:outline_level]
      collapsed = @current_row_collapsed || pending[:collapsed] if !collapsed && (@current_row_collapsed || pending[:collapsed])
      hidden = @current_row_hidden || pending[:hidden] if !hidden && (@current_row_hidden || pending[:hidden])
      height ||= pending[:height]

      values = DslHelpers.normalize_row_values(values)
      cell_styles = DslHelpers.normalize_row_styles(cell_styles)
      # The row is kept as given until the workbook resolves its styles, so it must not
      # change with the caller's arrays. Each array is copied once, when first needed.
      values_copied = false
      styles_copied = false

      # Auto-detect Date / Time for built-in styles
      values.each_with_index do |val, idx|
        cell_style = cell_styles.is_a?(Array) ? cell_styles[idx] : cell_styles

        if val.is_a?(Date) && !val.is_a?(DateTime) && cell_style.nil?
          auto_style = inherited_date_style(row_style, idx)
          unless auto_style
            style("__xlsxrb_date", number_format: "yyyy-mm-dd") unless @styles.key?("__xlsxrb_date")
            auto_style = "__xlsxrb_date"
          end
        elsif val.is_a?(Time) || val.is_a?(DateTime)
          # A value that claims to be a Time without being one, such as
          # ActiveSupport::TimeWithZone, is kept as a Time with the same wall clock.
          unless Time === val || val.is_a?(DateTime) # rubocop:disable Style/CaseEquality
            unless values_copied
              values = values.dup
              values_copied = true
            end
            values[idx] = Ooxml::WorksheetWriter::TimeLike.to_time(val)
          end
          next unless cell_style.nil?

          auto_style = inherited_date_style(row_style, idx)
          unless auto_style
            style("__xlsxrb_time", number_format: "yyyy-mm-dd hh:mm:ss") unless @styles.key?("__xlsxrb_time")
            auto_style = "__xlsxrb_time"
          end
        else
          next
        end
        unless styles_copied
          cell_styles = if cell_styles.nil?
                          []
                        elsif cell_styles.is_a?(Array)
                          cell_styles.dup
                        else
                          Array.new(values.size, cell_styles)
                        end
          styles_copied = true
        end
        cell_styles[idx] = auto_style
      end

      max_len = values.size
      max_len = [max_len, cell_styles.size].max if cell_styles.is_a?(Array)
      # See: https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
      raise ArgumentError, "Row contains #{max_len} columns, exceeding Excel limit of 16_384 columns" if @strict_excel_mode && max_len > 16_384

      style_lookup = cell_styles.is_a?(Array)
      # Number format codes of the cells' sheet styles, as of this call.
      format_codes = style_lookup ? nil : format_code_of(cell_styles)

      col_index = 0
      while col_index < max_len
        val = col_index < values.size ? values[col_index] : nil
        raise ArgumentError, "Invalid cell value type or value: #{val.class} for value #{val.inspect}" unless val.nil? || val.is_a?(String) || (val.is_a?(Numeric) && !(val.is_a?(Float) && (val.infinite? || val.nan?))) || val.is_a?(TrueClass) || val.is_a?(FalseClass) || val.is_a?(Date) || val.is_a?(Time) || val.is_a?(Elements::Formula) || (val.is_a?(Hash) && val.key?(:formula)) || val.is_a?(Elements::RichText) || val.is_a?(Elements::CellError) || (val.is_a?(Array) && val.first.is_a?(Hash) && (val.first.key?(:text) || val.first.key?("text")))
        # See: https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
        raise ArgumentError, "Cell text length #{val.length} exceeds Excel limit of 32,767 characters" if @strict_excel_mode && val.is_a?(String) && val.length > 32_767

        if val.is_a?(Array) || val.is_a?(Hash)
          unless values_copied
            values = values.dup
            values_copied = true
          end
          values[col_index] = if val.is_a?(Array)
                                # Coerce array of hashes to RichText
                                runs = val.map do |run|
                                  text = run[:text] || run["text"]
                                  font = run.reject { |k| k.to_s == "text" }
                                  { text: text, font: font.empty? ? nil : font }.compact
                                end
                                Elements::RichText.new(runs: runs)
                              else
                                # A Hash with :formula becomes the cell's formula.
                                Elements::Formula.new(
                                  expression: val[:formula] || val[:expression] || "",
                                  cached_value: val.key?(:value) ? val[:value] : val[:cached_value],
                                  type: val[:type] || (:shared if val.key?(:shared_index)),
                                  ref: val[:ref],
                                  shared_index: val[:shared_index],
                                  calculate_always: val[:calculate_always]
                                )
                              end
        end

        if style_lookup
          style_name = col_index < cell_styles.size ? cell_styles[col_index] : nil
          if style_name.is_a?(Hash)
            inline_name = "__inline_#{style_name.hash}"
            style(inline_name, **style_name) unless @styles.key?(inline_name)
            unless styles_copied
              cell_styles = cell_styles.dup
              styles_copied = true
            end
            cell_styles[col_index] = style_name = inline_name
          end
          format_code = style_name.nil? ? nil : format_code_of(style_name)
          (format_codes ||= [])[col_index] = format_code if format_code
        end
        col_index += 1
      end

      values = values.dup unless values_copied || values.frozen?
      cell_styles = cell_styles.dup if style_lookup && !styles_copied && !cell_styles.frozen?
      @rows << PendingRow.new(row_index, values, cell_styles, format_codes, max_len, height, hidden,
                              custom_height || !height.nil?, collapsed, outline_level, row_style, @style_log.clock)
      nil
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
    #   sheet.skip_rows(3)
    #
    # @param count [Integer] Number of rows to skip (must be >= 1).
    # @return [void]
    # @api public
    #: (?Integer count) -> void
    def skip_rows(count = 1)
      check_open!
      raise ArgumentError, "count must be positive: #{count.inspect}" if !count.is_a?(Integer) || count < 1

      @next_row_index += count
      nil
    end

    # Sets column formatting and properties for one or multiple columns.
    #
    # @example Set width for column A
    #   sheet.column("A", width: 20.0)
    #
    # @example Set width for a range of columns
    #   sheet.column("A".."D", width: 15.0)
    #
    # @param index [Integer, String, Range, Array] Column index (0-based), letter ("A"), or range ("A".."D").
    # @param width [Float, Integer, nil] Column width in character units (0 to 255).
    # @param collapsed [Boolean] Whether the column is collapsed.
    # @param style [Integer, Symbol, String, nil] Column style format or index.
    # @return [void]
    # @raise [ArgumentError] If width exceeds 255 in strict mode.
    # @api public
    #: (Integer | String | Symbol | Range[Integer | String | Symbol] | Array[Integer | String | Symbol] index, ?width: Float | Integer | nil, ?hidden: bool, ?custom_width: bool, ?outline_level: Integer | nil, ?collapsed: bool, ?style: (Integer | Symbol | String | Hash[Symbol, untyped])?, ?best_fit: bool) -> void
    def column(index, width: nil, hidden: false, custom_width: false, outline_level: nil, collapsed: false, style: nil, best_fit: false)
      raise ArgumentError, "Column width #{width} must be between 0 and 255 characters (Excel limitation)" if @strict_excel_mode && width && (width.negative? || width > 255)

      DslHelpers.normalize_column_indices(index).each do |idx|
        @column_style_clocks[idx] = @style_log.clock unless style.nil?
        existing_idx = @columns.find_index { |c| c.index == idx }
        if existing_idx
          prev = @columns[existing_idx]
          new_width = width.nil? ? prev.width : width
          new_custom_width = (custom_width || !width.nil?) || prev.custom_width?
          new_hidden = hidden || prev.hidden?
          new_outline = outline_level.nil? ? prev.outline_level : outline_level
          new_style = style.nil? ? prev.style_index : style
          unmapped = prev.unmapped_data ? prev.unmapped_data.dup : {}
          unmapped[:style_index] = new_style if new_style
          unmapped[:collapsed] = true if collapsed || prev.unmapped_data&.[](:collapsed)
          unmapped[:best_fit] = true if best_fit || prev.unmapped_data&.[](:best_fit)
          @columns[existing_idx] = Elements::Column.new(
            index: idx,
            width: new_width,
            hidden: new_hidden,
            custom_width: new_custom_width,
            outline_level: new_outline,
            style_index: new_style,
            unmapped_data: unmapped
          )
        else
          unmapped = style ? { style_index: style } : {}
          unmapped[:collapsed] = true if collapsed
          unmapped[:best_fit] = true if best_fit
          @columns << Elements::Column.new(
            index: idx,
            width: width,
            hidden: hidden,
            custom_width: custom_width || !width.nil?,
            outline_level: outline_level,
            style_index: style,
            unmapped_data: unmapped
          )
        end
      end
      @columns.sort_by!(&:index)
      nil
    end

    # Returns column configurations defined for this worksheet.
    #
    # @return [Array<Elements::Column>]
    # @api public
    #: () -> Array[Elements::Column]
    def columns
      @columns.map { |c| c.is_a?(Elements::Column) ? c : Elements::Column.new(**c) }
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

    # Groups multiple rows with outline hierarchy, either by range or with a block.
    #
    # @example Group rows with a block
    #   sheet.group_rows(level: 1, collapsed: true) do
    #     sheet.row(["Subtotal Item 1"])
    #     sheet.row(["Subtotal Item 2"])
    #   end
    #
    # @example Group rows by 0-based range
    #   sheet.group_rows(1..3, level: 1, collapsed: true)
    #
    # @param rows [Range, Array, Integer, nil] 0-based row index or range.
    # @param level [Integer] Grouping outline level (1 to 7).
    # @param collapsed [Boolean] Whether the group is collapsed.
    # @param hidden [Boolean] Whether the rows are hidden.
    # @yield Optional block executed with active group outline settings.
    # @return [void]
    # @api public
    #: (?untyped rows, ?level: Integer, ?collapsed: bool, ?hidden: bool) ?{ (self) -> void } -> void
    def group_rows(rows = nil, level: 1, collapsed: false, hidden: false)
      if block_given?
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
      elsif rows
        row_indices = case rows
                      when Range then rows.to_a
                      when Array then rows
                      else [rows.to_i]
                      end
        check_open!
        row_indices.each do |idx|
          existing_row = @rows.find { |r| r.index == idx }
          if existing_row
            existing_idx = @rows.index(existing_row)
            @rows[existing_idx] = existing_row.with_outline(
              hidden: hidden || existing_row.hidden,
              collapsed: collapsed || existing_row.collapsed,
              outline_level: level
            )
          else
            @pending_row_attributes[idx] = { outline_level: level, collapsed: collapsed, hidden: hidden }
          end
        end
      end
      nil
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
      return @sheet_properties.slice(:summary_below, :summary_right, :show_outline_symbols, :apply_styles) if summary_below.nil? && summary_right.nil? && show_outline_symbols.nil? && apply_styles.nil? && opts.empty?

      props = (@sheet_properties[:outline_properties] || {}).dup
      props[:summary_below] = summary_below unless summary_below.nil?
      props[:summary_right] = summary_right unless summary_right.nil?
      props[:show_outline_symbols] = show_outline_symbols unless show_outline_symbols.nil?
      props[:apply_styles] = apply_styles unless apply_styles.nil?
      props.merge!(opts)
      @sheet_properties[:outline_properties] = props
      @sheet_properties[:summary_below] = props[:summary_below] unless props[:summary_below].nil?
      @sheet_properties[:summary_right] = props[:summary_right] unless props[:summary_right].nil?
      nil
    end

    # Sets worksheet outline and grouping view properties.
    #
    # @param opts [Hash] Outline properties.
    # @return [void]
    # @api public
    #: (Hash[Symbol, untyped] opts) -> void
    def outline_properties=(opts)
      h = opts.respond_to?(:to_h) ? opts.to_h : (opts || {})
      outline_properties(**h)
      nil
    end

    # Adds a chart to the worksheet.
    #
    # @param type [Symbol, String, nil] Optional chart type (e.g. :bar, :line).
    # @param options [Hash] Chart configuration options.
    # @yield [builder]
    # @yieldparam builder [Xlsxrb::ChartBuilder]
    # @return [void]
    # @api public
    #: (?Symbol | String? type, **untyped options) ?{ (ChartBuilder) -> void } -> void
    def chart(type = nil, **options)
      options = options.merge(type: type) if type
      if block_given?
        builder = ChartBuilder.new
        builder.type = type if type
        yield builder
        options = builder.options.merge(options)
      end
      @charts << options
      nil
    end

    # Adds a hyperlink to a cell.
    #
    # @param cell [String, Integer] Cell coordinate (e.g. "A1").
    # @param url [String, nil] Target external URL.
    # @param display [String, nil] Optional display text.
    # @param tooltip [String, nil] Optional hover tooltip text.
    # @param location [String, nil] Optional internal sheet location (e.g. "Sheet2!A1").
    # @return [void]
    # @api public
    #: (String | Integer cell, ?(String | Elements::Hyperlink)? url, ?display: String?, ?tooltip: String?, ?location: String?) -> void
    def hyperlink(cell, url = nil, display: nil, tooltip: nil, location: nil)
      if url.is_a?(Elements::Hyperlink)
        link = { cell: cell }
        link[:url] = url.url if url.url
        link[:display] = url.display || display if url.display || display
        link[:tooltip] = url.tooltip || tooltip if url.tooltip || tooltip
        link[:location] = url.location || location if url.location || location
        @hyperlinks << link
        return
      end
      link = { cell: cell }
      link[:url] = url if url
      link[:display] = display if display
      link[:tooltip] = tooltip if tooltip
      link[:location] = location if location
      @hyperlinks << link
    end

    # Returns the auto-filter range for the worksheet, or nil if none.
    #
    # @return [String, nil]
    # @api public
    #: () -> String?
    attr_reader :auto_filter

    # Sets the auto-filter range for the worksheet.
    #
    # @param range [String, nil] Cell range (e.g. "A1:E100").
    # @return [String, nil]
    # @api public
    #: (String? range) -> String?
    def auto_filter=(range)
      @auto_filter = range&.to_s
    end

    # Configures filtering rules on a specific column of an auto-filter.
    #
    # @param col_id [Integer, String, Symbol] 0-based column index relative to filter range or column letter ("A", :B).
    # @param filter [Hash] Filter criteria (e.g. values, custom filters).
    # @return [void]
    # @api public
    #: (Integer | String | Symbol col_id, (Hash[untyped, untyped] | Array[untyped]) filter) -> void
    def filter_column(col_id, filter)
      col_id = Elements::Cell._column_index(col_id) if col_id.is_a?(String) || col_id.is_a?(Symbol)
      @filter_columns[col_id] = filter
      nil
    end

    # Configures or queries column sort state for the sheet.
    #
    # When called with no arguments, returns the {Elements::SortState} value object.
    # When called with arguments, updates sort state and returns nil.
    #
    # @param ref [String, nil] The sorted range.
    # @param sort_conditions [Array<Hash>, nil] List of sort condition definitions.
    # @param case_sensitive [Boolean] Whether sorting is case sensitive.
    # @param sort_method [String, nil] Sort method.
    # @param opts [Hash] Additional sort state options.
    # @return [Elements::SortState, nil, void]
    # @api public
    #: () -> Elements::SortState?
    #: (?String? ref, ?Array[Hash[untyped, untyped]]? sort_conditions, ?case_sensitive: bool, ?sort_method: String?, **untyped opts) -> void
    def sort_state(ref = nil, sort_conditions = nil, case_sensitive: false, sort_method: nil, **opts)
      if ref.nil? && sort_conditions.nil?
        return @sort_state ? Elements::SortState.new(**@sort_state) : nil
      end

      entry = { ref: ref, sort_conditions: sort_conditions, case_sensitive: case_sensitive }
      entry[:sort_method] = sort_method if sort_method
      @sort_state = entry.merge(opts)
      nil
    end

    # Sets column sort state for the sheet.
    #
    # @param opts [Elements::SortState, Hash, nil] Sort state options.
    # @return [void]
    # @api public
    #: (Elements::SortState | Hash[Symbol, untyped] | nil opts) -> void
    def sort_state=(opts)
      @sort_state = opts&.respond_to?(:to_h) ? opts.to_h : opts
      nil
    end

    # Adds a data validation rule to a cell, range, or coordinates.
    #
    # @param sqref [String, nil] Target cell or range reference (e.g. "B2:B100").
    # @param opts [Hash] Validation configuration (e.g. type: :list, formula1: '"Option1,Option2"', or coordinate options `columns:`, `rows:`, `except_rows:`, `list:`, `in:`).
    # @yield [data_validation_builder] Optional configuration block.
    # @return [void]
    # @api public
    #: (?String? sqref, **untyped opts) ?{ (DataValidationBuilder) -> void } -> void
    def data_validation(sqref = nil, **opts, &)
      normalized_list = DslHelpers.normalize_data_validation(sqref, **opts, &)
      @data_validations.concat(normalized_list)
    end

    # Adds a conditional formatting rule to a cell range.
    # Supports block configuration via {ConditionalFormatBuilder} and coordinate targeting (:columns, :rows).
    #
    # @param sqref [String, nil] Target cell range reference (e.g. "C2:C50").
    # @param opts [Hash] Rule options (e.g. type: :cellIs, operator: :greaterThan, formula: "100").
    # @option opts [Hash] :dxf Hierarchical differential format specification (e.g. `dxf: { font: { bold: false }, fill: { pattern: "none" } }`).
    # @option opts [Integer] :format_id Pre-registered differential format index (dxfId).
    # @yield [conditional_format_builder] Optional block for configuration.
    # @return [void]
    # @api public
    #: (?String? sqref, **untyped opts) ?{ (ConditionalFormatBuilder) -> void } -> void
    def conditional_format(sqref = nil, **opts, &)
      normalized_list = DslHelpers.normalize_conditional_format(sqref, **opts, &)
      @conditional_formats.concat(normalized_list)
    end

    # Adds an Excel Table (ListObject) to the sheet.
    #
    # @param ref [String, nil] The table range (e.g. "A1:D20").
    # @param columns [Array<String>, nil] Column header names.
    # @param name [String, nil] Table name identifier.
    # @param display_name [String, nil] Table display name.
    # @param style [String, nil] Built-in table style name (e.g. "TableStyleMedium2").
    # @param opts [Hash] Additional table properties.
    # @yield [table_builder] Optional configuration block.
    # @return [void]
    # @api public
    #: (?String? ref, ?columns: Array[String | Hash[Symbol, untyped]]?, ?name: String?, ?display_name: String?, ?style: String?, **untyped opts) ?{ (TableBuilder) -> void } -> void
    def table(ref = nil, columns: nil, name: nil, display_name: nil, style: nil, **opts)
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

      @tables << DslHelpers.normalize_table_options(ref, columns: columns, name: name, display_name: display_name, style: style, **opts)
    end

    # Adds a Pivot Table to the sheet.
    #
    # @param source_ref [String] Source data range (e.g. "DataSheet!A1:D100").
    # @param row_fields [Array<Integer, String>] Field names or indices for row headers.
    # @param data_fields [Array<Hash>] Data aggregation fields.
    # @param column_fields [Array<String, Integer>, nil] Field names or indices for columns.
    # @param ref [String, nil] Target cell for pivot table top-left corner (e.g. "E1").
    # @param name [String, nil] Pivot table name.
    # @param field_names [Array<String>, nil] Custom field display names.
    # @param items [Hash{Integer => Array}, nil] Items of each field, by 0-based field index.
    # @return [void]
    # @api public
    #: (String source_ref, row_fields: Array[String | Integer], data_fields: Array[String | Hash[Symbol, untyped]], ?column_fields: Array[String | Integer]?, ?ref: String?, ?name: String?, ?field_names: Array[String]?, ?items: Hash[Integer, Array[untyped]]?, **untyped opts) -> void
    def pivot_table(source_ref, row_fields:, data_fields:, column_fields: nil, ref: nil, name: nil, field_names: nil, items: nil, **opts)
      c_fields = column_fields || []
      target_ref = ref || "E1"
      @pivot_tables ||= []
      entry = {
        source_ref: source_ref, row_fields: row_fields,
        data_fields: data_fields, column_fields: c_fields,
        ref: target_ref, destination: target_ref, name: name,
        field_names: field_names, items: items
      }.merge(opts)
      @pivot_tables << entry
    end

    # Adds a cell comment / note.
    #
    # @param cell [String, Integer] Cell reference (e.g. "B2").
    # @param text [String] The comment text.
    # @param author [String] Author name.
    # @param visible [Boolean] Whether the comment is always shown instead of on hover.
    # @return [void]
    # @api public
    #: (String | Integer cell, String text, ?author: ::String, ?visible: bool) -> void
    def comment(cell, text, author: "Author", visible: false)
      ref_str = cell.to_s
      @comments << { ref: ref_str, text: text, author: author, visible: visible }
      nil
    end

    # Adds a threaded comment to a cell.
    #
    # @param cell [String, Integer] Cell reference (e.g. "B2").
    # @param text [String] The comment text.
    # @param author [String] Author display name.
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
      ref_str = cell.to_s
      entry = {
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
      @threaded_comments << entry
      nil
    end

    # Adds a sparkline group to the sheet.
    #
    # @param sparklines [Array<Hash>] List of { data_ref:, location_ref: } items.
    # @param type [String, Symbol, nil] "line" (default), "column", or "stacked".
    # @param opts [Hash] Color and formatting options.
    # @return [void]
    # @api public
    #: (sparklines: Array[String | Hash[Symbol, untyped]], ?type: (String | Symbol)?, **untyped opts) -> void
    def sparkline_group(sparklines:, type: nil, **opts)
      group = { sparklines: sparklines }
      group[:type] = type if type
      group.merge!(opts)
      @sparkline_groups << group
    end

    # Merges a range of cells into a single cell.
    #
    # @example Merge with string range
    #   sheet.merge("A1:C1")
    #
    # @param range [String] Cell range (e.g. "A1:B2").
    # @return [void]
    # @api public
    #: (String range) -> void
    def merge(range)
      canonical = DslHelpers.normalize_merge_range(range, strict_excel_mode: @strict_excel_mode)
      @merge_cells_ranges << canonical unless @merge_cells_ranges.include?(canonical)
    end

    # Configures or queries freeze panes (<pane state="frozen">).
    #
    # When called with no arguments, returns the {Elements::FreezePane} value object or nil.
    # When called with arguments, configures the freeze pane and returns nil.
    #
    # @param row [Integer, nil] Number of rows to freeze from top (0-based split position).
    # @param column [Integer, String, nil] Number of columns to freeze from left (0-based or letter).
    # @param top_left_cell [String, nil] Top-left cell reference in bottom-right pane.
    # @param active_pane [String, Symbol, nil] Active pane identifier.
    # @return [Elements::FreezePane, nil, void]
    # @api public
    #: () -> Elements::FreezePane?
    #: (?row: Integer | nil, ?column: (Integer | String)?, ?top_left_cell: String?, ?active_pane: (String | Symbol)?) -> void
    def freeze_pane(row: nil, column: nil, top_left_cell: nil, active_pane: nil)
      if row.nil? && column.nil? && top_left_cell.nil? && active_pane.nil?
        return @freeze_pane ? Elements::FreezePane.new(**@freeze_pane) : nil
      end

      c = column ? Elements::Cell._column_index(column) : 0
      entry = { row: row || 0, column: c }
      entry[:top_left_cell] = top_left_cell if top_left_cell
      entry[:active_pane] = active_pane.to_s if active_pane
      @freeze_pane = entry
      nil
    end

    # Sets or clears freeze pane configuration.
    #
    # @param opts [Elements::FreezePane, Hash, nil] Freeze pane options.
    # @return [void]
    # @api public
    #: (Elements::FreezePane | Hash[Symbol, untyped] | nil opts) -> void
    def freeze_pane=(opts)
      if opts.nil?
        @freeze_pane = nil
      else
        h = opts.respond_to?(:to_h) ? opts.to_h : opts
        c = h[:column] || h[:col]
        c = Elements::Cell._column_index(c) if c
        entry = { row: h[:row] || 0, column: c || 0 }
        entry[:top_left_cell] = h[:top_left_cell] if h[:top_left_cell]
        entry[:active_pane] = h[:active_pane].to_s if h[:active_pane]
        @freeze_pane = entry
      end
      nil
    end

    # Splits window panes without freezing or queries split pane settings (<pane>).
    #
    # When called with no arguments, returns the {Elements::SplitPane} value object or nil.
    # When called with arguments, configures the split pane and returns nil.
    #
    # @param x_split [Integer, nil] Horizontal split offset in points.
    # @param y_split [Integer, nil] Vertical split offset in points.
    # @param top_left_cell [String, nil] Top-left cell reference in bottom-right pane.
    # @param active_pane [String, Symbol, nil] Active pane identifier.
    # @return [Elements::SplitPane, nil, void]
    # @api public
    #: () -> Elements::SplitPane?
    #: (?x_split: Integer | nil, ?y_split: Integer | nil, ?top_left_cell: String?, ?active_pane: (String | Symbol)?) -> void
    def split_pane(x_split: nil, y_split: nil, top_left_cell: nil, active_pane: nil)
      if x_split.nil? && y_split.nil? && top_left_cell.nil? && active_pane.nil?
        return @split_pane ? Elements::SplitPane.new(**@split_pane) : nil
      end

      entry = { x_split: x_split || 0, y_split: y_split || 0, top_left_cell: top_left_cell }
      entry[:active_pane] = active_pane.to_s if active_pane
      @split_pane = entry
      nil
    end

    # Sets or clears split pane configuration.
    #
    # @param opts [Elements::SplitPane, Hash, nil] Split pane options.
    # @return [void]
    # @api public
    #: (Elements::SplitPane | Hash[Symbol, untyped] | nil opts) -> void
    def split_pane=(opts)
      if opts.nil?
        @split_pane = nil
      else
        h = opts.respond_to?(:to_h) ? opts.to_h : opts
        entry = { x_split: h[:x_split] || 0, y_split: h[:y_split] || 0, top_left_cell: h[:top_left_cell] }
        entry[:active_pane] = h[:active_pane].to_s if h[:active_pane]
        @split_pane = entry
      end
      nil
    end

    # Configures or queries the active cell and selection area (<selection>).
    #
    # When called with no arguments, returns the selection ({Elements::Selection}, Array, or nil).
    # When called with arguments, updates the selection and returns nil.
    #
    # @param active_cell [String, nil] Active cell reference (e.g. "A1") or range (e.g. "A1:D10").
    # @param sqref [String, nil] Selected range reference (e.g. "A1:D10").
    # @param pane [String, Symbol, nil] Target pane (:topLeft, :topRight, :bottomLeft, :bottomRight).
    # @return [Elements::Selection, Array<Elements::Selection>, nil, void]
    # @api public
    #: () -> (Elements::Selection | Array[Elements::Selection])?
    #: (?String? active_cell, ?sqref: String?, ?pane: (String | Symbol)?) -> void
    def selection(active_cell = nil, sqref: nil, pane: nil)
      if active_cell.nil? && sqref.nil? && pane.nil?
        return nil if @selection.nil?
        return @selection.map { |s| Elements::Selection.new(**s) } if @selection.is_a?(Array)

        return Elements::Selection.new(**@selection)
      end

      if sqref.nil? && active_cell.is_a?(String) && active_cell.include?(":")
        sqref = active_cell
        active_cell = sqref.split(":").first
      else
        sqref ||= active_cell
      end
      sel = { active_cell: active_cell, sqref: sqref }
      sel[:pane] = pane if pane
      @selection = sel
      nil
    end

    # Sets or clears the selection configuration.
    #
    # @param opts [Elements::Selection, Array<Elements::Selection, Hash>, Hash, nil]
    # @return [void]
    # @api public
    #: (Elements::Selection | Array[Elements::Selection | Hash[Symbol, untyped]] | Hash[Symbol, untyped] | nil opts) -> void
    def selection=(opts)
      @selection = if opts.nil?
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
      if left.nil? && right.nil? && top.nil? && bottom.nil? && header.nil? && footer.nil?
        return @page_margins ? Elements::PageMargins.new(**@page_margins) : nil
      end

      @page_margins = DslHelpers.normalize_page_margins(
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
      if opts.nil?
        @page_margins = nil
      else
        h = opts.respond_to?(:to_h) ? opts.to_h : opts
        @page_margins = DslHelpers.normalize_page_margins(**h)
      end
      nil
    end

    # Configures or queries page setup configuration (<pageSetup>).
    #
    # When called with no arguments, returns the {Elements::PageSetup} value object.
    # When called with keyword arguments, updates page setup and returns nil.
    #
    # @param opts [Hash] Page setup options (e.g. orientation: :landscape, paper_size: 9).
    # @return [Elements::PageSetup, void]
    # @api public
    #: () -> Elements::PageSetup
    #: (**untyped opts) -> void
    def page_setup(**opts)
      return Elements::PageSetup.new(**@page_setup) if opts.empty?

      @page_setup.merge!(opts)
      @sheet_properties[:fit_to_page] = true if (opts.key?(:fit_to_width) || opts.key?(:fit_to_height)) && !@sheet_properties.key?(:fit_to_page)
      nil
    end

    # Sets page setup configuration.
    #
    # @param opts [Hash, Elements::PageSetup]
    # @return [void]
    # @api public
    #: (Hash[Symbol, untyped] | Elements::PageSetup opts) -> void
    def page_setup=(opts)
      h = opts.respond_to?(:to_h) ? opts.to_h : (opts || {})
      @page_setup = h.transform_keys(&:to_sym)
      nil
    end

    # Returns whether page setup fits content to page width/height.
    #
    # @return [Boolean]
    # @api public
    #: () -> bool
    def fit_to_page?
      @sheet_properties[:fit_to_page] ? true : false
    end

    # Sets whether page setup fits content to page width/height.
    #
    # @param enabled [Boolean, nil] Whether to fit to page.
    # @return [Boolean]
    # @api public
    #: (bool? enabled) -> bool?
    def fit_to_page=(enabled)
      @sheet_properties[:fit_to_page] = enabled ? true : false
    end

    # Configures or queries headers and footers for printing (<headerFooter>).
    #
    # When called with no arguments, returns the {Elements::HeaderFooter} value object.
    # When called with keyword arguments, updates headers and footers and returns nil.
    #
    # @param opts [Hash] Header and footer options (e.g. odd_header: "&CHeader Text").
    # @return [Elements::HeaderFooter, void]
    # @api public
    #: () -> Elements::HeaderFooter
    #: (**untyped opts) -> void
    def header_footer(**opts)
      return Elements::HeaderFooter.new(**@header_footer) if opts.empty?

      @header_footer.merge!(opts)
      nil
    end

    # Sets header and footer configuration.
    #
    # @param opts [Hash, Elements::HeaderFooter]
    # @return [void]
    # @api public
    #: (Hash[Symbol, untyped] | Elements::HeaderFooter opts) -> void
    def header_footer=(opts)
      h = opts.respond_to?(:to_h) ? opts.to_h : (opts || {})
      @header_footer = h.transform_keys(&:to_sym)
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
      return Elements::PrintOptions.new(**@print_options) if opts.empty?

      opts.each { |k, v| @print_options[k.to_sym] = v }
      nil
    end

    # Replaces print options with the given hash or {Elements::PrintOptions} object.
    #
    # @param opts [Hash, Elements::PrintOptions]
    # @return [void]
    # @api public
    #: (Hash[Symbol, untyped] | Elements::PrintOptions opts) -> void
    def print_options=(opts)
      h = opts.respond_to?(:to_h) ? opts.to_h : (opts || {})
      @print_options = h.transform_keys(&:to_sym)
      nil
    end

    # Adds a defined name (named range or formula) scoped to this sheet or workbook.
    # Bare cell references and coordinate keywords are automatically formatted with sheet quotes and absolute $ references.
    #
    # @param name [String] Defined name identifier.
    # @param value [Object, nil] Cell reference, formula, or constant.
    # @param hidden [Boolean] Whether the defined name is hidden from Excel's UI.
    # @param columns [Integer, String, nil] Column index or letter for coordinate-based reference.
    # @param rows [Integer, Range, nil] Row or row range (0-based) for coordinate-based reference.
    # @param range [String, nil] Cell range reference (e.g. "A1:A10").
    # @param sheet [String, nil] Reference sheet name (defaults to current sheet).
    # @param scope [String, Symbol, nil] Scope identifier (:workbook or sheet name).
    # @return [void]
    # @api public
    #: (String name, ?untyped value, ?hidden: bool, ?columns: untyped, ?rows: untyped, ?range: untyped, ?sheet: String?, ?scope: (String | Symbol)?) -> void
    def defined_name(name, value = nil, hidden: false, columns: nil, rows: nil, range: nil, sheet: nil, scope: nil)
      ref_sheet = sheet || @name
      normalized_val = DslHelpers.normalize_defined_name_value(value, sheet: ref_sheet, columns: columns, rows: rows, range: range)
      entry = { name: name, value: normalized_val, hidden: hidden }
      local_sheet = if scope
                      scope == :workbook ? nil : scope.to_s
                    end
      entry[:local_sheet_name] = local_sheet if local_sheet
      @defined_names << entry
    end

    # Returns the print area for the worksheet, or nil if none.
    #
    # @example Get print area range
    #   sheet.print_area #=> "A1:H50"
    #
    # @return [String, nil]
    # @api public
    #: () -> String?
    attr_reader :print_area

    # Sets the print area for the worksheet.
    #
    # @param range [String, nil] Cell range (e.g. "A1:H50").
    # @return [String, nil]
    # @api public
    #: (String? range) -> String?
    def print_area=(range)
      @print_area = range&.to_s
    end

    # Sets or gets repeating print titles (rows and/or columns) for pagination.
    #
    # When called with no arguments, returns the {Elements::PrintTitles} value object or nil.
    # When called with arguments, updates print titles and returns nil.
    #
    # @example Set repeating header rows
    #   sheet.print_titles(rows: "1:2")
    #
    # @example Get repeating print titles
    #   sheet.print_titles #=> #<Xlsxrb::Elements::PrintTitles ...>
    #
    # @param rows [String, Integer, Range, nil] Repeating row range (e.g. "1:2" or 0-based 0..1).
    # @param columns [String, Integer, Symbol, Range, nil] Repeating column range (e.g. "A:B" or 0..1).
    # @return [Elements::PrintTitles, nil, void]
    # @api public
    #: () -> Elements::PrintTitles?
    #: (?rows: (String | Integer | Range[Integer])?, ?columns: (String | Integer | Symbol | Range[untyped])?) -> void
    def print_titles(rows: nil, columns: nil)
      return @print_titles ? Elements::PrintTitles.new(**@print_titles) : nil if rows.nil? && columns.nil?

      norm_rows = DslHelpers.normalize_print_titles_rows(rows)
      norm_cols = DslHelpers.normalize_print_titles_columns(columns)
      @print_titles = Elements::PrintTitles.new(rows: norm_rows, columns: norm_cols)
      nil
    end

    # Sets repeating print titles for the worksheet.
    #
    # @param opts [Elements::PrintTitles, Hash, nil] Options with :rows and/or :columns.
    # @return [Elements::PrintTitles, Hash, nil]
    # @api public
    #: (Elements::PrintTitles | Hash[Symbol, untyped] | nil opts) -> (Elements::PrintTitles | Hash[Symbol, untyped] | nil)
    def print_titles=(opts)
      if opts.nil?
        @print_titles = nil
      else
        h = opts.respond_to?(:to_h) ? opts.to_h : opts
        print_titles(**h)
      end
      opts
    end

    # Sets repeating title rows to appear at the top of each printed page.
    # Accepts 0-indexed row numbers (WriteXLSX / xlsxrb convention) or Excel range string ("1:2").
    #
    # @example Repeat first row
    #   sheet.repeat_rows(0)
    #
    # @example Repeat rows 1 through 3
    #   sheet.repeat_rows(0, 2)
    #
    # @param first [Integer, String] First row index (0-based) or range string (e.g. "1:2").
    # @param last [Integer, String, nil] Optional last row index (0-based, defaults to first).
    # @return [void]
    # @api public
    #: (Integer | String first, ?(Integer | String)? last) -> void
    def repeat_rows(first, last = nil)
      rows_str = if first.is_a?(String)
                   first.include?(":") ? first : "#{first}:#{last || first}"
                 else
                   last ||= first
                   "#{first + 1}:#{last + 1}"
                 end
      existing_cols = @print_titles&.columns
      print_titles(rows: rows_str, columns: existing_cols)
    end

    # Sets repeating title columns to appear at the left of each printed page.
    # Accepts 0-indexed column numbers (0 => "A"), column names ("A", "B"), or range string ("A:B").
    #
    # @example Repeat column A
    #   sheet.repeat_columns(0)
    #
    # @example Repeat columns A through C
    #   sheet.repeat_columns("A", "C")
    #
    # @param first [Integer, String] First column index (0-based) or column name/range.
    # @param last [Integer, String, nil] Optional last column index/name (defaults to first).
    # @return [void]
    # @api public
    #: (Integer | String first, ?(Integer | String)? last) -> void
    def repeat_columns(first, last = nil)
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
      existing_rows = @print_titles&.rows
      print_titles(rows: existing_rows, columns: cols_str)
    end

    # Sets or queries sheet-level protection with optional password hashing (<sheetProtection>).
    #
    # When called with no arguments, returns the {Elements::SheetProtection} value object or nil.
    # When called with keyword arguments, sets sheet protection and returns nil.
    #
    # @param opts [Hash] Protection options (e.g. password: "secret", select_locked_cells: true).
    # @return [Elements::SheetProtection, nil, void]
    # @api public
    #: () -> Elements::SheetProtection?
    #: (**untyped opts) -> void
    def sheet_protection(**opts)
      if opts.empty?
        @sheet_protection ? Elements::SheetProtection.new(**@sheet_protection) : nil
      else
        @sheet_protection = DslHelpers.normalize_protection_options(opts)
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
      if opts.nil? || opts == false
        @sheet_protection = nil
      elsif opts == true
        @sheet_protection = DslHelpers.normalize_protection_options(sheet: true)
      else
        h = opts.respond_to?(:to_h) ? opts.to_h : opts
        @sheet_protection = DslHelpers.normalize_protection_options(h)
      end
      nil
    end

    # Inserts an embedded image from raw binary data.
    #
    # @param file_data [String] Binary image data.
    # @param ext [String] File extension ("png", "jpeg", etc.).
    # @param from_column [Integer, nil] Top-left starting column (0-based).
    # @param from_row [Integer] Top-left starting row (0-based).
    # @param to_column [Integer, nil] Bottom-right ending column (0-based).
    # @param to_row [Integer] Bottom-right ending row (0-based).
    # @param opts [Hash] Additional image anchoring options.
    # @return [void]
    # @api public
    #: (String file_data, ?ext: ::String, ?from_column: ::Integer?, ?from_row: ::Integer, ?to_column: ::Integer?, ?to_row: ::Integer, **untyped opts) -> void
    def image(file_data, ext: "png", from_column: nil, from_row: 0, to_column: nil, to_row: 10, **opts)
      fc = from_column || 0
      tc = to_column || 5
      img = { file_data: file_data, ext: ext, from_column: fc, from_row: from_row, to_column: tc, to_row: to_row }
      img.merge!(opts)
      @images << img
    end

    # Gets the worksheet background image (<picture>).
    #
    # @return [Elements::BackgroundImage, nil]
    # @api public
    #: () -> Elements::BackgroundImage?
    def background_image
      return nil unless @background_image

      @background_image.is_a?(Elements::BackgroundImage) ? @background_image : Elements::BackgroundImage.new(**@background_image)
    end

    # Sets the worksheet background image (<picture>).
    #
    # @param val [String, Hash, Elements::Image, nil] Image binary data, file path, or image options hash.
    # @return [untyped]
    # @api public
    #: (untyped val) -> untyped
    def background_image=(val)
      if val.nil?
        @background_image = nil
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
      @background_image = {
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

    # Inserts a drawing shape (e.g. rectangle, callout, arrow).
    #
    # @param preset [String] Shape preset name (e.g. "rect", "roundRect").
    # @param text [String, nil] Shape label text.
    # @param from_column [Integer, nil] Top-left column.
    # @param from_row [Integer] Top-left row.
    # @param to_column [Integer, nil] Bottom-right column.
    # @param to_row [Integer] Bottom-right row.
    # @param opts [Hash] Additional shape formatting options.
    # @return [void]
    # @api public
    #: (?preset: String, ?text: String?, ?from_column: Integer?, ?from_row: Integer, ?to_column: Integer?, ?to_row: Integer, **untyped opts) -> void
    def shape(preset: "rect", text: nil, from_column: nil, from_row: 0, to_column: nil, to_row: 5, **opts)
      fc = from_column || 0
      tc = to_column || 5
      shape_data = { preset: preset, text: text, from_column: fc, from_row: from_row, to_column: tc, to_row: to_row }
      shape_data[:name] = opts.delete(:name) || "Shape #{@shapes.size + 1}"
      shape_data.merge!(opts)
      @shapes << shape_data
    end

    # Gets the sheet visibility state (:visible, :hidden, or :very_hidden).
    #
    # @return [Symbol]
    # @api public
    #: Symbol
    attr_reader :state

    # Sets the sheet visibility state (:visible, :hidden, or :very_hidden).
    #
    # @param new_state [Symbol, String] The new state.
    # @return [Symbol]
    # @api public
    #: (Symbol | String new_state) -> Symbol
    def state=(new_state)
      val = new_state.to_sym
      raise ArgumentError, "Invalid sheet state ':#{val}'. Allowed states are :visible, :hidden, :very_hidden" unless VALID_SHEET_STATES.include?(val)

      @state = val
    end

    # Gets the worksheet tab color.
    #
    # @return [String, Hash, nil]
    # @api public
    #: () -> (String | Hash[Symbol, untyped] | nil)
    def tab_color
      @sheet_properties[:tab_color]
    end

    # Sets the tab color of this worksheet.
    #
    # @param spec [String, Hash, Symbol, nil] Tab color hex string (e.g. "FF0000") or hash options.
    # @return [String, Hash, nil]
    # @api public
    #: (String | Symbol | Hash[Symbol, untyped] | nil spec) -> (String | Hash[Symbol, untyped] | nil)
    def tab_color=(spec)
      @sheet_properties[:tab_color] = spec
    end

    # Sets a single sheet property.
    #
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
      return Elements::SheetProperties.new(**@sheet_properties) if opts.empty?

      opts.each do |k, v|
        sym = k.to_sym
        if sym == :state
          val = v.to_sym
          raise ArgumentError, "Invalid sheet state ':#{val}'. Allowed states are :visible, :hidden, :very_hidden" unless VALID_SHEET_STATES.include?(val)

          @state = val
        elsif sym == :hidden
          @state = v ? :hidden : :visible
        else
          @sheet_properties[sym] = v
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
      h = opts.respond_to?(:to_h) ? opts.to_h : (opts || {})
      @sheet_properties = {}
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
      return Elements::SheetView.new(**@sheet_view) if opts.empty?

      opts.each { |k, v| @sheet_view[k.to_sym] = v }
      nil
    end

    # Replaces sheet view properties with the given hash or {Elements::SheetView} object.
    #
    # @param opts [Hash, Elements::SheetView]
    # @return [void]
    # @api public
    #: (Hash[Symbol, untyped] | Elements::SheetView opts) -> void
    def sheet_view=(opts)
      h = opts.respond_to?(:to_h) ? opts.to_h : (opts || {})
      @sheet_view = h.transform_keys(&:to_sym)
      nil
    end

    # Configures or queries sheet format properties (<sheetFormatPr>).
    #
    # When called with no arguments, returns the {Elements::SheetFormat} value object.
    # When called with keyword options, configures format properties and returns nil.
    #
    # @example Using keyword arguments
    #   sheet.sheet_format(default_row_height: 24, default_column_width: 15)
    #
    # @param opts [Hash] Keyword arguments for sheet format properties.
    # @return [Elements::SheetFormat, void]
    # @api public
    #: () -> Elements::SheetFormat
    #: (**untyped opts) -> void
    def sheet_format(**opts)
      return Elements::SheetFormat.new(**@sheet_format) if opts.empty?

      opts.each { |k, v| @sheet_format[k.to_sym] = v }
      nil
    end

    # Replaces sheet format properties with the given hash or {Elements::SheetFormat} object.
    #
    # @param opts [Hash, Elements::SheetFormat]
    # @return [void]
    # @api public
    #: (Hash[Symbol, untyped] | Elements::SheetFormat opts) -> void
    def sheet_format=(opts)
      h = opts.respond_to?(:to_h) ? opts.to_h : (opts || {})
      @sheet_format = h.transform_keys(&:to_sym)
      nil
    end

    # Inserts a horizontal page break after the specified 0-based row.
    #
    # @param row_index [Integer] 0-based row index.
    # @return [void]
    # @api public
    #: (Integer row_index) -> void
    def row_break(row_index)
      @row_breaks << { id: row_index, max: 16_383, man: 1 }
    end

    # Inserts a vertical page break after the specified column.
    #
    # @param col [Integer, String, Symbol] Column index (0-based) or letter ("B", :B).
    # @return [void]
    # @api public
    #: (Integer | String | Symbol col) -> void
    def column_break(col)
      col_index = Elements::Cell._column_index(col)
      @col_breaks << { id: col_index, max: 1_048_575, man: 1 }
    end

    # Adds a protected range to the worksheet.
    #
    # @param name [String] Name of the protected range.
    # @param sqref [String] Cell range (e.g. "A1:B10").
    # @param password [String, nil] Optional password or hash.
    # @param opts [Hash] Additional options.
    # @return [void]
    # @api public
    #: (String name, String sqref, ?password: String?, **untyped opts) -> void
    def protected_range(name, sqref, password: nil, **opts)
      entry = { name: name.to_s, sqref: sqref.to_s }
      entry[:password] = password if password
      entry.merge!(opts) unless opts.empty?
      @protected_ranges << DslHelpers.normalize_protection_options(entry)
    end

    # Adds a cell watch to the worksheet.
    #
    # @param ref [String] Cell reference (e.g. "A1").
    # @return [void]
    # @api public
    #: (String ref) -> void
    def cell_watch(ref)
      @cell_watches << ref.to_s
    end

    # Adds an ignored error entry for the worksheet.
    #
    # @param sqref [String] Cell range reference (e.g. "A1" or "A1:B10").
    # @param opts [Hash] Error check flags to ignore (e.g. number_stored_as_text: true, eval_error: true).
    # @return [void]
    # @api public
    #: (String sqref, ?eval_error: bool, ?two_digit_text_year: bool, ?number_stored_as_text: bool, ?formula: bool, ?formula_range: bool, ?unlocked_formula: bool, ?empty_cell_reference: bool, ?list_data_validation: bool, ?calculated_column: bool) -> void
    def ignored_error(sqref, **opts)
      @ignored_errors << { sqref: sqref.to_s }.merge(opts)
    end

    # Returns ignored error rules configured for this worksheet.
    #
    # @return [Array<Elements::IgnoredError>]
    # @api public
    #: () -> Array[Elements::IgnoredError]
    def ignored_errors
      @ignored_errors.map { |ie| ie.is_a?(Elements::IgnoredError) ? ie : Elements::IgnoredError.new(**ie) }.freeze
    end

    # Gets phonetic (furigana) guide properties for the sheet.
    #
    # @return [Elements::PhoneticProperties, nil]
    # @api public
    #: () -> Elements::PhoneticProperties?
    def phonetic_properties
      return nil unless @phonetic_properties

      @phonetic_properties.is_a?(Elements::PhoneticProperties) ? @phonetic_properties : Elements::PhoneticProperties.new(**@phonetic_properties)
    end

    # Sets phonetic (furigana) guide properties for the sheet.
    #
    # @param spec [Hash, Integer, nil] Options hash or font ID integer.
    # @return [Hash[Symbol, untyped], Integer, nil]
    # @api public
    #: (Hash[Symbol, untyped] | Integer | nil spec) -> (Hash[Symbol, untyped] | Integer | nil)
    def phonetic_properties=(spec)
      if spec.nil?
        @phonetic_properties = nil
      else
        props = if spec.is_a?(Hash)
                  spec.dup
                elsif spec.is_a?(Integer)
                  { font_id: spec }
                else
                  {}
                end
        @phonetic_properties = WorksheetBuilder.normalize_phonetic_properties(props)
      end
      spec
    end

    # Normalizes phonetic properties to OOXML schema values.
    #
    # @param props [Hash, nil]
    # @return [Hash[Symbol, untyped], nil]
    # @api private
    #: (Hash[untyped, untyped]? props) -> Hash[Symbol, untyped]?
    def self.normalize_phonetic_properties(props)
      return nil if props.nil? || props.empty?

      raw_font = props[:font_id] || props["font_id"] || props[:fontId] || props["fontId"] || 0
      type = props[:type] || props["type"]
      type_str = case type&.to_s
                 when "hiragana", "Hiragana" then "Hiragana"
                 when "halfwidth_katakana", "halfwidthKatakana" then "halfwidthKatakana"
                 when "fullwidth_katakana", "fullwidthKatakana" then "fullwidthKatakana"
                 when "no_conversion", "noConversion" then "noConversion"
                 else type&.to_s
                 end

      align = props[:alignment] || props["alignment"]
      align_str = case align&.to_s
                  when "no_control", "noControl" then "noControl"
                  when "left" then "left"
                  when "center" then "center"
                  when "distributed" then "distributed"
                  else align&.to_s
                  end

      res = { font_id: raw_font.to_i }
      res[:type] = type_str if type_str
      res[:alignment] = align_str if align_str
      res
    end

    # Gets the internal code name of the worksheet.
    #
    # @return [String, nil]
    # @api public
    #: () -> String?
    def code_name
      @sheet_properties[:code_name]
    end

    # Sets the internal code name of the worksheet.
    #
    # @param val [String, Symbol, nil] The code name.
    # @return [String, nil]
    # @api public
    #: (String | Symbol | nil val) -> String?
    def code_name=(val)
      @sheet_properties[:code_name] = val&.to_s
    end

    # Gets worksheet calculation properties (<sheetCalcPr>).
    #
    # @return [Elements::SheetCalcProperties, nil]
    # @api public
    #: () -> Elements::SheetCalcProperties?
    def sheet_calc_properties
      calc = @sheet_calc_pr || (@sheet_properties[:full_calc_on_load] ? { full_calc_on_load: true } : nil)
      return nil unless calc

      calc.is_a?(Elements::SheetCalcProperties) ? calc : Elements::SheetCalcProperties.new(**calc)
    end

    # Sets worksheet calculation properties (<sheetCalcPr>).
    #
    # @param spec [Hash, nil] Calculation properties hash.
    # @return [Hash, nil]
    # @api public
    #: (Hash[Symbol, untyped]? spec) -> Hash[Symbol, untyped]?
    def sheet_calc_properties=(spec)
      if spec.nil?
        @sheet_properties.delete(:full_calc_on_load)
        @sheet_calc_pr = nil
      else
        pr = spec.is_a?(Hash) ? spec.dup : {}
        @sheet_properties[:full_calc_on_load] = pr[:full_calc_on_load] unless pr[:full_calc_on_load].nil?
        @sheet_calc_pr = pr
      end
      spec
    end

    # Builds and returns the in-memory {Elements::Worksheet}. Cells carry their style
    # names; {WorkbookBuilder#build} resolves them to style indexes.
    #
    # @deprecated Read the sheet from the workbook that {Xlsxrb.build} returns. This
    #   method will be private in version 0.3.0.
    # @return [Elements::Worksheet]
    # @api public
    #: () -> Elements::Worksheet
    def build
      Deprecation.warn("WorksheetBuilder#build is deprecated; read the sheet from the workbook that Xlsxrb.build returns. It will be private in version 0.3.0.")
      snapshot(@rows.dup.freeze).worksheet
    end

    # Ends the sheet block: rows can no longer be added, and the sheet's state is taken
    # as it is now.
    #
    # @return [Snapshot]
    # @api private
    #: () -> Snapshot
    def close
      @closed = true
      snapshot(@rows.freeze)
    end

    # Internal: returns styles for later processing by WorkbookBuilder
    # @return [Hash{String => StyleBuilder}]
    #: Hash[String, StyleBuilder]
    attr_reader :styles

    private

    #: () -> void
    def check_open!
      raise Error, "Sheet '#{@name}' is closed: add its rows inside its sheet block" if @closed
    end

    # The style that a Date or Time cell without a cell style inherits, to give the cell
    # in its place: the row style, or the sheet's column style in a row without one. The
    # workbook then gives the cell a variant of that style with the default date or
    # date-time format, so it keeps the style's font, fill, and border, as StreamWriter
    # does. Returns nil for a style index or for no style.
    #: (untyped row_style, Integer column_index) -> (String | Symbol | Hash[Symbol, untyped])?
    def inherited_date_style(row_style, column_index)
      style = if !row_style.nil?
                row_style
              elsif @column_style_clocks.key?(column_index)
                @columns.find { |column| column.index == column_index }&.style_index
              end
      style if style.is_a?(String) || style.is_a?(Symbol) || style.is_a?(Hash)
    end

    # Returns the number format code of the sheet style named +style_name+, if it has one.
    #: (untyped style_name) -> String?
    def format_code_of(style_name)
      number_format = @styles[style_name]&.number_format
      return nil unless number_format

      number_format.is_a?(String) ? number_format : Ooxml::Utils::BUILTIN_NUM_FMT_CODES[number_format]
    end

    #: (Array[PendingRow] rows) -> Snapshot
    def snapshot(rows)
      facade_meta = {}
      facade_meta[:hyperlinks] = @hyperlinks unless @hyperlinks.empty?
      facade_meta[:auto_filter] = @auto_filter if @auto_filter
      facade_meta[:filter_columns] = @filter_columns unless @filter_columns.empty?
      facade_meta[:sort_state] = @sort_state if @sort_state
      facade_meta[:data_validations] = @data_validations unless @data_validations.empty?
      facade_meta[:conditional_formats] = @conditional_formats unless @conditional_formats.empty?
      facade_meta[:tables] = @tables unless @tables.empty?
      facade_meta[:pivot_tables] = @pivot_tables unless (@pivot_tables || []).empty?
      facade_meta[:comments] = @comments unless @comments.empty?
      facade_meta[:threaded_comments] = @threaded_comments unless @threaded_comments.empty?
      facade_meta[:sparkline_groups] = @sparkline_groups unless @sparkline_groups.empty?
      facade_meta[:merge_cells] = @merge_cells_ranges unless @merge_cells_ranges.empty?
      facade_meta[:freeze_pane] = @freeze_pane if @freeze_pane
      facade_meta[:split_pane] = @split_pane if @split_pane
      facade_meta[:selection] = @selection if @selection
      facade_meta[:page_margins] = @page_margins if @page_margins
      facade_meta[:page_setup] = @page_setup unless @page_setup.empty?
      facade_meta[:header_footer] = @header_footer unless @header_footer.empty?
      facade_meta[:print_options] = @print_options unless @print_options.empty?
      facade_meta[:sheet_protection] = @sheet_protection if @sheet_protection
      facade_meta[:images] = @images unless @images.empty?
      facade_meta[:shapes] = @shapes unless @shapes.empty?
      facade_meta[:sheet_properties] = @sheet_properties unless @sheet_properties.empty?
      facade_meta[:sheet_format] = @sheet_format unless @sheet_format.empty?
      facade_meta[:sheet_view] = @sheet_view unless @sheet_view.empty?
      facade_meta[:row_breaks] = @row_breaks unless @row_breaks.empty?
      facade_meta[:column_breaks] = @col_breaks unless @col_breaks.empty?
      facade_meta[:protected_ranges] = @protected_ranges unless @protected_ranges.empty?
      facade_meta[:cell_watches] = @cell_watches unless @cell_watches.empty?
      facade_meta[:ignored_errors] = @ignored_errors unless @ignored_errors.empty?
      facade_meta[:phonetic_properties] = @phonetic_properties if @phonetic_properties
      facade_meta[:sheet_calc_pr] = @sheet_calc_pr if @sheet_calc_pr
      facade_meta[:background_image] = @background_image if @background_image
      facade_meta[:print_area] = @print_area if @print_area
      facade_meta[:print_titles] = @print_titles if @print_titles

      hl_hash = {}
      @hyperlinks.each do |h|
        if h.is_a?(Elements::Hyperlink)
          ref_val = h.ref if h.respond_to?(:ref)
          ref_val ||= h.cell if h.respond_to?(:cell)
          ref = ref_val.to_s.upcase
          hl_hash[ref] = h unless ref.empty?
        else
          ref = (h[:cell] || h[:ref]).to_s.upcase
          entry = {}
          entry[:url] = h[:url] if h[:url]
          entry[:display] = h[:display] if h[:display]
          entry[:tooltip] = h[:tooltip] if h[:tooltip]
          entry[:location] = h[:location] if h[:location]
          hl_hash[ref] = Elements::Hyperlink.new(**entry) unless entry.empty?
        end
      end

      cm_hash = {}
      @comments.each do |c|
        obj = c.is_a?(Elements::Comment) ? c : Elements::Comment.new(**c)
        ref = (obj.respond_to?(:ref) ? obj.ref : obj[:ref]).to_s.upcase
        cm_hash[ref] = obj
      end

      tc_hash = {}
      @threaded_comments.each do |c|
        obj = c.is_a?(Elements::ThreadedComment) ? c : Elements::ThreadedComment.new(**c)
        ref = (obj.respond_to?(:ref) ? obj.ref : obj[:ref]).to_s.upcase
        (tc_hash[ref] ||= []) << obj
      end

      frame = Elements::Worksheet.new(
        name: @name, rows: [], columns: @columns, charts: @charts,
        conditional_formats: @conditional_formats,
        data_validations: @data_validations,
        unmapped_data: facade_meta.empty? ? {} : { facade: facade_meta },
        state: @state,
        hyperlinks: hl_hash,
        comments: @comments,
        threaded_comments: @threaded_comments,
        print_area: @print_area,
        print_titles: @print_titles,
        sparkline_groups: @sparkline_groups,
        merged_cells: @merge_cells_ranges,
        auto_filter: @auto_filter,
        tables: @tables,
        shapes: @shapes,
        freeze_pane: @freeze_pane,
        split_pane: @split_pane,
        filter_columns: @filter_columns,
        sort_state: @sort_state,
        page_setup: @page_setup,
        page_margins: @page_margins,
        header_footer: @header_footer,
        print_options: @print_options,
        sheet_protection: @sheet_protection,
        sheet_view: @sheet_view,
        sheet_properties: @sheet_properties,
        sheet_format: @sheet_format,
        row_breaks: @row_breaks,
        column_breaks: @col_breaks,
        protected_ranges: @protected_ranges,
        cell_watches: @cell_watches,
        ignored_errors: @ignored_errors,
        phonetic_properties: @phonetic_properties,
        selection: @selection,
        sheet_calc_properties: @sheet_calc_pr,
        background_image: @background_image
      )
      Snapshot.new(frame, rows, hl_hash, cm_hash, tc_hash, @column_style_clocks.dup)
    end
  end

  class WorksheetBuilder
    # A row as given to {WorksheetBuilder#row}, kept until the workbook resolves its
    # style names, so that its cells are created once.
    #
    # @api private
    class PendingRow
      #: Integer
      attr_reader :index
      # Cell values; Hashes with :formula and rich text runs are already converted.
      #: Array[untyped]
      attr_reader :values
      # Style of each cell: a name, an index, or nil; or one of these for every cell.
      #: untyped
      attr_reader :styles
      # Number format codes of the cells' sheet styles: an Array, one code, or nil.
      #: untyped
      attr_reader :format_codes
      # Number of columns the row spans, including styled cells past the last value.
      #: Integer
      attr_reader :width
      #: (Float | Integer)?
      attr_reader :height
      #: untyped
      attr_reader :hidden
      #: bool
      attr_reader :custom_height
      #: untyped
      attr_reader :collapsed
      #: Integer?
      attr_reader :outline_level
      # The row style as given: a name, an index, an inline Hash, or nil.
      #: untyped
      attr_reader :style
      # The style log clock when the row was added: its styles are the definitions made before.
      #: Integer
      attr_reader :clock

      #: (Integer index, Array[untyped] values, untyped styles, untyped format_codes, Integer width, (Float | Integer)? height, untyped hidden, bool custom_height, untyped collapsed, Integer? outline_level, untyped style, Integer clock) -> void
      def initialize(index, values, styles, format_codes, width, height, hidden, custom_height, collapsed, outline_level, style, clock)
        @index = index
        @values = values
        @styles = styles
        @format_codes = format_codes
        @width = width
        @height = height
        @hidden = hidden
        @custom_height = custom_height
        @collapsed = collapsed
        @outline_level = outline_level
        @style = style
        @clock = clock
      end

      #: (hidden: untyped, collapsed: untyped, outline_level: Integer?) -> PendingRow
      def with_outline(hidden:, collapsed:, outline_level:)
        PendingRow.new(@index, @values, @styles, @format_codes, @width, @height, hidden, @custom_height, collapsed, outline_level, @style, @clock)
      end

      # Returns whether every cell of the row is within Excel's row and column limits.
      #: () -> bool
      def within_limits?
        Elements::Cell.valid_coordinates?(@index, 0) && @width <= 16_384
      end

      # Returns the row's cells with their style names, as {WorksheetBuilder#build} gives them.
      #: () -> Array[Elements::Cell]
      def named_cells
        values = @values
        styles = @styles
        codes = @format_codes
        style_lookup = styles.is_a?(Array)
        code_lookup = codes.is_a?(Array)
        all_valid = within_limits?
        cells = [] #: Array[Elements::Cell]
        col_index = 0
        while col_index < @width
          value = col_index < values.size ? values[col_index] : nil
          style = style_lookup ? styles[col_index] : styles
          unless value.nil? && style.nil?
            code = code_lookup ? codes[col_index] : codes
            cells << PendingRow.cell(@index, col_index, value, style, code, all_valid || Elements::Cell.valid_coordinates?(@index, col_index))
          end
          col_index += 1
        end
        cells
      end

      #: (Array[Elements::Cell] cells, untyped style_index) -> Elements::Row
      def to_row(cells, style_index)
        Elements::Row.new(
          index: @index,
          cells: cells,
          height: @height,
          hidden: @hidden,
          custom_height: @custom_height,
          collapsed: @collapsed,
          outline_level: @outline_level,
          style_index: style_index,
          unmapped_data: Elements::EMPTY_HASH
        )
      end

      # Creates a cell for +value+ with +style+ (a name, an index, or nil).
      # +valid_coordinates+ tells whether the cell is within Excel's limits; a valid cell
      # has no errors, so it skips the checks of Cell.new.
      #: (Integer row_index, Integer column_index, untyped value, untyped style, String? format_code, bool valid_coordinates) -> Elements::Cell
      def self.cell(row_index, column_index, value, style, format_code, valid_coordinates)
        formula = nil
        if value.is_a?(Elements::Formula)
          formula = value
          value = formula.cached_value
          valid_coordinates &&= Elements::Cell.valid_value?(value)
        end
        raw_value = value&.to_s
        if valid_coordinates && !style.is_a?(Integer)
          Elements::Cell.fast_create(row_index, column_index, value, style, formula, raw_value, format_code)
        else
          Elements::Cell.new(row_index: row_index, column_index: column_index, value: value, formula: formula,
                             style: style, raw_value: raw_value, format_code: format_code)
        end
      end
    end

    # The state of a sheet when its block ended: the worksheet without its rows, the
    # pending rows, and the hyperlinks and comments to attach to their cells.
    #
    # @api private
    class Snapshot
      # The worksheet with every part except its rows.
      #: Elements::Worksheet
      attr_reader :frame
      #: Array[PendingRow]
      attr_reader :rows
      # The style log clock when each column's style was set.
      #: Hash[Integer, Integer]
      attr_reader :column_style_clocks

      #: (Elements::Worksheet frame, Array[PendingRow] rows, Hash[String, untyped] hyperlinks, Hash[String, untyped] comments, Hash[String, Array[untyped]] threaded_comments, Hash[Integer, Integer] column_style_clocks) -> void
      def initialize(frame, rows, hyperlinks, comments, threaded_comments, column_style_clocks)
        @frame = frame
        @rows = rows
        @column_style_clocks = column_style_clocks
        @hyperlinks = hyperlinks
        @comments = comments
        @threaded_comments = threaded_comments
        @links = !(hyperlinks.empty? && comments.empty? && threaded_comments.empty?)
      end

      # Returns the worksheet with +rows+, or with the rows' named cells by default.
      #: (?Array[Elements::Row]? rows, ?Array[Elements::Column]? columns) -> Elements::Worksheet
      def worksheet(rows = nil, columns = nil)
        rows ||= @rows.map { |row| row.to_row(attach_links(row.named_cells), row.style) }
        @frame.with(rows: rows, columns: columns || @frame.columns, errors: nil)
      end

      # Attaches the sheet's hyperlinks and comments to the cells they refer to.
      #: (Array[Elements::Cell] cells) -> Array[Elements::Cell]
      def attach_links(cells)
        return cells unless @links

        cells.map do |cell|
          hl = @hyperlinks[cell.ref]
          cm = @comments[cell.ref]
          tcs = @threaded_comments[cell.ref]
          if hl || cm || tcs
            cell.with(
              hyperlink: hl || cell.hyperlink,
              comment: cm || cell.comment,
              threaded_comments: (tcs ? (cell.threaded_comments + tcs) : cell.threaded_comments)
            )
          else
            cell
          end
        end
      end
    end
  end

  # Stands in for the {Elements::Worksheet} of a sheet defined in {Xlsxrb.build}, as
  # returned by {WorkbookBuilder#sheet}. It answers the same methods with the same
  # results. The worksheet is built the first time one of them is called, or when the
  # object is frozen (as Ractor.make_shareable does), so a sheet whose return value is
  # never used costs nothing.
  #
  # @deprecated Read the sheet from the workbook that {Xlsxrb.build} returns. In version
  #   0.3.0, {WorkbookBuilder#sheet} will return the sheet's {WorksheetBuilder}, and
  #   {WorkbookBuilder#[]} will return the {WorksheetBuilder} of an existing sheet.
  # @api public
  class DeferredWorksheet
    DEPRECATION_MESSAGE = "Using the sheet that WorkbookBuilder#sheet and #[] return is deprecated; " \
                          "read it from the workbook that Xlsxrb.build returns. In version 0.3.0, sheet will return " \
                          "the sheet's WorksheetBuilder, and #[] will return the WorksheetBuilder of an existing sheet."

    #: (WorksheetBuilder::Snapshot snapshot) -> void
    def initialize(snapshot)
      @snapshot = snapshot
      @built_worksheet = nil
    end

    # Returns the worksheet, building it on first use.
    #
    # @return [Elements::Worksheet]
    # @api public
    #: () -> Elements::Worksheet
    def worksheet = built_worksheet

    #: (untyped other) -> bool
    def ==(other)
      built_worksheet == (other.is_a?(DeferredWorksheet) ? other.built_worksheet : other)
    end

    #: (untyped other) -> bool
    def eql?(other)
      built_worksheet.eql?(other.is_a?(DeferredWorksheet) ? other.built_worksheet : other)
    end

    #: () -> Integer
    def hash = built_worksheet.hash

    #: () -> String
    def inspect = built_worksheet.inspect

    #: () -> String
    def to_s = built_worksheet.to_s

    #: (untyped printer) -> untyped
    def pretty_print(printer) = printer.pp(built_worksheet)

    # Builds the worksheet before freezing, so that a frozen stand-in still answers.
    #: () -> self
    def freeze
      built_worksheet
      @snapshot = nil
      super
    end

    # Elements::Worksheet methods, written out as the method_missing policy in
    # ARCHITECTURE.md requires. A test checks that they cover every public method.

    #: (String | Symbol | Integer ref_or_row, ?Integer | String | Symbol | nil col) -> Elements::Cell?
    def [](ref_or_row, col = nil) = built_worksheet[ref_or_row, col]

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def all?(...) = built_worksheet.all?(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def any?(...) = built_worksheet.any?(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def auto_filter(...) = built_worksheet.auto_filter(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def background_image(...) = built_worksheet.background_image(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def cell_watches(...) = built_worksheet.cell_watches(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def cells(...) = built_worksheet.cells(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def chain(...) = built_worksheet.chain(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def charts(...) = built_worksheet.charts(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def chunk(...) = built_worksheet.chunk(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def chunk_while(...) = built_worksheet.chunk_while(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def code_name(...) = built_worksheet.code_name(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def collect(...) = built_worksheet.collect(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def collect_concat(...) = built_worksheet.collect_concat(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def column_breaks(...) = built_worksheet.column_breaks(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def column_range(...) = built_worksheet.column_range(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def columns(...) = built_worksheet.columns(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def comment(...) = built_worksheet.comment(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def comments(...) = built_worksheet.comments(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def comments_by_ref(...) = built_worksheet.comments_by_ref(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def compact(...) = built_worksheet.compact(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def conditional_formats(...) = built_worksheet.conditional_formats(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def count(...) = built_worksheet.count(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def cycle(...) = built_worksheet.cycle(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def data_validations(...) = built_worksheet.data_validations(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def date1904?(...) = built_worksheet.date1904?(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def deconstruct(...) = built_worksheet.deconstruct(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def deconstruct_keys(...) = built_worksheet.deconstruct_keys(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def detect(...) = built_worksheet.detect(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def dimension(...) = built_worksheet.dimension(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def dimension_bounds(...) = built_worksheet.dimension_bounds(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def drop(...) = built_worksheet.drop(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def drop_while(...) = built_worksheet.drop_while(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def each(...) = built_worksheet.each(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def each_cell(...) = built_worksheet.each_cell(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def each_cons(...) = built_worksheet.each_cons(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def each_entry(...) = built_worksheet.each_entry(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def each_row(...) = built_worksheet.each_row(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def each_row_values(...) = built_worksheet.each_row_values(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def each_slice(...) = built_worksheet.each_slice(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def each_with_index(...) = built_worksheet.each_with_index(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def each_with_object(...) = built_worksheet.each_with_object(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def entries(...) = built_worksheet.entries(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def errors(...) = built_worksheet.errors(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def filter(...) = built_worksheet.filter(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def filter_columns(...) = built_worksheet.filter_columns(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def filter_map(...) = built_worksheet.filter_map(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def find(...) = built_worksheet.find(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def find_all(...) = built_worksheet.find_all(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def find_index(...) = built_worksheet.find_index(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def first(...) = built_worksheet.first(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def first_row(...) = built_worksheet.first_row(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def fit_to_page?(...) = built_worksheet.fit_to_page?(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def flat_map(...) = built_worksheet.flat_map(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def formatted_value(...) = built_worksheet.formatted_value(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def freeze_pane(...) = built_worksheet.freeze_pane(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def full_calc_on_load?(...) = built_worksheet.full_calc_on_load?(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def grep(...) = built_worksheet.grep(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def grep_v(...) = built_worksheet.grep_v(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def group_by(...) = built_worksheet.group_by(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def header_footer(...) = built_worksheet.header_footer(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def header_footer_images(...) = built_worksheet.header_footer_images(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def hidden?(...) = built_worksheet.hidden?(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def hyperlink(...) = built_worksheet.hyperlink(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def hyperlinks(...) = built_worksheet.hyperlinks(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def ignored_errors(...) = built_worksheet.ignored_errors(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def images(...) = built_worksheet.images(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def images_at(...) = built_worksheet.images_at(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def include?(...) = built_worksheet.include?(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def inject(...) = built_worksheet.inject(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def last_row(...) = built_worksheet.last_row(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def lazy(...) = built_worksheet.lazy(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def map(...) = built_worksheet.map(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def max(...) = built_worksheet.max(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def max_by(...) = built_worksheet.max_by(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def member?(...) = built_worksheet.member?(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def merged_cells(...) = built_worksheet.merged_cells(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def min(...) = built_worksheet.min(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def min_by(...) = built_worksheet.min_by(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def minmax(...) = built_worksheet.minmax(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def minmax_by(...) = built_worksheet.minmax_by(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def name(...) = built_worksheet.name(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def none?(...) = built_worksheet.none?(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def one?(...) = built_worksheet.one?(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def outline_properties(...) = built_worksheet.outline_properties(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def pad_empty_cells?(...) = built_worksheet.pad_empty_cells?(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def pad_empty_rows?(...) = built_worksheet.pad_empty_rows?(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def page_margins(...) = built_worksheet.page_margins(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def page_setup(...) = built_worksheet.page_setup(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def partition(...) = built_worksheet.partition(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def phonetic_properties(...) = built_worksheet.phonetic_properties(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def pivot_tables(...) = built_worksheet.pivot_tables(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def print_area(...) = built_worksheet.print_area(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def print_options(...) = built_worksheet.print_options(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def print_titles(...) = built_worksheet.print_titles(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def protected_ranges(...) = built_worksheet.protected_ranges(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def reduce(...) = built_worksheet.reduce(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def reject(...) = built_worksheet.reject(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def reverse_each(...) = built_worksheet.reverse_each(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def row_at(...) = built_worksheet.row_at(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def row_breaks(...) = built_worksheet.row_breaks(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def row_range(...) = built_worksheet.row_range(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def rows(...) = built_worksheet.rows(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def select(...) = built_worksheet.select(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def selection(...) = built_worksheet.selection(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def shapes(...) = built_worksheet.shapes(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def sheet_calc_properties(...) = built_worksheet.sheet_calc_properties(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def sheet_format(...) = built_worksheet.sheet_format(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def sheet_properties(...) = built_worksheet.sheet_properties(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def sheet_protection(...) = built_worksheet.sheet_protection(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def sheet_view(...) = built_worksheet.sheet_view(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def slice_after(...) = built_worksheet.slice_after(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def slice_before(...) = built_worksheet.slice_before(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def slice_when(...) = built_worksheet.slice_when(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def sort(...) = built_worksheet.sort(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def sort_by(...) = built_worksheet.sort_by(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def sort_state(...) = built_worksheet.sort_state(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def sparkline_groups(...) = built_worksheet.sparkline_groups(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def split_pane(...) = built_worksheet.split_pane(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def state(...) = built_worksheet.state(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def styles(...) = built_worksheet.styles(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def sum(...) = built_worksheet.sum(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def tab_color(...) = built_worksheet.tab_color(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def tables(...) = built_worksheet.tables(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def take(...) = built_worksheet.take(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def take_while(...) = built_worksheet.take_while(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def tally(...) = built_worksheet.tally(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def threaded_comment(...) = built_worksheet.threaded_comment(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def threaded_comments(...) = built_worksheet.threaded_comments(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def threaded_comments_at(...) = built_worksheet.threaded_comments_at(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def threaded_comments_by_ref(...) = built_worksheet.threaded_comments_by_ref(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def to_a(...) = built_worksheet.to_a(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def to_h(...) = built_worksheet.to_h(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def to_set(...) = built_worksheet.to_set(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def trim_empty_rows?(...) = built_worksheet.trim_empty_rows?(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def uniq(...) = built_worksheet.uniq(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def unmapped_data(...) = built_worksheet.unmapped_data(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def valid?(...) = built_worksheet.valid?(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def very_hidden?(...) = built_worksheet.very_hidden?(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def visible?(...) = built_worksheet.visible?(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def with(...) = built_worksheet.with(...)

    #: (*untyped, **untyped) ?{ (*untyped) -> untyped } -> untyped
    def zip(...) = built_worksheet.zip(...)

    protected

    # Returns the worksheet, building it on first use with a deprecation warning that
    # points at the caller of the public method.
    #: () -> Elements::Worksheet
    def built_worksheet
      @built_worksheet ||= begin
        Deprecation.warn(DEPRECATION_MESSAGE, uplevel: 2)
        @snapshot&.worksheet || raise(Error, "worksheet is not available")
      end
    end
  end

  # Builder context for configuring an Excel Table (ListObject) in block form.
  #
  # @api public
  class TableBuilder
    # @param ref [String, nil] Table range (e.g. "A1:D10").
    # @param columns [Array, nil] Column header names or definition hashes.
    # @param name [String, nil] Table name identifier.
    # @param display_name [String, nil] Table display name.
    # @param style [String, nil] Built-in table style name (e.g. "TableStyleMedium2").
    # @param opts [Hash] Additional table properties.
    #: (?String? ref, ?columns: untyped, ?name: String?, ?display_name: String?, ?style: String?, **untyped opts) -> void
    def initialize(ref = nil, columns: nil, name: nil, display_name: nil, style: nil, **opts)
      @options = opts.dup
      @options[:ref] = ref if ref
      @options[:columns] = columns.is_a?(Array) ? columns.dup : [columns] if columns
      @options[:name] = name if name
      @options[:display_name] = display_name if display_name
      @options[:style] = style if style
    end

    #: () -> String?
    def ref
      @options[:ref]
    end

    #: (String val) -> String
    def ref=(val)
      @options[:ref] = val.to_s
    end

    # Returns column definitions for the table.
    #
    # @return [Array<Elements::TableColumn>]
    # @api public
    #: () -> Array[Elements::TableColumn]
    def columns
      (@options[:columns] || []).each_with_index.map do |c, idx|
        if c.is_a?(Elements::TableColumn)
          c
        elsif c.is_a?(Hash)
          Elements::TableColumn.new(id: idx + 1, **c)
        else
          Elements::TableColumn.new(id: idx + 1, name: c.to_s)
        end
      end
    end

    #: (untyped cols) -> untyped
    def columns=(cols)
      @options[:columns] = cols.is_a?(Array) ? cols : [cols]
    end

    # Appends a single column name or column definition.
    #
    # @param col [String, Hash, Symbol] Column header name or definition.
    # @param opts [Hash] Column options (e.g. totals_row_function, totals_row_label, formula, totals_row_formula).
    # @return [void]
    #: (String | Symbol | Hash[Symbol, untyped] col, **untyped opts) -> void
    def column(col, **opts)
      col_spec = if opts.empty?
                   col
                 elsif col.is_a?(Hash)
                   col.merge(opts)
                 else
                   opts.merge(name: col.to_s)
                 end
      (@options[:columns] ||= []) << col_spec
      nil
    end

    # Appends a calculated column with formula and options.
    #
    # @param name [String, Symbol] The column name.
    # @param formula [String] The formula for the calculated column (e.g. "[@Qty]*[@Price]").
    # @param opts [Hash] Additional column options.
    # @return [void]
    #: (String | Symbol name, String formula, **untyped opts) -> void
    def calculated_column(name, formula, **opts)
      column(name, formula: formula, **opts)
    end

    #: () -> String?
    def name
      @options[:name]
    end

    #: (String val) -> String
    def name=(val)
      @options[:name] = val.to_s
    end

    #: () -> String?
    def display_name
      @options[:display_name]
    end

    #: (String val) -> String
    def display_name=(val)
      @options[:display_name] = val.to_s
    end

    #: () -> String?
    def style
      @options[:style]
    end

    #: (String val) -> String
    def style=(val)
      @options[:style] = val.to_s
    end

    #: () -> bool
    def show_row_stripes?
      @options[:show_row_stripes] ? true : false
    end

    #: (bool val) -> bool
    def show_row_stripes=(val)
      @options[:show_row_stripes] = val ? true : false
    end

    #: () -> bool
    def show_column_stripes?
      @options[:show_column_stripes] ? true : false
    end

    #: (bool val) -> bool
    def show_column_stripes=(val)
      @options[:show_column_stripes] = val ? true : false
    end

    #: () -> bool
    def show_first_column?
      @options[:show_first_column] ? true : false
    end

    #: (bool val) -> bool
    def show_first_column=(val)
      @options[:show_first_column] = val ? true : false
    end

    #: () -> bool
    def show_last_column?
      @options[:show_last_column] ? true : false
    end

    #: (bool val) -> bool
    def show_last_column=(val)
      @options[:show_last_column] = val ? true : false
    end

    #: () -> Integer?
    def totals_row_count
      @options[:totals_row_count]
    end

    #: (Integer val) -> Integer
    def totals_row_count=(val)
      @options[:totals_row_count] = val.to_i
    end

    # Returns whether the table has a totals row.
    #
    # @return [Boolean]
    #: () -> bool
    def totals_row?
      rc = @options[:totals_row_count]
      if rc
        rc.positive?
      else
        @options[:totals_row] ? true : false
      end
    end

    #: (bool val) -> bool
    def totals_row=(val)
      @options[:totals_row_count] = val ? 1 : 0
      val
    end

    #: () -> bool
    def auto_filter?
      @options[:auto_filter] != false
    end

    #: (bool val) -> bool
    def auto_filter=(val)
      @options[:auto_filter] = val ? true : false
    end

    #: () -> Hash[Symbol, untyped]
    def to_h
      @options.dup
    end
  end

  # Builder context for configuring data validation rules in block form.
  #
  # @api public
  class DataValidationBuilder
    # @param sqref [String, nil] Target cell or range reference.
    # @param opts [Hash] Additional validation properties.
    #: (?String? sqref, **untyped opts) -> void
    def initialize(sqref = nil, **opts)
      @options = opts.dup
      @options[:sqref] = sqref if sqref
    end

    #: () -> String?
    def sqref
      @options[:sqref]
    end

    #: (String val) -> String
    def sqref=(val)
      @options[:sqref] = val.to_s
    end

    #: () -> untyped
    def columns
      @options[:columns]
    end

    #: (untyped val) -> untyped
    def columns=(val)
      @options[:columns] = val
    end

    #: () -> untyped
    def rows
      @options[:rows]
    end

    #: (untyped val) -> untyped
    def rows=(val)
      @options[:rows] = val
    end

    #: () -> untyped
    def except_rows
      @options[:except_rows]
    end

    #: (untyped val) -> untyped
    def except_rows=(val)
      @options[:except_rows] = val
    end

    # Sets or gets the list validation choices.
    #
    # @example
    #   b.list = ["Option1", "Option2"]
    #
    # @return [Array<Object>, nil]
    # @api public
    #: () -> untyped
    def list
      @options[:list]
    end

    #: (Array[untyped] val) -> untyped
    def list=(val)
      @options[:list] = val
    end

    #: () -> untyped
    def type
      @options[:type]
    end

    #: (untyped val) -> untyped
    def type=(val)
      @options[:type] = val
    end

    #: () -> untyped
    def operator
      @options[:operator]
    end

    #: (untyped val) -> untyped
    def operator=(val)
      @options[:operator] = val
    end

    #: () -> untyped
    def formula1
      @options[:formula1]
    end

    #: (untyped val) -> untyped
    def formula1=(val)
      @options[:formula1] = val
    end

    #: () -> untyped
    def formula2
      @options[:formula2]
    end

    #: (untyped val) -> untyped
    def formula2=(val)
      @options[:formula2] = val
    end

    #: () -> bool
    def allow_blank?
      @options[:allow_blank] ? true : false
    end

    #: (bool val) -> bool
    def allow_blank=(val)
      @options[:allow_blank] = val ? true : false
    end

    #: () -> bool
    def show_input_message?
      @options[:show_input_message] ? true : false
    end

    #: (bool val) -> bool
    def show_input_message=(val)
      @options[:show_input_message] = val ? true : false
    end

    #: () -> bool
    def show_error_message?
      @options[:show_error_message] ? true : false
    end

    #: (bool val) -> bool
    def show_error_message=(val)
      @options[:show_error_message] = val ? true : false
    end

    #: () -> String?
    def error_title
      @options[:error_title]
    end

    #: (String val) -> String
    def error_title=(val)
      @options[:error_title] = val.to_s
    end

    #: () -> String?
    def error_message
      @options[:error]
    end

    #: (String val) -> String
    def error_message=(val)
      @options[:error] = val.to_s
    end

    #: () -> String?
    def prompt_title
      @options[:prompt_title]
    end

    #: (String val) -> String
    def prompt_title=(val)
      @options[:prompt_title] = val.to_s
    end

    #: () -> String?
    def prompt_message
      @options[:prompt]
    end

    #: (String val) -> String
    def prompt_message=(val)
      @options[:prompt] = val.to_s
    end

    #: () -> untyped
    def error_style
      @options[:error_style]
    end

    #: (untyped val) -> untyped
    def error_style=(val)
      @options[:error_style] = val
    end

    #: () -> Hash[Symbol, untyped]
    def to_h
      @options.dup
    end
  end

  # Builder context for configuring conditional formatting rules in block form.
  #
  # @api public
  class ConditionalFormatBuilder
    # @param sqref [String, nil] Target cell or range reference.
    # @param opts [Hash] Additional conditional formatting properties.
    #: (?String? sqref, **untyped opts) -> void
    def initialize(sqref = nil, **opts)
      @options = opts.dup
      @options[:sqref] = sqref if sqref
    end

    #: () -> String?
    def sqref
      @options[:sqref]
    end

    #: (String val) -> String
    def sqref=(val)
      @options[:sqref] = val.to_s
    end

    #: () -> untyped
    def columns
      @options[:columns]
    end

    #: (untyped val) -> untyped
    def columns=(val)
      @options[:columns] = val
    end

    #: () -> untyped
    def rows
      @options[:rows]
    end

    #: (untyped val) -> untyped
    def rows=(val)
      @options[:rows] = val
    end

    #: () -> untyped
    def except_rows
      @options[:except_rows]
    end

    #: (untyped val) -> untyped
    def except_rows=(val)
      @options[:except_rows] = val
    end

    #: () -> untyped
    def type
      @options[:type]
    end

    #: (untyped val) -> untyped
    def type=(val)
      @options[:type] = val
    end

    #: () -> untyped
    def operator
      @options[:operator]
    end

    #: (untyped val) -> untyped
    def operator=(val)
      @options[:operator] = val
    end

    #: () -> untyped
    def formula
      @options[:formula]
    end

    #: (untyped val) -> untyped
    def formula=(val)
      @options[:formula] = val
    end

    #: () -> untyped
    def font
      @options[:font]
    end

    #: (untyped val) -> untyped
    def font=(val)
      @options[:font] = val
    end

    #: () -> untyped
    def fill
      @options[:fill]
    end

    #: (untyped val) -> untyped
    def fill=(val)
      @options[:fill] = val
    end

    #: () -> untyped
    def fill_color
      @options[:fill_color]
    end

    #: (untyped val) -> untyped
    def fill_color=(val)
      @options[:fill_color] = val
    end

    #: () -> untyped
    def font_color
      @options[:font_color]
    end

    #: (untyped val) -> untyped
    def font_color=(val)
      @options[:font_color] = val
    end

    #: () -> bool
    def bold?
      @options[:bold] ? true : false
    end

    #: (bool val) -> bool
    def bold=(val)
      @options[:bold] = val ? true : false
    end

    #: () -> bool
    def italic?
      @options[:italic] ? true : false
    end

    #: (bool val) -> bool
    def italic=(val)
      @options[:italic] = val ? true : false
    end

    #: () -> untyped
    def underline
      @options[:underline]
    end

    #: (untyped val) -> untyped
    def underline=(val)
      @options[:underline] = val
    end

    #: () -> String?
    def font_name
      @options[:font_name]
    end

    #: (String val) -> String
    def font_name=(val)
      @options[:font_name] = val.to_s
    end

    #: () -> untyped
    def font_size
      @options[:font_size]
    end

    #: (untyped val) -> untyped
    def font_size=(val)
      @options[:font_size] = val
    end

    #: () -> bool
    def stop_if_true?
      @options[:stop_if_true] ? true : false
    end

    #: (bool val) -> bool
    def stop_if_true=(val)
      @options[:stop_if_true] = val ? true : false
    end

    #: () -> Integer?
    def priority
      @options[:priority]
    end

    #: (Integer val) -> Integer
    def priority=(val)
      @options[:priority] = val.to_i
    end

    #: () -> bool
    def above_average?
      @options[:above_average] ? true : false
    end

    #: (bool val) -> bool
    def above_average=(val)
      @options[:above_average] = val ? true : false
    end

    #: () -> bool
    def equal_average?
      @options[:equal_average] ? true : false
    end

    #: (bool val) -> bool
    def equal_average=(val)
      @options[:equal_average] = val ? true : false
    end

    #: () -> Integer?
    def std_dev
      @options[:std_dev]
    end

    #: (Integer val) -> Integer
    def std_dev=(val)
      @options[:std_dev] = val.to_i
    end

    #: () -> Integer?
    def rank
      @options[:rank]
    end

    #: (Integer val) -> Integer
    def rank=(val)
      @options[:rank] = val.to_i
    end

    #: () -> bool
    def percent?
      @options[:percent] ? true : false
    end

    #: (bool val) -> bool
    def percent=(val)
      @options[:percent] = val ? true : false
    end

    #: () -> bool
    def bottom?
      @options[:bottom] ? true : false
    end

    #: (bool val) -> bool
    def bottom=(val)
      @options[:bottom] = val ? true : false
    end

    #: () -> String?
    def text
      @options[:text]
    end

    #: (String val) -> String
    def text=(val)
      @options[:text] = val.to_s
    end

    #: () -> untyped
    def time_period
      @options[:time_period]
    end

    #: (untyped val) -> untyped
    def time_period=(val)
      @options[:time_period] = val
    end

    #: () -> untyped
    def color_scale
      @options[:color_scale]
    end

    #: (untyped val) -> untyped
    def color_scale=(val)
      @options[:color_scale] = val
    end

    #: () -> untyped
    def data_bar
      @options[:data_bar]
    end

    #: (untyped val) -> untyped
    def data_bar=(val)
      @options[:data_bar] = val
    end

    #: () -> untyped
    def icon_set
      @options[:icon_set]
    end

    #: (untyped val) -> untyped
    def icon_set=(val)
      @options[:icon_set] = val
    end

    #: () -> untyped
    def dxf
      @options[:dxf]
    end

    #: (untyped val) -> untyped
    def dxf=(val)
      @options[:dxf] = val
    end

    # Gets the differential format identifier (dxfId).
    #
    # @return [Integer, nil]
    # @api public
    #: () -> Integer?
    def dxf_id
      @options[:dxf_id]
    end

    # Sets the differential format identifier (dxfId).
    #
    # @param val [Integer]
    # @return [Integer]
    # @api public
    #: (Integer val) -> Integer
    def dxf_id=(val)
      @options[:dxf_id] = val.to_i
    end

    #: () -> Hash[Symbol, untyped]
    def to_h
      @options.dup
    end
  end
end
