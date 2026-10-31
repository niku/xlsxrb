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
    #: (String name, ?strict_excel_mode: bool, ?state: Symbol) -> void
    def initialize(name, strict_excel_mode: true, state: :visible)
      @name = name
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

    attr_reader :name, :defined_names

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
    # @api public
    #: (String | Symbol name, **untyped opts) ?{ (StyleBuilder) -> void } -> StyleBuilder
    def style(name, **opts)
      style_name = name.to_s
      style_builder = StyleBuilder.new(style_name)
      style_builder.apply_options!(**opts) unless opts.empty?
      yield style_builder if block_given?
      @styles[style_name] = style_builder
      style_builder
    end

    # Appends a row of cells to the worksheet.
    #
    # @example Add a row with values and array styles
    #   sheet.row(["ID", "Name", "Total"], styles: [:header, :header, :header])
    #
    # @example Add a row with a hash of column keys
    #   sheet.row({ "A" => "Invoice", "C" => 12345 })
    #
    # @param values [Array<Object>, Hash{String, Integer => Object}] The cell values.
    # @param styles [String, Symbol, Array<String | Symbol>, Hash, nil] Style names or inline styles to apply.
    # @param height [Float, Integer, nil] The row height in points (0 to 409).
    # @param hidden [Boolean] Whether the row is hidden.
    # @param custom_height [Boolean] Whether custom row height is enforced.
    # @param outline_level [Integer, nil] Grouping/outline hierarchy level.
    # @param collapsed [Boolean] Whether the row is collapsed.
    # @param style [Integer, Symbol, String, Hash, nil] Row style format or index.
    # @return [void]
    # @raise [ArgumentError] If limits are exceeded when strict_excel_mode is enabled.
    # @api public
    #: (Array[untyped] | Hash[untyped, untyped] values, ?styles: untyped, ?height: Float | Integer | nil, ?hidden: bool, ?custom_height: bool, ?collapsed: bool, ?outline_level: Integer | nil, ?style: (Integer | Symbol | String | Hash[Symbol, untyped])?) -> void
    def row(values, styles: nil, height: nil, hidden: false, custom_height: false, collapsed: false, outline_level: nil, style: nil)
      row_index = @next_row_index
      DslHelpers.validate_row_bounds!(row_index, height, strict_excel_mode: @strict_excel_mode)
      @next_row_index += 1

      pending = @pending_row_attributes.delete(row_index) || {}
      outline_level ||= @current_row_outline_level || pending[:outline_level]
      collapsed = @current_row_collapsed || pending[:collapsed] if !collapsed && (@current_row_collapsed || pending[:collapsed])
      hidden = @current_row_hidden || pending[:hidden] if !hidden && (@current_row_hidden || pending[:hidden])
      height ||= pending[:height]

      values = DslHelpers.normalize_row_values(values)
      styles = DslHelpers.normalize_row_styles(styles)

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

      max_len = values.size
      max_len = [max_len, styles.size].max if styles.is_a?(Array)
      # See: https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
      raise ArgumentError, "Row contains #{max_len} columns, exceeding Excel limit of 16_384 columns" if @strict_excel_mode && max_len > 16_384

      cells = Array.new(max_len)
      style_lookup = styles.is_a?(Array)

      col_index = 0
      while col_index < max_len
        val = col_index < values.size ? values[col_index] : nil
        raise ArgumentError, "Invalid cell value type or value: #{val.class} for value #{val.inspect}" unless val.nil? || val.is_a?(String) || (val.is_a?(Numeric) && !(val.is_a?(Float) && (val.infinite? || val.nan?))) || val.is_a?(TrueClass) || val.is_a?(FalseClass) || val.is_a?(Date) || val.is_a?(Time) || val.is_a?(Elements::Formula) || (val.is_a?(Hash) && val.key?(:formula)) || val.is_a?(Elements::RichText) || val.is_a?(Elements::CellError) || (val.is_a?(Array) && val.first.is_a?(Hash) && (val.first.key?(:text) || val.first.key?("text")))
        # See: https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
        raise ArgumentError, "Cell text length #{val.length} exceeds Excel limit of 32,767 characters" if @strict_excel_mode && val.is_a?(String) && val.length > 32_767

        if val.is_a?(Array) && val.first.is_a?(Hash) && (val.first.key?(:text) || val.first.key?("text"))
          # Coerce array of hashes to RichText
          runs = val.map do |run|
            text = run[:text] || run["text"]
            font = run.reject { |k| k.to_s == "text" }
            { text: text, font: font.empty? ? nil : font }.compact
          end
          val = Elements::RichText.new(runs: runs)
        end

        style_name = if style_lookup
                       col_index < styles.size ? styles[col_index] : nil
                     else
                       styles
                     end

        if style_name.is_a?(Hash)
          inline_name = "__inline_#{style_name.hash}"
          style(inline_name, **style_name) unless @styles.key?(inline_name)
          style_name = inline_name
        end
        if val.nil? && style_name.nil?
          col_index += 1
          next
        end
        sb = @styles[style_name]
        fmt_code = if sb&.num_fmt_id
                     sb.num_fmt_id.is_a?(String) ? sb.num_fmt_id : Ooxml::Utils::BUILTIN_NUM_FMT_CODES[sb.num_fmt_id]
                   end
        # If value is a Formula object or Hash with :formula, store it as the cell's formula
        cells[col_index] = if val.is_a?(Elements::Formula)
                             Elements::Cell.new(
                               row_index: row_index,
                               column_index: col_index,
                               value: val.cached_value,
                               formula: val,
                               style_index: style_name,
                               raw_value: val.cached_value&.to_s,
                               format_code: fmt_code
                             )
                           elsif val.is_a?(Hash) && val.key?(:formula)
                             cached_val = val.key?(:value) ? val[:value] : val[:cached_value]
                             f_obj = Elements::Formula.new(
                               expression: val[:formula],
                               cached_value: cached_val,
                               type: val[:type],
                               ref: val[:ref],
                               shared_index: val[:shared_index] || val[:si],
                               calculate_always: val[:calculate_always]
                             )
                             Elements::Cell.new(
                               row_index: row_index,
                               column_index: col_index,
                               value: cached_val,
                               formula: f_obj,
                               style_index: style_name,
                               raw_value: cached_val&.to_s,
                               format_code: fmt_code
                             )
                           else
                             Elements::Cell.new(
                               row_index: row_index,
                               column_index: col_index,
                               value: val,
                               style_index: style_name,
                               raw_value: val&.to_s,
                               format_code: fmt_code
                             )
                           end
        col_index += 1
      end

      cells.compact!
      unmapped = style ? { style_index: style } : {}
      @rows << Elements::Row.new(
        index: row_index,
        cells: cells,
        height: height,
        hidden: hidden,
        custom_height: custom_height || !height.nil?,
        collapsed: collapsed,
        outline_level: outline_level,
        unmapped_data: unmapped
      )
    end
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
    #   sheet.skip_row
    #
    # @example Skip multiple rows
    #   sheet.skip_rows(3)
    #
    # @param count [Integer] Number of rows to skip (must be >= 1).
    # @return [void]
    # @api public
    #: (?Integer count) -> void
    def skip_rows(count = 1)
      raise ArgumentError, "count must be positive: #{count.inspect}" if !count.is_a?(Integer) || count < 1

      @next_row_index += count
    end
    alias skip_row skip_rows

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
    #: (Integer | String | Range[Integer | String] | Array[Integer | String] index, ?width: Float | Integer | nil, ?hidden: bool, ?custom_width: bool, ?outline_level: Integer | nil, ?collapsed: bool, ?style: (Integer | Symbol | String | Hash[Symbol, untyped])?) -> void
    def column(index, width: nil, hidden: false, custom_width: false, outline_level: nil, collapsed: false, style: nil)
      raise ArgumentError, "Column width #{width} must be between 0 and 255 characters (Excel limitation)" if @strict_excel_mode && width && (width.negative? || width > 255)

      DslHelpers.normalize_column_indices(index).each do |idx|
        existing_idx = @columns.find_index { |c| c.index == idx }
        if existing_idx
          prev = @columns[existing_idx]
          new_width = width.nil? ? prev.width : width
          new_custom_width = (custom_width || !width.nil?) || prev.custom_width
          new_hidden = hidden || prev.hidden
          new_outline = outline_level.nil? ? prev.outline_level : outline_level
          new_style = style.nil? ? prev.style_index : style
          unmapped = prev.unmapped_data ? prev.unmapped_data.dup : {}
          unmapped[:style_index] = new_style if new_style
          unmapped[:collapsed] = true if collapsed || prev.unmapped_data&.[](:collapsed)
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
    alias group_column group_columns
    alias group_cols group_columns
    alias group_col group_columns

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
        row_indices.each do |idx|
          existing_row = @rows.find { |r| r.index == idx }
          if existing_row
            existing_idx = @rows.index(existing_row)
            @rows[existing_idx] = Elements::Row.new(
              index: existing_row.index,
              cells: existing_row.cells,
              height: existing_row.height,
              hidden: hidden || existing_row.hidden,
              custom_height: existing_row.custom_height,
              collapsed: collapsed || existing_row.collapsed,
              outline_level: level,
              unmapped_data: existing_row.unmapped_data
            )
          else
            @pending_row_attributes[idx] = { outline_level: level, collapsed: collapsed, hidden: hidden }
          end
        end
      end
    end
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
      props = (@sheet_properties[:outline_properties] || {}).dup
      props[:summary_below] = summary_below unless summary_below.nil?
      props[:summary_right] = summary_right unless summary_right.nil?
      symbols = show_outline_symbols.nil? ? show_symbols : show_outline_symbols
      props[:show_outline_symbols] = symbols unless symbols.nil?
      props[:apply_styles] = apply_styles unless apply_styles.nil?
      props.merge!(opts)
      @sheet_properties[:outline_properties] = props
      @sheet_properties[:summary_below] = props[:summary_below] unless props[:summary_below].nil?
      @sheet_properties[:summary_right] = props[:summary_right] unless props[:summary_right].nil?
      props
    end
    alias outline_pr outline_properties

    # Adds a chart to the worksheet.
    #
    # @param options [Hash] Chart configuration options.
    # @yield [builder]
    # @yieldparam builder [Xlsxrb::ChartBuilder]
    # @return [void]
    # @api public
    #: (**untyped options) ?{ (ChartBuilder) -> void } -> void
    def chart(**options)
      if block_given?
        builder = ChartBuilder.new
        yield builder
        options = builder.options.merge(options)
      end
      @charts << options
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

    # Sets an auto-filter range on the sheet.
    #
    # @param range [String] The cell range (e.g. "A1:E100").
    # @return [String]
    # @api public
    #: (String range) -> String
    def auto_filter(range)
      @auto_filter = range
    end

    # Configures filtering rules on a specific column of an auto-filter.
    #
    # @param col_id [Integer] 0-based column index relative to filter range.
    # @param filter [Hash] Filter criteria (e.g. values, custom filters).
    # @return [Hash]
    # @api public
    #: (Integer col_id, Hash[untyped, untyped] filter) -> Hash[untyped, untyped]
    def filter_column(col_id, filter)
      @filter_columns[col_id] = filter
    end

    # Configures column sort state for the sheet.
    #
    # @param ref [String] The sorted range.
    # @param sort_conditions [Array<Hash>] List of sort condition definitions.
    # @param opts [Hash] Additional sort state options.
    # @return [Hash]
    # @api public
    #: (String ref, Array[Hash[untyped, untyped]] sort_conditions, **untyped opts) -> Hash[untyped, untyped]
    def sort_state(ref, sort_conditions, **opts)
      @sort_state = { ref: ref, sort_conditions: sort_conditions }.merge(opts)
    end

    # Adds a data validation rule to a cell, range, or coordinates.
    #
    # @param sqref [String, nil] Target cell or range reference (e.g. "B2:B100").
    # @param opts [Hash] Validation configuration (e.g. type: :list, formula1: '"Option1,Option2"', or coordinate options `columns:`, `rows:`, `except_rows:`, `in:`).
    # @yield [data_validation_builder] Optional configuration block.
    # @return [void]
    # @api public
    #: (?String? sqref, **untyped opts) ?{ (DataValidationBuilder) -> void } -> void
    def validate_data(sqref = nil, **opts, &)
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
    # @param col_fields [Array<Integer, String>] Field names or indices for column headers.
    # @param dest_ref [String] Top-left destination cell (default "E1").
    # @param name [String, nil] Pivot table name.
    # @param field_names [Array<String>, nil] Custom field display names.
    # @param items [Array, nil] Item configuration.
    # @return [void]
    # @api public
    #: (String source_ref, row_fields: Array[String | Integer], data_fields: Array[String | Hash[Symbol, untyped]], ?col_fields: Array[String | Integer], ?dest_ref: String, ?name: String?, ?field_names: Array[String]?, ?items: Array[untyped]?, **untyped opts) -> void
    def pivot_table(source_ref, row_fields:, data_fields:, col_fields: [], dest_ref: "E1", name: nil, field_names: nil, items: nil)
      @pivot_tables ||= []
      @pivot_tables << {
        source_ref: source_ref, row_fields: row_fields,
        data_fields: data_fields, col_fields: col_fields,
        dest_ref: dest_ref, name: name,
        field_names: field_names, items: items
      }
    end

    # Adds a cell comment / note.
    #
    # @param cell [String, Integer] Cell reference (e.g. "B2").
    # @param text [String] The comment text.
    # @param author [String] Author name.
    # @return [void]
    # @api public
    #: (String | Integer cell, String text, ?author: ::String) -> void
    def comment(cell, text, author: "Author")
      @comments << { cell: cell, text: text, author: author }
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

    # Adds a single sparkline to the sheet.
    #
    # @example Add a line sparkline
    #   sheet.add_sparkline(location: "F1", range: "A1:E1")
    #
    # @example Add a column sparkline with high/low markers
    #   sheet.add_sparkline(location: "F2", range: "A2:E2", type: :column, high: true, low: true)
    #
    # @example Add a win/loss sparkline
    #   sheet.add_sparkline(location: "F3", range: "A3:E3", type: :win_loss)
    #
    # @param location [String, nil] Target cell for the sparkline (e.g. "F1").
    # @param range [String, nil] Data range for the sparkline (e.g. "A1:E1").
    # @param type [Symbol, String, nil] Sparkline type (:line, :column, :win_loss / :stacked).
    # @param opts [Hash] Additional sparkline options.
    # @return [void]
    # @api public
    #: (?location: String?, ?range: String?, ?type: (String | Symbol)?, **untyped opts) -> void
    def add_sparkline(location: nil, range: nil, type: nil, **opts)
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
    # @example Merge with string range
    #   sheet.merge("A1:C1")
    #
    # @example Merge with row and column bounds
    #   sheet.merge(row_start: 0, row_end: 2, col_start: "A", col_end: "C")
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
      canonical = DslHelpers.normalize_merge_range(
        range,
        row: row, col_start: col_start, col_end: col_end,
        row_start: row_start, row_end: row_end,
        strict_excel_mode: @strict_excel_mode
      )
      @merge_cells_ranges << canonical unless @merge_cells_ranges.include?(canonical)
    end

    # Freezes window panes at the specified row and column.
    #
    # @param row [Integer] Number of rows to freeze from top (0-based split position).
    # @param col [Integer, String] Number of columns to freeze from left (0-based or letter).
    # @return [void]
    # @api public
    #: (?row: Integer, ?col: (Integer | String)) -> void
    def freeze_pane(row: 0, col: 0)
      col = Elements::Cell.column_index(col)
      @freeze_pane = { row: row, col: col }
    end

    # Splits window panes without freezing.
    #
    # @param x_split [Integer] Horizontal split offset in points.
    # @param y_split [Integer] Vertical split offset in points.
    # @param top_left_cell [String, nil] Top-left cell reference in bottom-right pane.
    # @return [void]
    # @api public
    #: (?x_split: ::Integer, ?y_split: ::Integer, ?top_left_cell: String?) -> void
    def split_pane(x_split: 0, y_split: 0, top_left_cell: nil)
      @split_pane = { x_split: x_split, y_split: y_split, top_left_cell: top_left_cell }
    end

    # Sets the active cell and selection area for the worksheet.
    #
    # @param active_cell [String] Active cell reference (e.g. "A1").
    # @param sqref [String, nil] Selected range reference (e.g. "A1:D10").
    # @param pane [String, Symbol, nil] Target pane (:topLeft, :topRight, :bottomLeft, :bottomRight).
    # @return [void]
    # @api public
    #: (String active_cell, ?sqref: String?, ?pane: (String | Symbol)?) -> void
    def select_cell(active_cell, sqref: nil, pane: nil)
      @selection = { active_cell: active_cell, sqref: sqref || active_cell }
      @selection[:pane] = pane if pane
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
      @page_margins = DslHelpers.normalize_page_margins(
        left: left, right: right, top: top, bottom: bottom, header: header, footer: footer
      )
    end

    # Sets page setup configuration (orientation, paper size, fit to page, scaling).
    #
    # @param opts [Hash] Page setup options (e.g. orientation: :landscape, paper_size: 9).
    # @return [void]
    # @api public
    #: (**untyped opts) -> void
    def page_setup(**opts)
      @page_setup.merge!(opts)
      return unless (opts.key?(:fit_to_width) || opts.key?(:fit_to_height)) && !@sheet_properties.key?(:fit_to_page)

      @sheet_properties[:fit_to_page] = true
    end

    # Configures whether page setup fits content to page width/height.
    # When enabled, optionally sets fit_to_width and/or fit_to_height on page_setup.
    #
    # @param enabled [Boolean, nil] Whether to fit to page.
    # @param width [Integer, nil] Pages wide to fit to.
    # @param height [Integer, nil] Pages tall to fit to.
    # @return [Boolean, void]
    # @api public
    #: (?bool? enabled, ?width: Integer?, ?height: Integer?) -> (bool | void)
    def fit_to_page(enabled = nil, width: nil, height: nil)
      if enabled.nil? && width.nil? && height.nil?
        return @sheet_properties[:fit_to_page] ? true : false
      end

      @sheet_properties[:fit_to_page] = (enabled.nil? || (enabled ? true : false))
      page_setup(fit_to_width: width) if width
      page_setup(fit_to_height: height) if height
    end
    alias fit_to_page= fit_to_page
    alias fit_to_page? fit_to_page

    # Configures headers and footers for printing.
    #
    # @param opts [Hash] Header and footer options (e.g. odd_header: "&CHeader Text").
    # @return [void]
    # @api public
    #: (**untyped opts) -> void
    def header_footer(**opts)
      @header_footer.merge!(opts)
    end

    # Sets a print option (e.g. grid_lines: true, headings: true).
    #
    # @param name [Symbol] Option name.
    # @param value [Object] Option value.
    # @return [void]
    # @api public
    #: (Symbol name, untyped value) -> void
    def print_options(name, value)
      @print_options[name] = value
    end

    # Adds a defined name (named range or formula) scoped to this sheet or workbook.
    # Bare cell references and coordinate keywords are automatically formatted with sheet quotes and absolute $ references.
    #
    # @param name [String] Defined name identifier.
    # @param value [Object, nil] Cell reference, formula, or constant.
    # @param hidden [Boolean] Whether the defined name is hidden from Excel's UI.
    # @param col [Integer, String, nil] Column index or letter for coordinate-based reference.
    # @param rows [Integer, Range, nil] Row or row range (1-based) for coordinate-based reference.
    # @param range [String, nil] Cell range reference (e.g. "A1:A10").
    # @param sheet [String, nil] Target sheet name (defaults to current sheet).
    # @return [void]
    # @api public
    #: (String name, ?untyped value, ?hidden: bool, ?col: untyped, ?rows: untyped, ?range: untyped, ?sheet: String?, ?local: bool) -> void
    def defined_name(name, value = nil, hidden: false, col: nil, rows: nil, range: nil, sheet: nil, local: false)
      ref_sheet = sheet || @name
      normalized_val = DslHelpers.normalize_defined_name_value(value, sheet: ref_sheet, col: col, rows: rows, range: range)
      entry = { name: name, value: normalized_val, hidden: hidden }
      local_sheet = local ? (@name || sheet) : sheet
      entry[:local_sheet_name] = local_sheet if local_sheet
      @defined_names << entry
    end

    # Sets or gets the print area for the worksheet.
    #
    # @example Set print area range
    #   sheet.print_area("A1:H50")
    #
    # @example Get print area range
    #   sheet.print_area #=> "A1:H50"
    #
    # @param range [String, nil] Cell range (e.g. "A1:H50").
    # @return [String, nil]
    # @api public
    #: (?String? range) -> String?
    def print_area(range = nil)
      return @print_area if range.nil?

      @print_area = range
    end

    # Sets or gets repeating print titles (rows and/or columns) for pagination.
    #
    # @example Set repeating header rows
    #   sheet.print_titles(rows: "1:2")
    #
    # @example Get repeating print titles
    #   sheet.print_titles #=> { rows: "1:2", cols: nil }
    #
    # @param rows [String, nil] Repeating row range (e.g. "1:2").
    # @param cols [String, nil] Repeating column range (e.g. "A:B").
    # @return [Hash, nil]
    # @api public
    #: (?rows: String?, ?cols: String?) -> Hash[Symbol, untyped]?
    def print_titles(rows: nil, cols: nil)
      return @print_titles if rows.nil? && cols.nil?

      @print_titles = { rows: rows, cols: cols }
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
      @print_titles = (@print_titles || {}).merge(rows: rows_str)
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
                   first_name = Utils.col_index_to_name(first)
                   last_name = if last.is_a?(Integer)
                                 Utils.col_index_to_name(last)
                               elsif last
                                 last.to_s
                               else
                                 first_name
                               end
                   "#{first_name}:#{last_name}"
                 end
      @print_titles = (@print_titles || {}).merge(cols: cols_str)
    end

    # Sets sheet-level protection with optional password hashing.
    #
    # @param opts [Hash] Protection options (e.g. password: "secret", select_locked_cells: true).
    # @return [void]
    # @api public
    #: (**untyped opts) -> void
    def protect_sheet(**opts)
      @sheet_protection = DslHelpers.normalize_protection_options(opts)
    end

    # Inserts an embedded image from raw binary data.
    #
    # @param file_data [String] Binary image data.
    # @param ext [String] File extension ("png", "jpeg", etc.).
    # @param from_col [Integer] Top-left starting column (0-based).
    # @param from_row [Integer] Top-left starting row (0-based).
    # @param to_col [Integer] Bottom-right ending column (0-based).
    # @param to_row [Integer] Bottom-right ending row (0-based).
    # @param opts [Hash] Additional image anchoring options.
    # @return [void]
    # @api public
    #: (String file_data, ?ext: ::String, ?from_col: ::Integer, ?from_row: ::Integer, ?to_col: ::Integer, ?to_row: ::Integer, **untyped opts) -> void
    def image(file_data, ext: "png", from_col: 0, from_row: 0, to_col: 5, to_row: 10, **opts)
      img = { file_data: file_data, ext: ext, from_col: from_col, from_row: from_row, to_col: to_col, to_row: to_row }
      img.merge!(opts)
      @images << img
    end

    # Configures the worksheet background image (<picture>).
    #
    # @example Set background image from binary data
    #   sheet.background_image(png_bytes, ext: "png")
    #
    # @example Set background image from file path
    #   sheet.background_image("path/to/background.png")
    #
    # @param file_data [String, Hash, Elements::Image, nil] Image binary data, file path, or image options hash.
    # @param ext [String, nil] Image extension ("png", "jpeg", "jpg", "gif", etc.).
    # @param opts [Hash] Additional image options.
    # @return [Hash, nil]
    # @api public
    #: (?untyped file_data, ?ext: ::String?, **untyped opts) -> Hash[Symbol, untyped]?
    def background_image(file_data = nil, ext: nil, **opts)
      return @background_image if file_data.nil? && ext.nil? && opts.empty?

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

      @background_image = {
        data: data&.b,
        file_data: data&.b,
        ext: final_ext,
        filename: "background.#{final_ext}",
        content_type: mime
      }.merge(opts).freeze
    end
    alias background_picture background_image
    alias picture background_image

    def background_image=(val)
      background_image(val)
    end
    alias background_picture= background_image=
    alias picture= background_image=

    # Inserts a drawing shape (e.g. rectangle, callout, arrow).
    #
    # @param preset [String] Shape preset name (e.g. "rect", "roundRect").
    # @param text [String, nil] Shape label text.
    # @param from_col [Integer] Top-left column.
    # @param from_row [Integer] Top-left row.
    # @param to_col [Integer] Bottom-right column.
    # @param to_row [Integer] Bottom-right row.
    # @param opts [Hash] Additional shape formatting options.
    # @return [void]
    # @api public
    #: (**untyped opts) -> void
    def shape(preset: "rect", text: nil, from_col: 0, from_row: 0, to_col: 5, to_row: 5, **opts)
      shape = { preset: preset, text: text, from_col: from_col, from_row: from_row, to_col: to_col, to_row: to_row }
      shape[:name] = opts.delete(:name) || "Shape #{@shapes.size + 1}"
      shape.merge!(opts)
      @shapes << shape
    end

    # Gets or sets the sheet visibility state (:visible, :hidden, or :very_hidden).
    #
    # @param new_state [Symbol, nil] The new state.
    # @return [Symbol]
    # @api public
    #: (?Symbol? new_state) -> Symbol
    def state(new_state = nil)
      if new_state
        val = new_state.to_sym
        raise ArgumentError, "Invalid sheet state ':#{val}'. Allowed states are :visible, :hidden, :very_hidden" unless VALID_SHEET_STATES.include?(val)

        @state = val
      else
        @state
      end
    end

    # Gets or sets the worksheet tab color.
    #
    # @overload tab_color()
    #   Returns the current tab color specification.
    #   @return [String, Hash, nil]
    # @overload tab_color(spec, **opts)
    #   Sets the tab color using an RGB hex string or color options.
    #   @param spec [String, Hash, Symbol, nil] RGB hex string (e.g. "FF0000"), theme/indexed hash, or nil.
    #   @param opts [Hash] Color keyword arguments (e.g. theme: 1, tint: 0.5, indexed: 9, auto: true).
    #   @return [String, Hash, nil]
    # @api public
    #: (?String | Symbol | Hash[Symbol, untyped] | nil spec, **untyped opts) -> (String | Hash[Symbol, untyped] | nil)
    def tab_color(spec = nil, **opts)
      if spec.nil? && opts.empty?
        @sheet_properties[:tab_color]
      else
        color = if opts.empty?
                  spec
                else
                  spec.is_a?(Hash) ? spec.merge(opts) : opts
                end
        @sheet_properties[:tab_color] = color
      end
    end
    alias tab_color= tab_color

    # Sets a sheet-level property (e.g. tab_color: "FF0000").
    #
    # @param name [Symbol] Property name.
    # @param value [Object] Property value.
    # @return [void]
    # @api public
    #: (Symbol name, untyped value) -> void
    def sheet_properties(name, value)
      sym = name.to_sym
      if sym == :state
        val = value.to_sym
        raise ArgumentError, "Invalid sheet state ':#{val}'. Allowed states are :visible, :hidden, :very_hidden" unless VALID_SHEET_STATES.include?(val)

        @state = val
        return
      elsif sym == :hidden
        @state = value ? :hidden : :visible
        return
      end

      @sheet_properties[name] = value
    end

    # Sets a sheet view display property (e.g. show_grid_lines: true, zoom_scale: 120).
    #
    # @param name [Symbol] View property name.
    # @param value [Object] View property value.
    # @return [void]
    # @api public
    #: (Symbol name, untyped value) -> void
    def sheet_view(name, value)
      @sheet_view[name] = value
    end

    # Configures sheet format properties (e.g. default_row_height, default_col_width).
    #
    # @example Using keyword arguments
    #   sheet.sheet_format(default_row_height: 24, default_col_width: 15)
    #
    # @example Using property name and value
    #   sheet.sheet_format(:default_row_height, 24)
    #
    # @param name [Symbol, String, Hash, nil] Property name or hash of properties.
    # @param value [Object, nil] Property value when name is provided.
    # @param opts [Hash] Keyword arguments for sheet format properties.
    # @return [Hash]
    # @api public
    #: (?Symbol | String | Hash[Symbol, untyped] | nil name, ?untyped value, **untyped opts) -> Hash[Symbol, untyped]
    def sheet_format(name = nil, value = nil, **opts)
      if name.is_a?(Hash)
        @sheet_format.merge!(name)
      elsif name && value
        @sheet_format[name.to_sym] = value
      end
      @sheet_format.merge!(opts) unless opts.empty?
      @sheet_format
    end

    # Inserts a horizontal page break before the specified row.
    #
    # @param row_num [Integer] 1-based row number.
    # @return [void]
    # @api public
    #: (Integer row_num) -> void
    def page_break_row(row_num)
      @row_breaks << row_num
    end

    # Inserts a vertical page break before the specified column.
    #
    # @param col_index [Integer, String] Column index (0-based) or letter ("B").
    # @return [void]
    # @api public
    #: (Integer | String col_index) -> void
    def page_break_col(col_index)
      col_index = Elements::Cell.column_index(col_index)
      @col_breaks << col_index
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
    alias add_protected_range protected_range

    # Adds a cell watch to the worksheet.
    #
    # @param ref [String] Cell reference (e.g. "A1").
    # @return [void]
    # @api public
    #: (String ref) -> void
    def cell_watch(ref)
      @cell_watches << ref.to_s
    end
    alias add_cell_watch cell_watch

    # Adds an ignored error entry for the worksheet.
    #
    # @param sqref [String] Cell range reference (e.g. "A1" or "A1:B10").
    # @param opts [Hash] Error check flags to ignore (e.g. number_stored_as_text: true, eval_error: true).
    # @return [void]
    # @api public
    #: (String sqref, ?eval_error: bool, ?two_digit_text_year: bool, ?number_stored_as_text: bool, ?formula: bool, ?formula_range: bool, ?unlocked_formula: bool, ?empty_cell_reference: bool, ?list_data_validation: bool, ?calculated_column: bool) -> void
    def ignore_error(sqref, **opts)
      @ignored_errors << { sqref: sqref.to_s }.merge(opts)
    end
    alias add_ignored_error ignore_error

    # Returns ignored error rules configured for this worksheet.
    #
    # @return [Array<Hash[Symbol, untyped]>]
    # @api public
    #: () -> Array[Hash[Symbol, untyped]]
    def ignored_errors
      @ignored_errors.dup
    end

    # Gets or sets phonetic (furigana) guide properties for the worksheet.
    #
    # @overload phonetic_properties()
    #   Returns the current phonetic properties.
    #   @return [Hash, nil]
    # @overload phonetic_properties(props = nil, font_id: nil, type: nil, alignment: nil)
    #   Sets phonetic properties.
    #   @param spec [Hash, Integer, nil]
    #   @param font_id [Integer, nil]
    #   @param type [Symbol, String, nil]
    #   @param alignment [Symbol, String, nil]
    #   @return [Hash, nil]
    # @api public
    #: (?Hash[Symbol, untyped] | Integer | nil spec, ?font_id: Integer?, ?type: (Symbol | String)?, ?alignment: (Symbol | String)?, **untyped opts) -> Hash[Symbol, untyped]?
    def phonetic_properties(spec = nil, font_id: nil, type: nil, alignment: nil, **opts)
      if spec.nil? && font_id.nil? && type.nil? && alignment.nil? && opts.empty?
        @phonetic_properties
      else
        props = if spec.is_a?(Hash)
                  spec.merge(opts)
                elsif spec.is_a?(Integer)
                  { font_id: spec, type: type, alignment: alignment }.merge(opts)
                else
                  { font_id: font_id, type: type, alignment: alignment }.merge(opts)
                end
        props.compact!
        @phonetic_properties = WorksheetBuilder.normalize_phonetic_properties(props)
      end
    end
    alias phonetic_properties= phonetic_properties
    alias phonetic_pr phonetic_properties
    alias phonetic_pr= phonetic_properties

    # Normalizes phonetic properties to OOXML schema values.
    #
    # @param props [Hash, nil]
    # @return [Hash[Symbol, untyped], nil]
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

    # Sets or gets the internal code name of the worksheet.
    #
    # @param val [String, Symbol, nil] The code name.
    # @return [String, nil]
    # @api public
    #: (?String | Symbol | nil val) -> String?
    def code_name(val = nil)
      return @sheet_properties[:code_name] if val.nil?

      @sheet_properties[:code_name] = val.to_s
    end
    alias code_name= code_name

    # Configures worksheet calculation properties (<sheetCalcPr>).
    #
    # @example Instruct Excel to recalculate formulas when loading sheet
    #   sheet.sheet_calc_pr(full_calc_on_load: true)
    #
    # @param spec [Hash, nil] Calculation properties hash.
    # @param full_calc_on_load [Boolean, nil] Whether to recalculate when loading.
    # @return [Hash, nil]
    # @api public
    #: (?Hash[Symbol, untyped]? spec, ?full_calc_on_load: bool?, **untyped opts) -> Hash[Symbol, untyped]?
    def sheet_calc_pr(spec = nil, full_calc_on_load: nil, **opts)
      if spec.nil? && full_calc_on_load.nil? && opts.empty?
        return @sheet_calc_pr || (@sheet_properties[:full_calc_on_load] ? { full_calc_on_load: true } : nil)
      end

      pr = spec.is_a?(Hash) ? spec.dup : {}
      pr[:full_calc_on_load] = full_calc_on_load unless full_calc_on_load.nil?
      pr.merge!(opts) unless opts.empty?
      @sheet_properties[:full_calc_on_load] = pr[:full_calc_on_load] unless pr[:full_calc_on_load].nil?
      @sheet_calc_pr = pr
    end
    alias sheet_calc_pr= sheet_calc_pr
    alias calc_pr sheet_calc_pr
    alias calc_pr= sheet_calc_pr

    # Builds and returns the in-memory {Elements::Worksheet}.
    #
    # @return [Elements::Worksheet]
    # @api public
    #: () -> Elements::Worksheet
    def build
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
      facade_meta[:col_breaks] = @col_breaks unless @col_breaks.empty?
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
        ref = h[:cell].to_s.upcase
        entry = {}
        entry[:url] = h[:url] if h[:url]
        entry[:display] = h[:display] if h[:display]
        entry[:tooltip] = h[:tooltip] if h[:tooltip]
        entry[:location] = h[:location] if h[:location]
        hl_hash[ref] = entry unless entry.empty?
      end

      cm_hash = {}
      @comments.each do |c|
        ref = (c[:ref] || c[:cell]).to_s.upcase
        cm_hash[ref] = c
      end

      rows = if hl_hash.empty? && cm_hash.empty?
               @rows
             else
               @rows.map do |row|
                 new_cells = row.cells.map do |cell|
                   hl = hl_hash[cell.ref]
                   cm = cm_hash[cell.ref]
                   if hl || cm
                     cell.with(hyperlink: hl || cell.hyperlink, comment: cm || cell.comment)
                   else
                     cell
                   end
                 end
                 row.with(cells: new_cells)
               end
             end

      Elements::Worksheet.new(
        name: @name, rows: rows, columns: @columns, charts: @charts,
        conditional_formatting: @conditional_formats,
        data_validations: @data_validations,
        unmapped_data: facade_meta.empty? ? {} : { facade: facade_meta },
        state: @state,
        hyperlinks: hl_hash,
        comments: @comments,
        print_area: @print_area,
        print_titles: @print_titles,
        sparkline_groups: @sparkline_groups,
        merged_cells: @merge_cells_ranges,
        auto_filter: @auto_filter,
        tables: @tables,
        shapes: @shapes,
        freeze_pane: @freeze_pane,
        page_setup: @page_setup,
        page_margins: @page_margins,
        header_footer: @header_footer,
        print_options: @print_options,
        sheet_protection: @sheet_protection,
        sheet_view: @sheet_view,
        sheet_properties: @sheet_properties,
        sheet_format: @sheet_format,
        row_breaks: @row_breaks,
        col_breaks: @col_breaks,
        protected_ranges: @protected_ranges,
        cell_watches: @cell_watches,
        ignored_errors: @ignored_errors,
        phonetic_properties: @phonetic_properties,
        selection: @selection,
        sheet_calc_pr: @sheet_calc_pr,
        background_image: @background_image
      )
    end

    # Internal: returns styles for later processing by WorkbookBuilder
    # @return [Hash{String => StyleBuilder}]
    #: Hash[String, StyleBuilder]
    attr_reader :styles
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

    #: (untyped key) -> untyped
    def [](key)
      @options[key.to_sym]
    end

    #: (untyped key, untyped value) -> untyped
    def []=(key, value)
      @options[key.to_sym] = value
    end

    #: (?String? val) -> String?
    def ref(val = nil)
      if val
        @options[:ref] = val.to_s
      else
        @options[:ref]
      end
    end

    #: (String val) -> String
    def ref=(val)
      @options[:ref] = val.to_s
    end

    #: (*untyped cols) -> untyped
    def columns(*cols)
      if cols.empty?
        @options[:columns]
      elsif cols.size == 1 && cols.first.is_a?(Array)
        @options[:columns] = cols.first
      else
        @options[:columns] = cols
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
    end
    alias add_column column

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

    #: (?String? val) -> String?
    def name(val = nil)
      if val
        @options[:name] = val.to_s
      else
        @options[:name]
      end
    end

    #: (String val) -> String
    def name=(val)
      @options[:name] = val.to_s
    end

    #: (?String? val) -> String?
    def display_name(val = nil)
      if val
        @options[:display_name] = val.to_s
      else
        @options[:display_name]
      end
    end

    #: (String val) -> String
    def display_name=(val)
      @options[:display_name] = val.to_s
    end

    #: (?String? val) -> String?
    def style(val = nil)
      if val
        @options[:style] = val.to_s
      else
        @options[:style]
      end
    end

    #: (String val) -> String
    def style=(val)
      @options[:style] = val.to_s
    end

    #: (?bool? val) -> bool?
    def show_row_stripes(val = nil)
      if val.nil?
        @options[:show_row_stripes]
      else
        @options[:show_row_stripes] = val ? true : false
      end
    end

    #: (bool val) -> bool
    def show_row_stripes=(val)
      @options[:show_row_stripes] = val ? true : false
    end

    #: (?bool? val) -> bool?
    def show_column_stripes(val = nil)
      if val.nil?
        @options[:show_column_stripes]
      else
        @options[:show_column_stripes] = val ? true : false
      end
    end

    #: (bool val) -> bool
    def show_column_stripes=(val)
      @options[:show_column_stripes] = val ? true : false
    end

    #: (?bool? val) -> bool?
    def show_first_column(val = nil)
      if val.nil?
        @options[:show_first_column]
      else
        @options[:show_first_column] = val ? true : false
      end
    end

    #: (bool val) -> bool
    def show_first_column=(val)
      @options[:show_first_column] = val ? true : false
    end

    #: (?bool? val) -> bool?
    def show_last_column(val = nil)
      if val.nil?
        @options[:show_last_column]
      else
        @options[:show_last_column] = val ? true : false
      end
    end

    #: (bool val) -> bool
    def show_last_column=(val)
      @options[:show_last_column] = val ? true : false
    end

    #: (?Integer? val) -> Integer?
    def total_row_count(val = nil)
      if val.nil?
        @options[:totals_row_count] || @options[:total_row_count]
      else
        @options[:totals_row_count] = val.to_i
      end
    end
    alias totals_row_count total_row_count

    #: (Integer val) -> Integer
    def total_row_count=(val)
      @options[:totals_row_count] = val.to_i
    end
    alias totals_row_count= total_row_count=

    #: (?bool? val) -> bool?
    def total_row(val = nil)
      if val.nil?
        rc = @options[:totals_row_count] || @options[:total_row_count]
        rc ? rc.positive? : @options[:total_row]
      else
        @options[:totals_row_count] = val ? 1 : 0
        val ? true : false
      end
    end
    alias total_row= total_row
    alias totals_row total_row
    alias totals_row= total_row

    #: (?bool? val) -> bool?
    def auto_filter(val = nil)
      if val.nil?
        @options[:auto_filter]
      else
        @options[:auto_filter] = val ? true : false
      end
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

    #: (untyped key) -> untyped
    def [](key)
      @options[key.to_sym]
    end

    #: (untyped key, untyped value) -> untyped
    def []=(key, value)
      @options[key.to_sym] = value
    end

    #: (?String? val) -> String?
    def sqref(val = nil)
      if val
        @options[:sqref] = val.to_s
      else
        @options[:sqref]
      end
    end

    #: (String val) -> String
    def sqref=(val)
      @options[:sqref] = val.to_s
    end
    alias range sqref
    alias range= sqref=

    #: (?untyped val) -> untyped
    def columns(val = nil)
      if val
        @options[:columns] = val
      else
        @options[:columns]
      end
    end

    #: (untyped val) -> untyped
    def columns=(val)
      @options[:columns] = val
    end
    alias col columns
    alias col= columns=

    #: (?untyped val) -> untyped
    def rows(val = nil)
      if val
        @options[:rows] = val
      else
        @options[:rows]
      end
    end

    #: (untyped val) -> untyped
    def rows=(val)
      @options[:rows] = val
    end
    alias row rows
    alias row= rows=

    #: (?untyped val) -> untyped
    def except_rows(val = nil)
      if val
        @options[:except_rows] = val
      else
        @options[:except_rows]
      end
    end

    #: (untyped val) -> untyped
    def except_rows=(val)
      @options[:except_rows] = val
    end

    #: (?Array[untyped]? val) -> untyped
    def in(val = nil)
      if val
        @options[:in] = val
      else
        @options[:in]
      end
    end

    #: (Array[untyped] val) -> untyped
    def in=(val)
      @options[:in] = val
    end

    #: (?untyped val) -> untyped
    def type(val = nil)
      if val
        @options[:type] = val
      else
        @options[:type]
      end
    end

    #: (untyped val) -> untyped
    def type=(val)
      @options[:type] = val
    end

    #: (?untyped val) -> untyped
    def operator(val = nil)
      if val
        @options[:operator] = val
      else
        @options[:operator]
      end
    end

    #: (untyped val) -> untyped
    def operator=(val)
      @options[:operator] = val
    end

    #: (?untyped val) -> untyped
    def formula1(val = nil)
      if val
        @options[:formula1] = val
      else
        @options[:formula1]
      end
    end

    #: (untyped val) -> untyped
    def formula1=(val)
      @options[:formula1] = val
    end

    #: (?untyped val) -> untyped
    def formula2(val = nil)
      if val
        @options[:formula2] = val
      else
        @options[:formula2]
      end
    end

    #: (untyped val) -> untyped
    def formula2=(val)
      @options[:formula2] = val
    end

    #: (?bool? val) -> bool?
    def allow_blank(val = nil)
      if val.nil?
        @options[:allow_blank]
      else
        @options[:allow_blank] = val ? true : false
      end
    end

    #: (bool val) -> bool
    def allow_blank=(val)
      @options[:allow_blank] = val ? true : false
    end

    #: (?bool? val) -> bool?
    def show_input_message(val = nil)
      if val.nil?
        @options[:show_input_message]
      else
        @options[:show_input_message] = val ? true : false
      end
    end

    #: (bool val) -> bool
    def show_input_message=(val)
      @options[:show_input_message] = val ? true : false
    end

    #: (?bool? val) -> bool?
    def show_error_message(val = nil)
      if val.nil?
        @options[:show_error_message]
      else
        @options[:show_error_message] = val ? true : false
      end
    end

    #: (bool val) -> bool
    def show_error_message=(val)
      @options[:show_error_message] = val ? true : false
    end

    #: (?String? val) -> String?
    def error_title(val = nil)
      if val
        @options[:error_title] = val.to_s
      else
        @options[:error_title]
      end
    end

    #: (String val) -> String
    def error_title=(val)
      @options[:error_title] = val.to_s
    end

    #: (?String? val) -> String?
    def error_message(val = nil)
      if val
        @options[:error] = val.to_s
      else
        @options[:error]
      end
    end

    #: (String val) -> String
    def error_message=(val)
      @options[:error] = val.to_s
    end
    alias error error_message
    alias error= error_message=

    #: (?String? val) -> String?
    def prompt_title(val = nil)
      if val
        @options[:prompt_title] = val.to_s
      else
        @options[:prompt_title]
      end
    end

    #: (String val) -> String
    def prompt_title=(val)
      @options[:prompt_title] = val.to_s
    end

    #: (?String? val) -> String?
    def prompt_message(val = nil)
      if val
        @options[:prompt] = val.to_s
      else
        @options[:prompt]
      end
    end

    #: (String val) -> String
    def prompt_message=(val)
      @options[:prompt] = val.to_s
    end
    alias prompt prompt_message
    alias prompt= prompt_message=

    #: (?untyped val) -> untyped
    def error_style(val = nil)
      if val
        @options[:error_style] = val
      else
        @options[:error_style]
      end
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

    #: (untyped key) -> untyped
    def [](key)
      @options[key.to_sym]
    end

    #: (untyped key, untyped value) -> untyped
    def []=(key, value)
      @options[key.to_sym] = value
    end

    #: (?String? val) -> String?
    def sqref(val = nil)
      if val
        @options[:sqref] = val.to_s
      else
        @options[:sqref]
      end
    end

    #: (String val) -> String
    def sqref=(val)
      @options[:sqref] = val.to_s
    end
    alias range sqref
    alias range= sqref=

    #: (?untyped val) -> untyped
    def columns(val = nil)
      if val
        @options[:columns] = val
      else
        @options[:columns]
      end
    end

    #: (untyped val) -> untyped
    def columns=(val)
      @options[:columns] = val
    end
    alias col columns
    alias col= columns=

    #: (?untyped val) -> untyped
    def rows(val = nil)
      if val
        @options[:rows] = val
      else
        @options[:rows]
      end
    end

    #: (untyped val) -> untyped
    def rows=(val)
      @options[:rows] = val
    end
    alias row rows
    alias row= rows=

    #: (?untyped val) -> untyped
    def except_rows(val = nil)
      if val
        @options[:except_rows] = val
      else
        @options[:except_rows]
      end
    end

    #: (untyped val) -> untyped
    def except_rows=(val)
      @options[:except_rows] = val
    end

    #: (?untyped val) -> untyped
    def type(val = nil)
      if val
        @options[:type] = val
      else
        @options[:type]
      end
    end

    #: (untyped val) -> untyped
    def type=(val)
      @options[:type] = val
    end

    #: (?untyped val) -> untyped
    def operator(val = nil)
      if val
        @options[:operator] = val
      else
        @options[:operator]
      end
    end

    #: (untyped val) -> untyped
    def operator=(val)
      @options[:operator] = val
    end

    #: (?untyped val) -> untyped
    def formula(val = nil)
      if val
        @options[:formula] = val
      else
        @options[:formula]
      end
    end

    #: (untyped val) -> untyped
    def formula=(val)
      @options[:formula] = val
    end

    #: (*untyped vals) -> untyped
    def formulas(*vals)
      if vals.empty?
        @options[:formulas]
      elsif vals.size == 1 && vals.first.is_a?(Array)
        @options[:formulas] = vals.first
      else
        @options[:formulas] = vals
      end
    end

    #: (untyped vals) -> untyped
    def formulas=(vals)
      @options[:formulas] = vals.is_a?(Array) ? vals : [vals]
    end

    #: (?untyped val) -> untyped
    def font(val = nil)
      if val
        @options[:font] = val
      else
        @options[:font]
      end
    end

    #: (untyped val) -> untyped
    def font=(val)
      @options[:font] = val
    end

    #: (?untyped val) -> untyped
    def fill(val = nil)
      if val
        @options[:fill] = val
      else
        @options[:fill]
      end
    end

    #: (untyped val) -> untyped
    def fill=(val)
      @options[:fill] = val
    end

    #: (?untyped val) -> untyped
    def fill_color(val = nil)
      if val
        @options[:fill_color] = val
      else
        @options[:fill_color]
      end
    end

    #: (untyped val) -> untyped
    def fill_color=(val)
      @options[:fill_color] = val
    end

    #: (?untyped val) -> untyped
    def font_color(val = nil)
      if val
        @options[:font_color] = val
      else
        @options[:font_color]
      end
    end

    #: (untyped val) -> untyped
    def font_color=(val)
      @options[:font_color] = val
    end

    #: (?bool? val) -> bool?
    def bold(val = nil)
      if val.nil?
        @options[:bold]
      else
        @options[:bold] = val ? true : false
      end
    end

    #: (bool val) -> bool
    def bold=(val)
      @options[:bold] = val ? true : false
    end

    #: (?bool? val) -> bool?
    def italic(val = nil)
      if val.nil?
        @options[:italic]
      else
        @options[:italic] = val ? true : false
      end
    end

    #: (bool val) -> bool
    def italic=(val)
      @options[:italic] = val ? true : false
    end

    #: (?untyped val) -> untyped
    def underline(val = nil)
      if val.nil?
        @options[:underline]
      else
        @options[:underline] = val
      end
    end

    #: (untyped val) -> untyped
    def underline=(val)
      @options[:underline] = val
    end

    #: (?untyped val) -> untyped
    def size(val = nil)
      if val
        @options[:size] = val
      else
        @options[:size]
      end
    end

    #: (untyped val) -> untyped
    def size=(val)
      @options[:size] = val
    end
    alias sz size
    alias sz= size=

    #: (?bool? val) -> bool?
    def stop_if_true(val = nil)
      if val.nil?
        @options[:stop_if_true]
      else
        @options[:stop_if_true] = val ? true : false
      end
    end

    #: (bool val) -> bool
    def stop_if_true=(val)
      @options[:stop_if_true] = val ? true : false
    end

    #: (?Integer? val) -> Integer?
    def priority(val = nil)
      if val
        @options[:priority] = val.to_i
      else
        @options[:priority]
      end
    end

    #: (Integer val) -> Integer
    def priority=(val)
      @options[:priority] = val.to_i
    end

    #: (?bool? val) -> bool?
    def above_average(val = nil)
      if val.nil?
        @options[:above_average]
      else
        @options[:above_average] = val ? true : false
      end
    end

    #: (bool val) -> bool
    def above_average=(val)
      @options[:above_average] = val ? true : false
    end

    #: (?bool? val) -> bool?
    def equal_average(val = nil)
      if val.nil?
        @options[:equal_average]
      else
        @options[:equal_average] = val ? true : false
      end
    end

    #: (bool val) -> bool
    def equal_average=(val)
      @options[:equal_average] = val ? true : false
    end

    #: (?Integer? val) -> Integer?
    def std_dev(val = nil)
      if val
        @options[:std_dev] = val.to_i
      else
        @options[:std_dev]
      end
    end

    #: (Integer val) -> Integer
    def std_dev=(val)
      @options[:std_dev] = val.to_i
    end

    #: (?Integer? val) -> Integer?
    def rank(val = nil)
      if val
        @options[:rank] = val.to_i
      else
        @options[:rank]
      end
    end

    #: (Integer val) -> Integer
    def rank=(val)
      @options[:rank] = val.to_i
    end

    #: (?bool? val) -> bool?
    def percent(val = nil)
      if val.nil?
        @options[:percent]
      else
        @options[:percent] = val ? true : false
      end
    end

    #: (bool val) -> bool
    def percent=(val)
      @options[:percent] = val ? true : false
    end

    #: (?bool? val) -> bool?
    def bottom(val = nil)
      if val.nil?
        @options[:bottom]
      else
        @options[:bottom] = val ? true : false
      end
    end

    #: (bool val) -> bool
    def bottom=(val)
      @options[:bottom] = val ? true : false
    end

    #: (?String? val) -> String?
    def text(val = nil)
      if val
        @options[:text] = val.to_s
      else
        @options[:text]
      end
    end

    #: (String val) -> String
    def text=(val)
      @options[:text] = val.to_s
    end

    #: (?untyped val) -> untyped
    def time_period(val = nil)
      if val
        @options[:time_period] = val
      else
        @options[:time_period]
      end
    end

    #: (untyped val) -> untyped
    def time_period=(val)
      @options[:time_period] = val
    end

    #: (?untyped val) -> untyped
    def color_scale(val = nil)
      if val
        @options[:color_scale] = val
      else
        @options[:color_scale]
      end
    end

    #: (untyped val) -> untyped
    def color_scale=(val)
      @options[:color_scale] = val
    end

    #: (?untyped val) -> untyped
    def data_bar(val = nil)
      if val
        @options[:data_bar] = val
      else
        @options[:data_bar]
      end
    end

    #: (untyped val) -> untyped
    def data_bar=(val)
      @options[:data_bar] = val
    end

    #: (?untyped val) -> untyped
    def icon_set(val = nil)
      if val
        @options[:icon_set] = val
      else
        @options[:icon_set]
      end
    end

    #: (untyped val) -> untyped
    def icon_set=(val)
      @options[:icon_set] = val
    end

    #: (?untyped val) -> untyped
    def dxf(val = nil)
      if val
        @options[:dxf] = val
      else
        @options[:dxf]
      end
    end

    #: (untyped val) -> untyped
    def dxf=(val)
      @options[:dxf] = val
    end

    #: (?Integer? val) -> Integer?
    def format_id(val = nil)
      if val
        @options[:format_id] = val.to_i
      else
        @options[:format_id]
      end
    end

    #: (Integer val) -> Integer
    def format_id=(val)
      @options[:format_id] = val.to_i
    end

    #: () -> Hash[Symbol, untyped]
    def to_h
      @options.dup
    end
  end
end
