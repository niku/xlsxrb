# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  # DSL context for building a single in-memory worksheet in {Xlsxrb.build}.
  #
  # @api public
  class WorksheetBuilder
    # @param name [String] The worksheet name.
    # @param strict_excel_mode [Boolean] Whether to enforce Microsoft Excel limits.
    #: (String name, ?strict_excel_mode: bool) -> void
    def initialize(name, strict_excel_mode: true)
      @name = name
      @strict_excel_mode = strict_excel_mode
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
      @row_breaks = []
      @col_breaks = []
    end

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
    # @return [void]
    # @raise [ArgumentError] If limits are exceeded when strict_excel_mode is enabled.
    # @api public
    #: (Array[untyped] | Hash[untyped, untyped] values, ?styles: untyped, ?height: Float | Integer | nil, ?hidden: bool, ?custom_height: bool, ?outline_level: Integer | nil) -> void
    def row(values, styles: nil, height: nil, hidden: false, custom_height: false, outline_level: nil)
      row_index = @rows.size
      # See: https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
      if @strict_excel_mode
        raise ArgumentError, "Row index #{row_index} exceeds Excel limit of 1,048,576 rows" if row_index >= 1_048_576
        raise ArgumentError, "Row height #{height} must be between 0 and 409 points (Excel limitation)" if height && (height.negative? || height > 409)
      end

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

      max_len = values.size
      max_len = [max_len, styles.size].max if styles.is_a?(Array)
      # See: https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
      raise ArgumentError, "Row contains #{max_len} columns, exceeding Excel limit of 16_384 columns" if @strict_excel_mode && max_len > 16_384

      cells = Array.new(max_len)
      style_lookup = styles.is_a?(Array)

      col_index = 0
      while col_index < max_len
        val = col_index < values.size ? values[col_index] : nil
        raise ArgumentError, "Invalid cell value type or value: #{val.class} for value #{val.inspect}" unless val.nil? || val.is_a?(String) || (val.is_a?(Numeric) && !(val.is_a?(Float) && (val.infinite? || val.nan?))) || val.is_a?(TrueClass) || val.is_a?(FalseClass) || val.is_a?(Date) || val.is_a?(Time) || val.is_a?(Elements::Formula) || (val.is_a?(Hash) && val.key?(:formula)) || val.is_a?(Elements::RichText) || (val.is_a?(Array) && val.first.is_a?(Hash) && (val.first.key?(:text) || val.first.key?("text")))
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
        # If value is a Formula object or Hash with :formula, store it as the cell's formula
        cells[col_index] = if val.is_a?(Elements::Formula)
                             Elements::Cell.new(
                               row_index: row_index,
                               column_index: col_index,
                               value: nil,
                               formula: val,
                               style_index: style_name
                             )
                           elsif val.is_a?(Hash) && val.key?(:formula)
                             f_obj = Elements::Formula.new(val[:formula])
                             Elements::Cell.new(
                               row_index: row_index,
                               column_index: col_index,
                               value: val[:value],
                               formula: f_obj,
                               style_index: style_name
                             )
                           else
                             Elements::Cell.new(
                               row_index: row_index,
                               column_index: col_index,
                               value: val,
                               style_index: style_name
                             )
                           end
        col_index += 1
      end

      cells.compact!
      @rows << Elements::Row.new(
        index: row_index,
        cells: cells,
        height: height,
        hidden: hidden,
        custom_height: custom_height || !height.nil?,
        outline_level: outline_level
      )
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
    # @param hidden [Boolean] Whether the column is hidden.
    # @param custom_width [Boolean] Whether custom width is explicitly set.
    # @param outline_level [Integer, nil] Grouping/outline level.
    # @return [void]
    # @raise [ArgumentError] If width exceeds 255 in strict mode.
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

      indices.each do |idx|
        @columns << Elements::Column.new(
          index: idx,
          width: width,
          hidden: hidden,
          custom_width: custom_width || !width.nil?,
          outline_level: outline_level
        )
      end
    end

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
    #: (String | Integer cell, ?String? url, ?display: String?, ?tooltip: String?, ?location: String?) -> void
    def hyperlink(cell, url = nil, display: nil, tooltip: nil, location: nil)
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

    # Adds a data validation rule to a cell or range.
    #
    # @param sqref [String] Target cell or range reference (e.g. "B2:B100").
    # @param opts [Hash] Validation configuration (e.g. type: :list, formula1: '"Option1,Option2"').
    # @return [void]
    # @api public
    #: (String sqref, **untyped opts) -> void
    def validate_data(sqref, **opts)
      @data_validations << opts.merge(sqref: sqref)
    end

    # Adds a conditional formatting rule to a cell range.
    #
    # @param sqref [String] Target cell range reference (e.g. "C2:C50").
    # @param opts [Hash] Rule options (e.g. type: :cellIs, operator: :greaterThan, formula: "100").
    # @return [void]
    # @api public
    #: (String sqref, **untyped opts) -> void
    def conditional_format(sqref, **opts)
      @conditional_formats << opts.merge(sqref: sqref)
    end

    # Adds an Excel Table (ListObject) to the sheet.
    #
    # @param ref [String] The table range (e.g. "A1:D20").
    # @param columns [Array<String>] Column header names.
    # @param name [String, nil] Table name identifier.
    # @param display_name [String, nil] Table display name.
    # @param style [String, nil] Built-in table style name (e.g. "TableStyleMedium2").
    # @param opts [Hash] Additional table properties.
    # @return [void]
    # @api public
    #: (String ref, columns: Array[String | Hash[Symbol, untyped]], ?name: String?, ?display_name: String?, ?style: String?, **untyped opts) -> void
    def table(ref, columns:, name: nil, display_name: nil, style: nil, **opts)
      tbl = { ref: ref, columns: columns }
      tbl[:name] = name if name
      tbl[:display_name] = display_name if display_name
      tbl[:style] = style if style
      tbl.merge!(opts)
      @tables << tbl
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
    # @param type [String, nil] "line" (default), "column", or "stacked".
    # @param opts [Hash] Color and formatting options.
    # @return [void]
    # @api public
    #: (sparklines: Array[String | Hash[Symbol, untyped]], ?type: String?, **untyped opts) -> void
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
        return if @merge_cells_ranges.include?(range)

        @merge_cells_ranges << range
      else
        r_start = row || row_start || 0
        r_end = row || row_end || 0
        c_start = Elements::Cell.column_index(col_start || 0)
        c_end = Elements::Cell.column_index(col_end || 0)
        start_ref = "#{Xlsxrb::Elements::Cell.column_letter(c_start)}#{r_start + 1}"
        end_ref = "#{Xlsxrb::Elements::Cell.column_letter(c_end)}#{r_end + 1}"
        @merge_cells_ranges << "#{start_ref}:#{end_ref}"
      end
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
      @page_margins = { left: left, right: right, top: top, bottom: bottom, header: header, footer: footer }.compact
    end

    # Sets page setup configuration (orientation, paper size, fit to page, scaling).
    #
    # @param opts [Hash] Page setup options (e.g. orientation: :landscape, paper_size: 9).
    # @return [void]
    # @api public
    #: (**untyped opts) -> void
    def page_setup(**opts)
      @page_setup.merge!(opts)
    end

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

    # Sets sheet-level protection with optional password hashing.
    #
    # @param opts [Hash] Protection options (e.g. password: "secret", select_locked_cells: true).
    # @return [void]
    # @api public
    #: (**untyped opts) -> void
    def protect_sheet(**opts)
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
      @sheet_protection = normalized
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

    # Sets a sheet-level property (e.g. tab_color: "FF0000").
    #
    # @param name [Symbol] Property name.
    # @param value [Object] Property value.
    # @return [void]
    # @api public
    #: (Symbol name, untyped value) -> void
    def sheet_properties(name, value)
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
      facade_meta[:sheet_view] = @sheet_view unless @sheet_view.empty?
      facade_meta[:row_breaks] = @row_breaks unless @row_breaks.empty?
      facade_meta[:col_breaks] = @col_breaks unless @col_breaks.empty?

      Elements::Worksheet.new(
        name: @name, rows: @rows, columns: @columns, charts: @charts,
        unmapped_data: facade_meta.empty? ? {} : { facade: facade_meta }
      )
    end

    # Internal: returns styles for later processing by WorkbookBuilder
    # @return [Hash{String => StyleBuilder}]
    #: Hash[String, StyleBuilder]
    attr_reader :styles
  end
end
