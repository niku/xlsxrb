# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  # Streaming row implementation that parses cells on-demand / lazily.
  # Provides O(1) memory consumption even for rows with tens of thousands of columns.
  #
  # @example Streaming cells one-by-one (O(1) memory)
  #   row.each_cell do |cell|
  #     puts "#{cell.ref}: #{cell.value}"
  #   end
  #
  # @example Random access or array conversion (cached on-demand)
  #   cell = row[0]
  #   values = row.to_a
  #
  # @api public
  class StreamRow
    [Enumerable].each { |m| include m }

    attr_reader :index, :height, :hidden, :custom_height, :outline_level, :style_index

    # @param index [Integer] 0-based row index.
    # @param xml_bytes [String] Raw ASCII-8BIT XML bytes.
    # @param from [Integer] Byte offset where cells start.
    # @param to [Integer] Byte offset where cells end.
    # @param shared_strings [Array<String>] Shared strings table.
    # @param prefix [String] XML namespace prefix (e.g. "x:" or "").
    # @param height [Float, Integer, nil] Row height in points.
    # @param hidden [Boolean] Whether the row is hidden.
    # @param custom_height [Boolean] Whether custom height is set.
    # @param outline_level [Integer, nil] Grouping/outline level.
    # @param style_index [Integer, nil] Style index.
    # @param date1904 [Boolean] Whether the 1904 date system is active.
    #: (index: Integer, xml_bytes: String, from: Integer, to: Integer, shared_strings: Array[String], ?prefix: String, ?height: Float | Integer | nil, ?hidden: bool, ?custom_height: bool, ?outline_level: Integer | nil, ?style_index: Integer | nil, ?styles: Hash[untyped, untyped]?, ?date1904: bool) -> void
    def initialize(index:, xml_bytes:, from:, to:, shared_strings:, prefix: "", height: nil, hidden: false,
                   custom_height: false, outline_level: nil, style_index: nil, styles: nil, date1904: false)
      @index = index
      @xml = xml_bytes
      @from = from
      @to = to
      @shared_strings = shared_strings
      @prefix = prefix
      @height = height
      @hidden = hidden
      @custom_height = custom_height
      @outline_level = outline_level
      @style_index = style_index
      @styles = styles
      @date1904 = date1904 ? true : false
      @cells = nil
    end

    # rubocop:disable Style/OptionalBooleanParameter
    def self.fast_create(index, xml_bytes, from, to, shared_strings, prefix = "", height = nil, hidden = false, custom_height = false, outline_level = nil, style_index = nil, styles = nil, date1904 = false)
      inst = allocate
      inst.instance_variable_set(:@index, index)
      inst.instance_variable_set(:@xml, xml_bytes)
      inst.instance_variable_set(:@from, from)
      inst.instance_variable_set(:@to, to)
      inst.instance_variable_set(:@shared_strings, shared_strings)
      inst.instance_variable_set(:@prefix, prefix)
      inst.instance_variable_set(:@height, height)
      inst.instance_variable_set(:@hidden, hidden)
      inst.instance_variable_set(:@custom_height, custom_height)
      inst.instance_variable_set(:@outline_level, outline_level)
      inst.instance_variable_set(:@style_index, style_index)
      inst.instance_variable_set(:@styles, styles)
      inst.instance_variable_set(:@date1904, date1904 ? true : false)
      inst.instance_variable_set(:@cells, nil)
      inst
    end
    # rubocop:enable Style/OptionalBooleanParameter

    # Returns whether the row uses the 1904 date system.
    #
    # @return [Boolean]
    # @api public
    #: () -> bool
    def date1904?
      @date1904 ? true : false
    end

    # Iterate over cells in this streaming row one by one.
    #
    # @yield [cell]
    # @yieldparam cell [Elements::Cell]
    # @return [Enumerator, void]
    # @api public
    #: () { (Elements::Cell) -> void } -> void
    #: () -> Enumerator[Elements::Cell, void]
    def each_cell(&block)
      return enum_for(:each_cell) unless block

      if @cells
        @cells.each(&block)
      else
        Ooxml::WorksheetParser.fast_scan_cells_direct(@xml, @from, @to, @shared_strings, @index, @prefix, @styles, @date1904, &block)
      end
    end

    # Iterate over cells in this streaming row.
    #
    # @yield [cell]
    # @yieldparam cell [Elements::Cell]
    # @return [Enumerator, void]
    # @api public
    #: () { (Elements::Cell) -> void } -> void
    #: () -> Enumerator[Elements::Cell, void]
    def each(&)
      each_cell(&)
    end

    # Returns all cells as an Array. Cached on first access.
    #
    # @return [Array<Elements::Cell>]
    # @api public
    #: () -> Array[Elements::Cell]
    def cells
      @cells ||= begin
        arr = []
        Ooxml::WorksheetParser.fast_scan_cells_direct(@xml, @from, @to, @shared_strings, @index, @prefix, @styles, @date1904) do |c|
          arr << c
        end
        arr.freeze
      end
    end

    # Access a cell by 0-based column index, or access row attributes via Symbol.
    #
    # @param col_index [Integer, Symbol] Column index or attribute symbol.
    # @return [Elements::Cell, Object, nil]
    # @api public
    #: (Integer | Symbol col_index) -> untyped
    def [](col_index)
      case col_index
      when Symbol
        case col_index
        when :cells then cells
        when :values then values
        when :raw_values then raw_values
        when :formatted_values then formatted_values
        when :formulas then formulas
        when :index then index
        when :height then height
        when :hidden then hidden
        when :custom_height then custom_height
        when :outline_level then outline_level
        when :style_index then style_index
        when :date1904 then date1904?
        when :attrs
          h = { height: height, hidden: hidden, custom_height: custom_height, outline_level: outline_level }
          h[:style_index] = style_index if style_index
          h
        end
      else
        cells[col_index]
      end
    end

    # Access a cell by 0-based column index.
    #
    # @param col_index [Integer] 0-based column index.
    # @return [Elements::Cell, nil]
    # @api public
    #: (Integer col_index) -> Elements::Cell?
    def cell_at(col_index)
      cells.find { |c| c.column_index == col_index }
    end

    # Convert row cells to an Array of raw values (sparse columns get nil).
    #
    # @return [Array<Object>]
    # @api public
    #: () -> Array[untyped]
    def to_a
      values
    end

    # Returns cell values as an Array (sparse columns get nil).
    #
    # @param type_cast [Boolean] Whether to coerce date/time serial numbers into Date/Time instances.
    # @return [Array<Object>]
    # @api public
    #: (?type_cast: bool) -> Array[untyped]
    def values(type_cast: false)
      if @cells
        values_from_cells(type_cast)
      else
        Ooxml::WorksheetParser.fast_scan_row_values(@xml, @from, @to, @shared_strings, @styles, @date1904, type_cast: type_cast)
      end
    end

    # Convert row cells to an Array of raw string values (sparse columns get nil).
    #
    # @return [Array<String, nil>]
    # @api public
    #: () -> Array[String?]
    def raw_values
      return [] if cells.empty?

      max_col = cells.map(&:column_index).max || 0
      arr = Array.new(max_col + 1)
      cells.each do |cell|
        arr[cell.column_index] = cell.raw_value
      end
      arr
    end

    # Convert row cells to an Array of formatted string representations (sparse columns get nil).
    #
    # @return [Array<String, nil>]
    # @api public
    #: () -> Array[String?]
    def formatted_values
      return [] if cells.empty?

      max_col = cells.map(&:column_index).max || 0
      arr = Array.new(max_col + 1)
      cells.each do |cell|
        arr[cell.column_index] = cell.formatted_value
      end
      arr
    end

    # Convert row cells to an Array of formula expressions without leading '=' (sparse columns get nil).
    #
    # @return [Array<String, nil>]
    # @api public
    #: () -> Array[String?]
    def formulas
      return [] if cells.empty?

      max_col = cells.map(&:column_index).max || 0
      arr = Array.new(max_col + 1)
      cells.each do |cell|
        arr[cell.column_index] = cell.formula_expression
      end
      arr
    end

    # Returns whether the row is valid according to OOXML specifications.
    #
    # @return [Boolean]
    # @api public
    #: () -> bool
    def valid?
      true
    end

    # Returns whether the row contains no cell values (or all cell values are nil or empty strings).
    #
    # @return [Boolean]
    # @api public
    #: () -> bool
    def empty?
      vals = values
      return true if vals.empty?

      vals.all? { |v| v.nil? || (v.is_a?(String) && v.empty?) }
    end

    # Unmapped metadata for compatibility with Elements::Row.
    #
    # @return [Hash]
    # @api public
    #: () -> Hash[untyped, untyped]
    def unmapped_data
      @style_index ? { style_index: @style_index } : Elements::EMPTY_HASH
    end

    # Validation errors for compatibility with Elements::Row.
    #
    # @return [Array<String>]
    # @api public
    #: () -> Array[String]
    def errors
      Elements::EMPTY_ERRORS
    end

    # Human-readable representation.
    #
    # @return [String]
    # @api public
    #: () -> String
    def inspect
      "#<#{self.class.name} index=#{index} height=#{height.inspect} hidden=#{hidden}>"
    end

    private

    #: (bool) -> Array[untyped]
    def values_from_cells(type_cast)
      return [] if @cells.nil? || @cells.empty?

      max_col = @cells.map(&:column_index).max || 0
      arr = Array.new(max_col + 1)
      @cells.each do |c|
        val = c.value
        if type_cast && val.is_a?(Numeric)
          fmt_type = if c.style_index && @styles
                       NumberFormatter.format_type_for(c.style_index, @styles)
                     elsif c.format_code
                       NumberFormatter.format_type(c.format_code)
                     end
          begin
            case fmt_type
            when :date
              val = Ooxml::Utils.serial_to_date(val, date1904: @date1904)
            when :datetime, :time
              val = Ooxml::Utils.serial_to_datetime(val, date1904: @date1904)
            end
          rescue StandardError
            # Keep val numeric on failure
          end
        end
        arr[c.column_index] = val
      end
      arr
    end
  end
end
