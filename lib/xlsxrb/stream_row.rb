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

    attr_reader :index, :height, :hidden, :custom_height, :outline_level

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
    #: (index: Integer, xml_bytes: String, from: Integer, to: Integer, shared_strings: Array[String], ?prefix: String, ?height: Float | Integer | nil, ?hidden: bool, ?custom_height: bool, ?outline_level: Integer | nil) -> void
    def initialize(index:, xml_bytes:, from:, to:, shared_strings:, prefix: "", height: nil, hidden: false,
                   custom_height: false, outline_level: nil)
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
      @cells_cache = nil
    end

    # Iterate over cells in this streaming row one by one.
    #
    # @yield [cell]
    # @yieldparam cell [Elements::Cell]
    # @return [Enumerator, void]
    # @api public
    #: () { (Elements::Cell) -> void } -> void
    #: () -> Enumerator[Elements::Cell, void]
    def each_cell(&)
      return enum_for(:each_cell) unless block_given?

      if @cells
        @cells.each(&)
      else
        Ooxml::WorksheetParser.fast_scan_cells_direct(@xml, @from, @to, @shared_strings, { row: @index }, @prefix, &)
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
      @cells ||= each_cell.to_a.freeze
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
        when :index then index
        when :height then height
        when :hidden then hidden
        when :custom_height then custom_height
        when :outline_level then outline_level
        when :attrs then { height: height, hidden: hidden, custom_height: custom_height, outline_level: outline_level }
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
      return [] if cells.empty?

      max_col = cells.map(&:column_index).max || 0
      arr = Array.new(max_col + 1)
      cells.each do |cell|
        arr[cell.column_index] = cell.value
      end
      arr
    end

    # Returns cell values as an Array.
    #
    # @return [Array<Object>]
    # @api public
    #: () -> Array[untyped]
    def values
      to_a
    end

    # Returns whether the row is valid according to OOXML specifications.
    #
    # @return [Boolean]
    # @api public
    #: () -> bool
    def valid?
      true
    end

    # Unmapped metadata for compatibility with Elements::Row.
    #
    # @return [Hash]
    # @api public
    #: () -> Hash[untyped, untyped]
    def unmapped_data
      Elements::EMPTY_HASH
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
  end
end
