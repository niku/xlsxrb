# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Elements
    # Represents a single row in a worksheet.
    # All row and column indices are 0-based.
    #
    # @example Access cell by index or symbol
    #   row = sheet.row_at(0)
    #   cell = row[0]       # cell at column 0
    #   row.to_a            # array of cell values
    #
    # @api public
    Row = Data.define(:index, :cells, :height, :hidden, :custom_height, :outline_level, :unmapped_data, :errors) do
      include Enumerable

      # @param index [Integer] 0-based row index.
      # @param cells [Array<Elements::Cell>] Cells in this row.
      # @param height [Float, Integer, nil] Row height in points.
      # @param hidden [Boolean] Whether the row is hidden.
      # @param custom_height [Boolean] Whether custom height is set.
      # @param outline_level [Integer, nil] Grouping/outline level.
      # @param unmapped_data [Hash] Additional metadata.
      # @param errors [Array<String>, nil] Validation errors.
      #: (index: Integer, ?cells: Array[Elements::Cell], ?height: Float | Integer | nil, ?hidden: bool, ?custom_height: bool, ?outline_level: Integer | nil, ?unmapped_data: Hash[untyped, untyped], ?errors: Array[String]?) -> void
      def initialize(index:, cells: EMPTY_CELLS, height: nil, hidden: false, custom_height: false, outline_level: nil,
                     unmapped_data: EMPTY_HASH, errors: nil)
        computed_errors = errors || self.class.validate(index, cells)
        computed_errors = computed_errors.freeze unless computed_errors.frozen?
        cells = cells.freeze unless cells.frozen?
        super(index: index, cells: cells, height: height, hidden: hidden,
              custom_height: custom_height, outline_level: outline_level,
              unmapped_data: unmapped_data, errors: computed_errors)
      end

      # Access a cell by 0-based column index, or access row attributes via Symbol.
      #
      # @example
      #   row[0]          #=> Cell at column 0
      #   row[:height]    #=> 25.0
      #   row[:cells]     #=> [Cell, Cell, ...]
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

      # Iterate over cells in this row.
      #
      # @example
      #   row.each do |cell|
      #     puts cell.value
      #   end
      #
      # @yield [cell]
      # @yieldparam cell [Elements::Cell]
      # @return [Enumerator, void]
      # @api public
      #: () { (Elements::Cell) -> void } -> void
      #: | () -> Enumerator[Elements::Cell, void]
      def each(&)
        return to_enum(:each) unless block_given?

        cells.each(&)
      end

      # Iterate over cells in this row.
      #
      # @example
      #   row.each_cell do |cell|
      #     puts "#{cell.ref}: #{cell.value}"
      #   end
      #
      # @yield [cell]
      # @yieldparam cell [Elements::Cell]
      # @return [Enumerator, void]
      # @api public
      #: () { (Elements::Cell) -> void } -> void
      #: | () -> Enumerator[Elements::Cell, void]
      def each_cell(&)
        return to_enum(:each_cell) unless block_given?

        cells.each(&)
      end

      # Convert row cells to an Array of raw values.
      #
      # @example
      #   row.to_a #=> ["ID", "Name", "Total"]
      #
      # @return [Array<Object>]
      # @api public
      #: () -> Array[untyped]
      def to_a
        return [] if cells.empty?

        max_col = cells.map(&:column_index).max
        arr = Array.new(max_col + 1)
        cells.each do |cell|
          arr[cell.column_index] = cell.value
        end
        arr
      end

      # Returns whether the row is valid according to OOXML specifications.
      #
      # @return [Boolean]
      #: () -> bool
      def valid?
        errors.empty?
      end

      # Returns the cell at the given 0-based column index, or nil.
      #
      # @param column_index [Integer] 0-based column index.
      # @return [Elements::Cell, nil]
      # @api public
      #: (Integer column_index) -> Elements::Cell?
      def cell_at(column_index)
        cells.find { |c| c.column_index == column_index }
      end

      # Returns cell values as an Array (sparse columns get nil).
      #
      # @return [Array<Object>]
      # @api public
      #: () -> Array[untyped]
      def values
        return [] if cells.empty?

        max_col = cells.max_by(&:column_index).column_index
        result = Array.new(max_col + 1)
        cells.each { |c| result[c.column_index] = c.value }
        result
      end

      # Validates row index and cells against OOXML limits.
      #
      # @param index [Integer]
      # @param cells [Array<Elements::Cell>]
      # @return [Array<String>] List of errors.
      #: (untyped index, untyped cells) -> Array[String]
      def self.validate(index, cells)
        return EMPTY_ERRORS if index.is_a?(Integer) && index >= 0 && index < 1_048_576 && cells.is_a?(Array)

        errs = []
        if !index.is_a?(Integer) || index.negative?
          errs << "index must be a non-negative Integer (got #{index.inspect})"
        elsif index >= 1_048_576
          errs << "index must be < 1048576 (got #{index}, max row is 1048576)"
        end
        errs << "cells must be an Array (got #{cells.class})" unless cells.is_a?(Array)
        errs
      end
    end
  end
end
