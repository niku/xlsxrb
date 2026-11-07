# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Elements
    # Mixin providing coordinate-based and random-access cell/row lookups
    # for in-memory worksheet structures.
    #
    # Expects the including class to provide a `#rows` method returning an `Array<Elements::Row>`.
    #
    # @api public
    module CoordinateAccess
      # Returns a Hash mapping Excel cell references (e.g. "A1") to Cell objects.
      #
      # @return [Hash<String, Elements::Cell>]
      #: () -> Hash[String, Elements::Cell]
      def cells_hash
        h = {}
        rows.each do |r|
          r.cells.each do |c|
            ref = "#{Cell._column_letter(c.column_index)}#{c.row_index + 1}"
            h[ref] = c
          end
        end
        h
      end

      # Returns all cells ordered by row and column index.
      #
      # @return [Array<Elements::Cell>]
      # @api public
      #: () -> Array[Elements::Cell]
      def cells
        cells_hash.values.sort_by { |c| [c.row_index, c.column_index] }
      end

      # Access a cell by Excel-style reference (e.g. "A1") or 0-based coordinate indices (row, col).
      #
      # @example Coordinate string
      #   sheet["A1"] #=> #<Elements::Cell value="Hello">
      #
      # @example Row and column indices (0-based)
      #   sheet[0, 0] #=> #<Elements::Cell value="Hello">
      #   sheet[0, "A"] #=> #<Elements::Cell value="Hello">
      #
      # @param ref_or_row [String, Symbol, Integer] Cell reference (e.g. "A1" or :A1) or 0-based row index.
      # @param col [Integer, String, Symbol, nil] Optional 0-based column index or column letter.
      # @return [Elements::Cell, nil]
      # @api public
      #: (String | Symbol | Integer ref_or_row, ?Integer | String | Symbol | nil col) -> Elements::Cell?
      def [](ref_or_row, col = nil)
        if col
          row_at(ref_or_row)&.cell_at(col)
        else
          cells_hash[ref_or_row.to_s.upcase]
        end
      end

      # Returns the row at the given 0-based index, or nil.
      #
      # @param index [Integer] 0-based row index.
      # @return [Elements::Row, nil]
      # @api public
      #: (Integer index) -> Elements::Row?
      def row_at(index)
        rows.find { |r| r.index == index }
      end

      # Returns the first row in the sheet, or nil.
      #
      # @return [Elements::Row, nil]
      # @api public
      #: () -> Elements::Row?
      def first_row
        rows.min_by(&:index)
      end

      # Returns the last row in the sheet, or nil.
      #
      # @return [Elements::Row, nil]
      # @api public
      #: () -> Elements::Row?
      def last_row
        rows.max_by(&:index)
      end

      # Returns 0-based bounds { min_row: r1, max_row: r2, min_column: c1, max_column: c2 } from dimension ref.
      #
      # @example
      #   sheet.dimension_bounds #=> { min_row: 0, max_row: 4, min_column: 0, max_column: 2 }
      #
      # @return [Hash{Symbol => Integer}, nil]
      # @api public
      #: () -> Hash[Symbol, Integer]?
      def dimension_bounds
        return nil unless respond_to?(:dimension)

        ref = dimension
        return nil unless ref

        parts = ref.split(":")
        return nil unless [1, 2].include?(parts.size)

        ref1, ref2 = parts
        first_coords = Cell._parse_ref(ref1)
        return nil unless first_coords

        last_coords = ref2 ? Cell._parse_ref(ref2) : first_coords
        return nil unless last_coords

        min_r, min_c = first_coords
        max_r, max_c = last_coords

        {
          min_row: min_r,
          max_row: max_r,
          min_column: min_c,
          max_column: max_c
        }.freeze
      end

      # Returns the 1-based index of the first column containing data, or nil.
      #
      # @deprecated Use {#column_range} or {#dimension_bounds} instead.
      # @return [Integer, nil]
      # @api public
      #: () -> Integer?
      def first_column
        warn "[DEPRECATION] `first_column` is deprecated; use `column_range` or `dimension_bounds` instead.", uplevel: 1
        min_c = nil
        rows.each do |r|
          r.cells.each do |c|
            min_c = c.column_index if min_c.nil? || c.column_index < min_c
          end
        end
        min_c ? min_c + 1 : nil
      end

      # @deprecated Use {#column_range} or {#dimension_bounds} instead. Will be removed in 0.2.0.
      #: () -> Integer?
      def first_col
        warn "[DEPRECATION] `first_col` is deprecated; use `column_range` or `dimension_bounds` instead.", uplevel: 1
        first_column
      end

      # Returns the 1-based index of the last column containing data, or nil.
      #
      # @deprecated Use {#column_range} or {#dimension_bounds} instead.
      # @return [Integer, nil]
      # @api public
      #: () -> Integer?
      def last_column
        warn "[DEPRECATION] `last_column` is deprecated; use `column_range` or `dimension_bounds` instead.", uplevel: 1
        max_c = nil
        rows.each do |r|
          r.cells.each do |c|
            max_c = c.column_index if max_c.nil? || c.column_index > max_c
          end
        end
        max_c ? max_c + 1 : nil
      end

      # @deprecated Use {#column_range} or {#dimension_bounds} instead. Will be removed in 0.2.0.
      #: () -> Integer?
      def last_col
        warn "[DEPRECATION] `last_col` is deprecated; use `column_range` or `dimension_bounds` instead.", uplevel: 1
        last_column
      end

      # Returns the 0-based range of rows containing data, or nil.
      #
      # @example
      #   sheet.row_range #=> 0..9
      #
      # @return [Range<Integer>, nil]
      # @api public
      #: () -> Range[Integer]?
      def row_range
        r_first = first_row&.index
        r_last = last_row&.index
        return nil unless r_first && r_last

        (r_first..r_last)
      end

      # Returns the 0-based range of columns containing data, or nil.
      #
      # @example
      #   sheet.column_range #=> 0..4
      #
      # @return [Range<Integer>, nil]
      # @api public
      #: () -> Range[Integer]?
      def column_range
        indices = [] #: Array[Integer]
        rows.each do |r|
          r.cells.each do |c|
            indices << c.column_index
          end
        end
        return nil if indices.empty?

        (indices.min..indices.max)
      end

      # Returns the raw cell value at the given Excel-style reference (e.g. "A1").
      #
      # @deprecated Use `sheet["A1"]&.value` instead.
      # @example
      #   sheet.cell_value("A1") #=> "Sales Report"
      #
      # @param ref [String] Cell reference (e.g. "A1").
      # @return [Object, nil]
      # @api public
      #: (String ref) -> untyped
      def cell_value(ref)
        warn "[DEPRECATION] `cell_value` is deprecated; use `sheet[\"#{ref}\"]&.value` instead.", uplevel: 1
        parsed = Cell._parse_ref(ref)
        return nil unless parsed

        row_idx, col_idx = parsed
        row = row_at(row_idx)
        return nil unless row

        cell = row.cell_at(col_idx)
        cell&.value
      end
    end
  end
end
