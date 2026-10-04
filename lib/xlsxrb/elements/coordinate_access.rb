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
            ref = "#{Cell.column_letter(c.column_index)}#{c.row_index + 1}"
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

      # Access a cell by its Excel-style reference (e.g. "A1").
      #
      # @example
      #   sheet["A1"] #=> #<Elements::Cell value="Hello">
      #
      # @param ref [String, Symbol] Cell reference (e.g. "A1" or :A1).
      # @return [Elements::Cell, nil]
      # @api public
      #: (String | Symbol ref) -> Elements::Cell?
      def [](ref)
        cells_hash[ref.to_s.upcase]
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

      # Returns the 1-based index of the first column containing data, or nil.
      #
      # @return [Integer, nil]
      # @api public
      #: () -> Integer?
      def first_column
        min_c = nil
        rows.each do |r|
          r.cells.each do |c|
            min_c = c.column_index if min_c.nil? || c.column_index < min_c
          end
        end
        min_c ? min_c + 1 : nil
      end
      alias first_col first_column

      # Returns the 1-based index of the last column containing data, or nil.
      #
      # @return [Integer, nil]
      # @api public
      #: () -> Integer?
      def last_column
        max_c = nil
        rows.each do |r|
          r.cells.each do |c|
            max_c = c.column_index if max_c.nil? || c.column_index > max_c
          end
        end
        max_c ? max_c + 1 : nil
      end
      alias last_col last_column

      # Returns the raw cell value at the given Excel-style reference (e.g. "A1").
      #
      # @example
      #   sheet.cell_value("A1") #=> "Sales Report"
      #
      # @param ref [String] Cell reference (e.g. "A1").
      # @return [Object, nil]
      # @api public
      #: (String ref) -> untyped
      def cell_value(ref)
        parsed = Cell.parse_ref(ref)
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
