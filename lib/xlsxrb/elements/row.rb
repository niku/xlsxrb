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
    # rubocop:disable Style/DataInheritance -- Required as class syntax for mutant subject matcher
    class Row < Data.define(:index, :cells, :height, :hidden, :custom_height, :collapsed, :outline_level, :unmapped_data, :errors)
      # rubocop:enable Style/DataInheritance
      [Enumerable].each { |m| include m }

      # @param index [Integer] 0-based row index.
      # @param cells [Array<Elements::Cell>] Cells in this row.
      # @param height [Float, Integer, nil] Row height in points.
      # @param hidden [Boolean] Whether the row is hidden.
      # @param custom_height [Boolean] Whether custom height is set.
      # @param collapsed [Boolean] Whether the row is collapsed.
      # @param outline_level [Integer, nil] Grouping/outline level.
      # @param unmapped_data [Hash] Additional metadata.
      # @param errors [Array<String>, nil] Validation errors.
      #: (index: Integer, ?cells: Array[Elements::Cell], ?height: Float | Integer | nil, ?hidden: bool, ?custom_height: bool, ?collapsed: bool, ?outline_level: Integer | nil, ?unmapped_data: Hash[untyped, untyped], ?errors: Array[String]?) -> void
      def initialize(index:, cells: EMPTY_CELLS, height: nil, hidden: false, custom_height: false, collapsed: false,
                     outline_level: nil, unmapped_data: EMPTY_HASH, errors: nil)
        computed_errors = errors || self.class.validate(index, cells)
        computed_errors = computed_errors.freeze unless computed_errors.frozen?
        cells = cells.freeze unless cells.frozen?
        super(index: index, cells: cells, height: height, hidden: hidden,
              custom_height: custom_height, collapsed: collapsed, outline_level: outline_level,
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
          when :values then values
          when :raw_values then raw_values
          when :formatted_values then formatted_values
          when :formulas then formulas
          when :index then index
          when :height then height
          when :hidden then hidden
          when :custom_height then custom_height
          when :collapsed then collapsed
          when :outline_level then outline_level
          when :style_index then style_index
          when :hidden? then hidden?
          when :collapsed? then collapsed?
          when :custom_height? then custom_height?
          when :attributes then attributes
          when :cells_hash then cells_hash
          when :attrs
            h = { height: height, hidden: hidden, custom_height: custom_height, collapsed: collapsed, outline_level: outline_level }
            h[:style_index] = style_index if style_index
            h
          end
        else
          cells[col_index]
        end
      end

      # Returns a new Row with missing cells padded up to the maximum column index.
      #
      # @return [Elements::Row]
      # @api public
      #: () -> Elements::Row
      def pad_empty_cells
        return self if cells.empty?

        max_col = cells.map(&:column_index).max || 0
        return self if cells.size == max_col + 1

        is_d1904 = cells.first ? cells.first.date1904? : false
        cell_map = {}
        cells.each { |c| cell_map[c.column_index] = c }
        padded = Array.new(max_col + 1)
        (0..max_col).each do |c_idx|
          padded[c_idx] = cell_map[c_idx] || Cell.fast_create(index, c_idx, nil, nil, nil, nil, nil, is_d1904)
        end
        Row.new(
          index: index,
          cells: padded,
          height: height,
          hidden: hidden,
          custom_height: custom_height,
          collapsed: collapsed,
          outline_level: outline_level,
          unmapped_data: unmapped_data,
          errors: errors
        )
      end

      # Returns a coordinate-keyed Hash of cell values (e.g. {"A1" => "value", "B1" => 123}).
      #
      # @param type_cast [Boolean] Whether to coerce date/time serial numbers into Date/Time instances.
      # @return [Hash{String => Object}]
      # @api public
      #: (?type_cast: bool) -> Hash[String, untyped]
      def cells_hash(type_cast: false)
        result = {}
        cells.each do |c|
          val = c.value
          if type_cast && val.is_a?(Numeric) && c.format_code
            fmt_type = NumberFormatter.format_type(c.format_code)
            begin
              case fmt_type
              when :date
                val = Ooxml::Utils.serial_to_date(val, date1904: c.date1904?)
              when :datetime, :time
                val = Ooxml::Utils.serial_to_datetime(val, date1904: c.date1904?)
              end
            rescue StandardError
              # Keep val numeric on failure
            end
          end
          result[c.ref] = val
        end
        result
      end

      # Returns whether the row is hidden.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def hidden?
        hidden == true
      end

      # Returns whether the row is collapsed.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def collapsed?
        collapsed == true
      end

      # Returns whether custom height is set.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def custom_height?
        custom_height == true
      end

      # Returns a frozen Hash of row attributes.
      #
      # @return [Hash{Symbol => Object}]
      # @api public
      #: () -> Hash[Symbol, untyped]
      def attributes
        h = {
          index: index,
          height: height,
          hidden: hidden,
          custom_height: custom_height,
          collapsed: collapsed,
          outline_level: outline_level
        }
        h[:style_index] = style_index if style_index
        h.freeze
      end

      # Returns style_index if present in unmapped_data.
      #
      # @return [Integer, nil]
      #: () -> Integer?
      def style_index
        unmapped_data[:style_index]
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
      #: () -> Enumerator[Elements::Cell, void]
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
      #: () -> Enumerator[Elements::Cell, void]
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

      # Returns whether the row contains no cell values (or all cell values are nil or empty strings).
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def empty?
        return true if cells.empty?

        cells.all? { |c| c.value.nil? || (c.value.is_a?(String) && c.value.empty?) }
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
      # @param type_cast [Boolean] Whether to coerce date/time serial numbers into Date/Time instances.
      # @return [Array<Object>]
      # @api public
      #: (?type_cast: bool) -> Array[untyped]
      def values(type_cast: false)
        return [] if cells.empty?

        max_col = cells.max_by(&:column_index)&.column_index || 0
        result = Array.new(max_col + 1)
        cells.each do |c|
          val = c.value
          if type_cast && val.is_a?(Numeric) && c.format_code
            fmt_type = NumberFormatter.format_type(c.format_code)
            begin
              case fmt_type
              when :date
                val = Ooxml::Utils.serial_to_date(val, date1904: c.date1904?)
              when :datetime, :time
                val = Ooxml::Utils.serial_to_datetime(val, date1904: c.date1904?)
              end
            rescue StandardError
              # Keep val numeric on failure
            end
          end
          result[c.column_index] = val
        end
        result
      end

      # Returns unparsed raw cell values as an Array (sparse columns get nil).
      #
      # @return [Array<String, nil>]
      # @api public
      #: () -> Array[String?]
      def raw_values
        return [] if cells.empty?

        max_col = cells.max_by(&:column_index).column_index
        result = Array.new(max_col + 1)
        cells.each { |c| result[c.column_index] = c.raw_value }
        result
      end

      # Returns formatted string representations of cell values as an Array (sparse columns get nil).
      #
      # @return [Array<String, nil>]
      # @api public
      #: () -> Array[String?]
      def formatted_values
        return [] if cells.empty?

        max_col = cells.max_by(&:column_index).column_index
        result = Array.new(max_col + 1)
        cells.each { |c| result[c.column_index] = c.formatted_value }
        result
      end

      # Returns formula expressions without leading '=' as an Array (sparse columns get nil).
      #
      # @return [Array<String, nil>]
      # @api public
      #: () -> Array[String?]
      def formulas
        return [] if cells.empty?

        max_col = cells.max_by(&:column_index).column_index
        result = Array.new(max_col + 1)
        cells.each { |c| result[c.column_index] = c.formula_expression }
        result
      end

      # Returns whether the row index is within valid OOXML range (0..1048575).
      #
      # @param index [Object]
      # @return [Boolean]
      #: (untyped index) -> bool
      def self.valid_index?(index)
        case index
        when Integer
          index >= 0 && index < 1_048_576
        else
          false
        end
      end

      # Validates row index and cells against OOXML limits.
      #
      # @param index [Integer]
      # @param cells [Array<Elements::Cell>]
      # @return [Array<String>] List of errors.
      #: (untyped index, untyped cells) -> Array[String]
      def self.validate(index, cells)
        errs = []
        case index
        when Integer
          if index.negative?
            errs << "index must be a non-negative Integer (got #{index})"
          elsif index >= 1_048_576
            errs << "index must be < 1048576 (got #{index}, max row is 1048575)"
          end
        else
          errs << "index must be a non-negative Integer (got #{index.inspect})"
        end
        errs << "cells must be an Array (got #{cells.class})" unless cells.is_a?(Array)
        errs.empty? ? EMPTY_ERRORS : errs
      end
    end
  end
end
