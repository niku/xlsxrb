# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Elements
    # Represents a single worksheet in a workbook.
    Worksheet = Data.define(:name, :rows, :columns, :charts, :unmapped_data, :errors) do
      include Enumerable

      def initialize(name:, rows: [], columns: [], charts: [], unmapped_data: {}, errors: nil)
        computed_errors = errors || self.class.validate(name, rows)
        super(name: name, rows: rows.freeze, columns: columns.freeze, charts: charts.freeze,
              unmapped_data: unmapped_data, errors: computed_errors.freeze)
      end

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

      def cells
        # Ensure ordered traversal
        cells_hash.values.sort_by { |c| [c.row_index, c.column_index] }
      end

      def [](ref)
        cells_hash[ref.to_s.upcase]
      end

      def each(&)
        return to_enum(:each) unless block_given?

        cells.each(&)
      end

      def each_cell(&)
        return to_enum(:each_cell) unless block_given?

        cells.each(&)
      end

      def each_row(&)
        return to_enum(:each_row) unless block_given?

        rows.each(&)
      end

      def valid?
        errors.empty?
      end

      # Returns the row at the given 0-based index, or nil.
      def row_at(index)
        rows.find { |r| r.index == index }
      end

      def first_row
        rows.min_by(&:index)
      end

      def last_row
        rows.max_by(&:index)
      end

      # Returns cell value at Excel-style reference (e.g. "A1").
      def cell_value(ref)
        parsed = Cell.parse_ref(ref)
        return nil unless parsed

        row_idx, col_idx = parsed
        row = row_at(row_idx)
        return nil unless row

        cell = row.cell_at(col_idx)
        cell&.value
      end

      # Returns a new Worksheet with the specified cell updated.
      #
      # @param ref [String] The cell reference (e.g. "B1").
      # @param value [Object] The new cell value.
      # @param style_index [Integer, String, nil] Optional new style index.
      # @param formula [Elements::Formula, nil] Optional new formula.
      # @return [Worksheet] A new Worksheet instance.
      def update_cell(ref, value: nil, style_index: nil, formula: nil)
        parsed = Cell.parse_ref(ref)
        raise ArgumentError, "invalid cell reference: #{ref}" unless parsed

        row_idx, col_idx = parsed
        existing_row = row_at(row_idx)

        if existing_row
          existing_cell = existing_row.cell_at(col_idx)
          new_cell = if existing_cell
                       existing_cell.with(
                         value: value || existing_cell.value,
                         style_index: style_index || existing_cell.style_index,
                         formula: formula || existing_cell.formula
                       )
                     else
                       Cell.new(row_index: row_idx, column_index: col_idx, value: value, style_index: style_index, formula: formula)
                     end

          # Replace cell in the existing row
          new_cells = existing_row.cells.reject { |c| c.column_index == col_idx }
          new_cells << new_cell
          new_cells.sort_by!(&:column_index)

          new_row = existing_row.with(cells: new_cells)
          new_rows = rows.map { |r| r.index == row_idx ? new_row : r }
        else
          # Row doesn't exist, create it
          new_cell = Cell.new(row_index: row_idx, column_index: col_idx, value: value, style_index: style_index, formula: formula)
          new_row = Row.new(index: row_idx, cells: [new_cell])
          new_rows = (rows + [new_row]).sort_by!(&:index)
        end
        with(rows: new_rows)
      end

      def self.validate(name, rows)
        errs = []
        if name.nil? || !name.is_a?(String) || name.empty?
          errs << "worksheet name must be a non-empty String (got #{name.inspect})"
        else
          errs << "worksheet name cannot exceed 31 characters (got #{name.size})" if name.size > 31
          errs << "worksheet name cannot contain \\, /, ?, *, [, or ]" if name.match?(%r{[\\/?*\[\]]})
        end
        errs << "rows must be an Array (got #{rows.class})" unless rows.is_a?(Array)
        if rows.is_a?(Array)
          indices = rows.map(&:index)
          if indices.uniq.size != indices.size
            dups = indices.select { |i| indices.count(i) > 1 }.uniq
            errs << "duplicate row index: #{dups.join(", ")} — row indices within a sheet must be unique"
          end
        end
        errs
      end
    end
  end
end
