# frozen_string_literal: true

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

      def cells
        # Ensure ordered traversal
        @cells ||= cells_hash.values.sort_by { |c| [c.row_index, c.column_index] }
      end

      def [](ref)
        cells_hash[ref.to_s.upcase]
      end

      def each(&block)
        return to_enum(:each) unless block_given?
        cells.each(&block)
      end

      def each_cell(&block)
        return to_enum(:each_cell) unless block_given?
        cells.each(&block)
      end

      def each_row(&block)
        return to_enum(:each_row) unless block_given?
        rows.each(&block)
      end

      def cells
        return to_enum(:cells) unless block_given?
        rows.each do |row|
          row.each { |cell| yield cell }
        end
      end

      def valid?
        errors.empty?
      end

      # Returns the row at the given 0-based index, or nil.
      def row_at(index)
        rows.find { |r| r.index == index }
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

      def self.validate(name, rows)
        errs = []
        errs << "worksheet name must be a non-empty String (got #{name.inspect})" if name.nil? || (name.is_a?(String) && name.empty?)
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
