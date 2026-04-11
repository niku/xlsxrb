# frozen_string_literal: true

module Xlsxrb
  module Elements
    # Represents a single worksheet in a workbook.
    Worksheet = Data.define(:name, :rows, :columns, :charts, :unmapped_data, :errors) do
      def initialize(name:, rows: [], columns: [], charts: [], unmapped_data: {}, errors: nil)
        computed_errors = errors || self.class.validate(name, rows)
        super(name: name, rows: rows.freeze, columns: columns.freeze, charts: charts.freeze,
              unmapped_data: unmapped_data, errors: computed_errors.freeze)
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
        errs << "name must be present" if name.nil? || (name.is_a?(String) && name.empty?)
        errs << "rows must be an Array" unless rows.is_a?(Array)
        if rows.is_a?(Array)
          indices = rows.map(&:index)
          errs << "duplicate row indices" if indices.uniq.size != indices.size
        end
        errs
      end
    end
  end
end
