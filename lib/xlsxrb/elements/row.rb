# frozen_string_literal: true

module Xlsxrb
  module Elements
    # Represents a single row in a worksheet.
    # index is 0-based.
    Row = Data.define(:index, :cells, :height, :hidden, :custom_height, :outline_level, :unmapped_data, :errors) do
      def initialize(index:, cells: [], height: nil, hidden: false, custom_height: false, outline_level: nil,
                     unmapped_data: {}, errors: nil)
        computed_errors = errors || self.class.validate(index, cells)
        super(index: index, cells: cells.freeze, height: height, hidden: hidden,
              custom_height: custom_height, outline_level: outline_level,
              unmapped_data: unmapped_data, errors: computed_errors.freeze)
      end

      def valid?
        errors.empty?
      end

      # Returns the cell at the given 0-based column index, or nil.
      def cell_at(column_index)
        cells.find { |c| c.column_index == column_index }
      end

      # Returns cell values as an Array (sparse columns get nil).
      def values
        return [] if cells.empty?

        max_col = cells.max_by(&:column_index).column_index
        result = Array.new(max_col + 1)
        cells.each { |c| result[c.column_index] = c.value }
        result
      end

      def self.validate(index, cells)
        errs = []
        errs << "index must be >= 0" if !index.is_a?(Integer) || index.negative?
        errs << "cells must be an Array" unless cells.is_a?(Array)
        errs
      end
    end
  end
end
