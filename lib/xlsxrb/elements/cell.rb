# frozen_string_literal: true

module Xlsxrb
  module Elements
    # Represents a single cell in a worksheet.
    # All indices are 0-based.
    Cell = Data.define(:row_index, :column_index, :value, :formula, :style_index, :unmapped_data, :errors) do
      def initialize(row_index:, column_index:, value: nil, formula: nil, style_index: nil, unmapped_data: {}, errors: nil)
        computed_errors = errors || self.class.validate(row_index, column_index, value)
        super(row_index: row_index, column_index: column_index, value: value, formula: formula,
              style_index: style_index, unmapped_data: unmapped_data, errors: computed_errors.freeze)
      end

      def valid?
        errors.empty?
      end

      # Excel-style reference (e.g. "A1").
      def ref
        "#{self.class.column_letter(column_index)}#{row_index + 1}"
      end

      # Cache column letters up to Excel's limit (16,384)
      @column_letters = (0...16_384).map do |index|
        result = +""
        i = index
        loop do
          result.prepend(("A".ord + (i % 26)).chr)
          i = (i / 26) - 1
          break if i.negative?
        end
        result.freeze
      end.freeze

      # Converts a 0-based column index to a letter (0 -> "A", 25 -> "Z", 26 -> "AA").
      def self.column_letter(index)
        @column_letters[index] || begin
          result = +""
          i = index
          loop do
            result.prepend(("A".ord + (i % 26)).chr)
            i = (i / 26) - 1
            break if i.negative?
          end
          result
        end
      end

      # Parses an Excel-style reference to [row_index, col_index] (both 0-based).
      def self.parse_ref(ref)
        match = ref.match(/\A([A-Z]+)(\d+)\z/)
        return nil unless match

        col = match[1].chars.reduce(0) { |acc, c| (acc * 26) + (c.ord - "A".ord + 1) } - 1
        row = match[2].to_i - 1
        [row, col]
      end

      def self.validate(row_index, column_index, value)
        errs = []
        errs << "row_index must be >= 0" if !row_index.is_a?(Integer) || row_index.negative?
        errs << "column_index must be >= 0" if !column_index.is_a?(Integer) || column_index.negative?
        errs << "row_index must be < 1048576" if row_index.is_a?(Integer) && row_index >= 1_048_576
        errs << "column_index must be < 16384" if column_index.is_a?(Integer) && column_index >= 16_384
        errs << "unsupported value type: #{value.class}" unless value.nil? || value.is_a?(String) || value.is_a?(Numeric) || value.is_a?(TrueClass) || value.is_a?(FalseClass) || value.is_a?(Date) || value.is_a?(Time)
        errs
      end
    end
  end
end
