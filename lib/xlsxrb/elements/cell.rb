# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Elements
    # Represents a single cell in a worksheet.
    # All indices are 0-based.
    Cell = Data.define(:row_index, :column_index, :value, :formula, :style_index, :unmapped_data, :errors) do
      def initialize(row_index:, column_index:, value: nil, formula: nil, style_index: nil, unmapped_data: EMPTY_HASH, errors: nil)
        computed_errors = errors || self.class.validate(row_index, column_index, value)
        computed_errors = computed_errors.freeze unless computed_errors.frozen?
        super(row_index: row_index, column_index: column_index, value: value, formula: formula,
              style_index: style_index, unmapped_data: unmapped_data, errors: computed_errors)
      end

      def valid?
        errors.empty?
      end

      # Excel-style reference (e.g. "A1").
      def ref
        "#{self.class.column_letter(column_index)}#{row_index + 1}"
      end

      def [](key)
        case key
        when :value then value
        when :formula then formula
        when :style_index then style_index
        when :ref then ref
        when :column_index then column_index
        when :row_index then row_index
        when :type
          case value
          when String then "s"
          when true, false then "b"
          end
        end
      end

      def content
        value
      end

      def to_s
        value.to_s
      end

      def to_i
        value.to_i
      end

      def to_f
        value.to_f
      end

      def to_date
        return value if value.is_a?(Date)

        # Excel epoch is 1899-12-30
        if value.is_a?(Numeric)
          Date.new(1899, 12, 30) + value.to_i
        else
          begin
            Date.parse(value.to_s)
          rescue StandardError
            nil
          end
        end
      end

      def to_time
        return value if value.is_a?(Time)

        if value.is_a?(Numeric)
          days = value.to_f
          base_time = Time.utc(1899, 12, 30)
          base_time + (days * 86_400)
        else
          begin
            Time.parse(value.to_s)
          rescue StandardError
            nil
          end
        end
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
        raise ArgumentError, "Column index must be a non-negative Integer, got #{index.inspect}" unless index.is_a?(Integer) && index >= 0

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

      # Converts a column letter (e.g. "A", :AA) to a 0-based column index.
      # If passed an integer or string/symbol representing an integer, it validates and returns the integer.
      def self.column_index(letter)
        if letter.is_a?(Integer)
          raise ArgumentError, "Column index must be >= 0, got #{letter}" if letter.negative?

          return letter
        end

        str = letter.to_s
        if str.match?(/\A-?\d+\z/)
          val = str.to_i
          raise ArgumentError, "Column index must be >= 0, got #{val}" if val.negative?

          return val
        end

        raise ArgumentError, "Invalid column letter: #{letter.inspect}" unless str.match?(/\A[a-zA-Z]+\z/)

        str.upcase.chars.reduce(0) { |acc, c| (acc * 26) + (c.ord - "A".ord + 1) } - 1
      end

      # Parses an Excel-style reference to [row_index, col_index] (both 0-based).
      def self.parse_ref(ref)
        return nil unless ref

        bytes = ref.b
        len = bytes.bytesize
        col = 0
        i = 0
        while i < len
          b = bytes.getbyte(i)
          if b.between?(65, 90)
            col = (col * 26) + (b - 64)
            i += 1
          elsif b.between?(97, 122)
            col = (col * 26) + (b - 96)
            i += 1
          else
            break
          end
        end
        return nil if i.zero? || i == len

        row = bytes.byteslice(i, len - i).to_i - 1
        [row, col - 1]
      end

      def self.validate(row_index, column_index, value)
        if row_index.is_a?(Integer) && row_index >= 0 && row_index < 1_048_576 &&
           column_index.is_a?(Integer) && column_index >= 0 && column_index < 16_384 &&
           (value.nil? || value.is_a?(String) || value.is_a?(Numeric) || value == true || value == false || value.is_a?(Date) || value.is_a?(Time) || value.is_a?(Formula) || (value.is_a?(Hash) && value.key?(:formula)) || value.is_a?(RichText) || value.is_a?(CellError))
          return EMPTY_ERRORS
        end

        errs = []
        errs << "row_index must be a non-negative Integer (got #{row_index.inspect})" if !row_index.is_a?(Integer) || row_index.negative?
        errs << "column_index must be a non-negative Integer (got #{column_index.inspect})" if !column_index.is_a?(Integer) || column_index.negative?
        errs << "row_index must be < 1048576 (got #{row_index}, max row is 1048575)" if row_index.is_a?(Integer) && row_index >= 1_048_576
        errs << "column_index must be < 16384 (got #{column_index}, max column is XFD=16383)" if column_index.is_a?(Integer) && column_index >= 16_384
        errs << "unsupported value type: #{value.class} (#{value.inspect}) — supported types: String, Numeric, true/false, Date, Time, or nil" unless value.nil? || value.is_a?(String) || value.is_a?(Numeric) || value == true || value == false || value.is_a?(Date) || value.is_a?(Time)
        errs
      end
    end
  end
end
