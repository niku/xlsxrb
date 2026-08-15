# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Elements
    # Represents column formatting in a worksheet.
    # index is 0-based.
    #
    # @example
    #   col = Elements::Column.new(index: 0, width: 25.0)
    #
    # @api public
    Column = Data.define(:index, :width, :hidden, :custom_width, :outline_level, :unmapped_data, :errors) do
      # @param index [Integer] 0-based column index.
      # @param width [Float, Integer, nil] Column width in characters.
      # @param hidden [Boolean] Whether the column is hidden.
      # @param custom_width [Boolean] Whether custom width flag is set.
      # @param outline_level [Integer, nil] Grouping/outline level.
      # @param unmapped_data [Hash] Additional metadata.
      # @param errors [Array<String>, nil] Validation errors.
      #: (index: Integer, ?width: Float | Integer | nil, ?hidden: bool, ?custom_width: bool, ?outline_level: Integer | nil, ?unmapped_data: Hash[untyped, untyped], ?errors: Array[String]?) -> void
      def initialize(index:, width: nil, hidden: false, custom_width: false, outline_level: nil,
                     unmapped_data: {}, errors: nil)
        computed_errors = errors || self.class.validate(index)
        super(index: index, width: width, hidden: hidden, custom_width: custom_width,
              outline_level: outline_level, unmapped_data: unmapped_data,
              errors: computed_errors.freeze)
      end

      # Returns whether the column definition is valid according to OOXML specifications.
      #
      # @return [Boolean]
      #: () -> bool
      def valid?
        errors.empty?
      end

      # Validates column index against OOXML limits.
      #
      # @param index [Integer]
      # @return [Array<String>] List of errors.
      #: (untyped index) -> Array[String]
      def self.validate(index)
        errs = []
        errs << "index must be a non-negative Integer (got #{index.inspect})" if !index.is_a?(Integer) || index.negative?
        errs << "index must be < 16384 (got #{index}, max column is XFD=16383)" if index.is_a?(Integer) && index >= 16_384
        errs
      end
    end
  end
end
