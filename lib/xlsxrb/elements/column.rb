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
    # rubocop:disable Style/DataInheritance -- Required as class syntax for mutant subject matcher
    class Column < Data.define(:index, :width, :hidden, :custom_width, :outline_level, :unmapped_data, :errors)
      # rubocop:enable Style/DataInheritance
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

      # Returns whether the column index is within valid OOXML range (0..16383).
      #
      # @param index [Object]
      # @return [Boolean]
      #: (untyped index) -> bool
      def self.valid_index?(index)
        case index
        when Integer
          index >= 0 && index < 16_384
        else
          false
        end
      end

      # Validates column index against OOXML limits.
      #
      # @param index [Integer]
      # @return [Array<String>] List of errors.
      #: (untyped index) -> Array[String]
      def self.validate(index)
        errs = []
        case index
        when Integer
          if index.negative?
            errs << "index must be a non-negative Integer (got #{index})"
          elsif index >= 16_384
            errs << "index must be < 16384 (got #{index}, max column is XFD=16383)"
          end
        else
          errs << "index must be a non-negative Integer (got #{index.inspect})"
        end
        errs
      end
    end
  end
end
