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
    class Column < Data.define(:index, :width, :hidden, :custom_width, :outline_level, :style_index, :unmapped_data, :errors, :best_fit, :collapsed)
      # rubocop:enable Style/DataInheritance

      undef_method :hidden, :custom_width, :best_fit, :collapsed
      # @param index [Integer] 0-based column index.
      # @param width [Float, Integer, nil] Column width in characters.
      # @param hidden [Boolean] Whether the column is hidden.
      # @param custom_width [Boolean] Whether custom width flag is set.
      # @param outline_level [Integer, nil] Grouping/outline level.
      # @param style_index [Integer, String, Symbol, Hash, nil] Optional style identifier or inline style.
      # @param unmapped_data [Hash] Additional metadata.
      # @param errors [Array<String>, nil] Validation errors.
      # @param best_fit [Boolean] Whether the column automatically fits content.
      # @param collapsed [Boolean] Whether the column is collapsed.
      #: (index: Integer, ?width: Float | Integer | nil, ?hidden: bool, ?custom_width: bool, ?outline_level: Integer | nil, ?style_index: (Integer | String | Symbol | Hash[Symbol, untyped])?, ?unmapped_data: Hash[untyped, untyped], ?errors: Array[String]?, ?best_fit: bool, ?collapsed: bool) -> void
      def initialize(index:, width: nil, hidden: false, custom_width: false, outline_level: nil,
                     style_index: nil, unmapped_data: {}, errors: nil, best_fit: false, collapsed: false)
        @hidden = hidden ? true : false
        @custom_width = custom_width ? true : false
        @best_fit = best_fit ? true : false
        @collapsed = collapsed ? true : false
        computed_errors = errors || self.class.validate(index)
        super(**{ index: index, width: width, hidden: @hidden, custom_width: @custom_width,
                  outline_level: outline_level, style_index: style_index, unmapped_data: unmapped_data,
                  errors: computed_errors.freeze, best_fit: @best_fit, collapsed: @collapsed })
      end

      # Returns whether the column definition is valid according to OOXML specifications.
      #
      # @return [Boolean]
      #: () -> bool
      def valid?
        errors.empty?
      end

      # Returns whether the column is hidden.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def hidden?
        @hidden
      end

      # Returns whether custom width is set.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def custom_width?
        @custom_width
      end

      # Returns whether best fit is enabled.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def best_fit?
        @best_fit
      end

      # Returns whether the column is collapsed.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def collapsed?
        @collapsed
      end

      # Deconstructs column attributes for pattern matching.
      #
      # @param _keys [Array<Symbol>, nil]
      # @return [Hash{Symbol => Object}]
      # @api public
      #: (Array[Symbol]? _keys) -> Hash[Symbol, untyped]
      def deconstruct_keys(_keys)
        {
          index: index,
          width: width,
          hidden: hidden?,
          custom_width: custom_width?,
          best_fit: best_fit?,
          collapsed: collapsed?,
          outline_level: outline_level,
          style_index: style_index
        }
      end

      # Returns the column letter (e.g. "A", "B", "AA").
      #
      # @return [String]
      # @api public
      #: () -> String
      def column_letter
        Cell._column_letter(index)
      end

      # Returns the style identifier or inline style.
      #
      # @return [Integer, String, Symbol, Hash, nil]
      # @api public
      #: () -> (Integer | String | Symbol | Hash[Symbol, untyped])?
      def style
        style_index
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
      # @api private
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
