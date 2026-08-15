# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Elements
    EMPTY_ERRORS = [].freeze
    EMPTY_HASH = {}.freeze
    EMPTY_CELLS = [].freeze

    # Represents an Excel formula with an optional cached value and calculation properties.
    #
    # @example Create a formula
    #   formula = Elements::Formula.new(expression: "SUM(A1:A10)")
    #
    # @api public
    Formula = Data.define(:expression, :cached_value, :type, :ref, :shared_index, :calculate_always, :aca, :bx, :dt2d, :dtr, :r1, :r2) do
      # @param expression [String] Excel formula expression without leading '=' (e.g. "SUM(A1:A10)").
      # @param cached_value [Object, nil] Optional precomputed value.
      # @param type [Symbol, String, nil] Formula type (:shared, :array, etc.).
      # @param ref [String, nil] Target cell or range reference.
      # @param shared_index [Integer, nil] Shared formula index.
      # @param calculate_always [Boolean, nil] Force Excel to recalculate on open.
      # @param aca [Boolean, nil] Always calculate array attribute.
      # @param bx [Boolean, nil] Assigns to array formula.
      # @param dt2d [Boolean, nil] 2D data table reference.
      # @param dtr [Boolean, nil] 1D data table reference.
      # @param r1 [String, nil] First table reference.
      # @param r2 [String, nil] Second table reference.
      #: (expression: String, ?cached_value: untyped, ?type: untyped, ?ref: String?, ?shared_index: Integer?, ?calculate_always: bool?, ?aca: bool?, ?bx: bool?, ?dt2d: bool?, ?dtr: bool?, ?r1: String?, ?r2: String?) -> void
      def initialize(expression:, cached_value: nil, type: nil, ref: nil, shared_index: nil, calculate_always: nil, aca: nil, bx: nil, dt2d: nil, dtr: nil, r1: nil, r2: nil) # rubocop:disable Naming/MethodParameterName
        super
      end
    end

    # Represents a cell error value (e.g. #N/A, #REF!, #DIV/0!).
    # @api public
    VALID_ERROR_CODES = %w[#NULL! #DIV/0! #VALUE! #REF! #NAME? #NUM! #N/A #GETTING_DATA].freeze
    CellError = Data.define(:code) do
      # @param code [String] Valid error code string (e.g. "#N/A").
      #: (code: String) -> void
      def initialize(code:)
        raise ArgumentError, "invalid error code: #{code.inspect} (must be one of #{VALID_ERROR_CODES.join(", ")})" unless VALID_ERROR_CODES.include?(code)

        super
      end

      # @return [String]
      # @api public
      #: () -> String
      def to_s
        code
      end
    end

    # Represents a rich text string with multiple formatting runs.
    #
    # @example
    #   rt = Elements::RichText.new(runs: [{ text: "Hello ", font: { bold: true } }, { text: "World" }])
    #
    # @api public
    RichText = Data.define(:runs) do
      # Returns concatenated plain text of all runs.
      #
      # @return [String]
      # @api public
      #: () -> String
      def to_s
        runs.map { |r| r[:text] }.join
      end
    end
  end
end
