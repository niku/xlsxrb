# frozen_string_literal: true

module Xlsxrb
  module Elements
    # Represents a formula with an optional cached value.
    # Optional: type (:shared, :array), ref (range), shared_index (si for shared formulas)
    Formula = Data.define(:expression, :cached_value, :type, :ref, :shared_index, :calculate_always, :aca, :bx, :dt2d, :dtr, :r1, :r2) do
      def initialize(expression:, cached_value: nil, type: nil, ref: nil, shared_index: nil, calculate_always: nil, aca: nil, bx: nil, dt2d: nil, dtr: nil, r1: nil, r2: nil) # rubocop:disable Naming/MethodParameterName
        super
      end
    end

    # Represents a cell error value (e.g. #N/A, #REF!, #DIV/0!).
    VALID_ERROR_CODES = %w[#NULL! #DIV/0! #VALUE! #REF! #NAME? #NUM! #N/A #GETTING_DATA].freeze
    CellError = Data.define(:code) do
      def initialize(code:)
        raise ArgumentError, "invalid error code: #{code.inspect} (must be one of #{VALID_ERROR_CODES.join(", ")})" unless VALID_ERROR_CODES.include?(code)

        super
      end

      def to_s
        code
      end
    end

    # Represents a rich text string with formatting runs.
    # runs: array of hashes, each with :text and optional :font (hash of font properties).
    # Font properties: :bold, :italic, :underline, :sz, :color, :name
    RichText = Data.define(:runs) do
      def to_s
        runs.map { |r| r[:text] }.join
      end
    end
  end
end
