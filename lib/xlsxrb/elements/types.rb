# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Elements
    EMPTY_ERRORS = [].freeze
    EMPTY_HASH = {}.freeze
    EMPTY_CELLS = [].freeze
    EMPTY_ARRAY = [].freeze

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

      # Returns the formula expression string.
      #
      # @return [String]
      # @api public
      #: () -> String
      def to_s
        expression
      end

      # Returns whether Excel is forced to recalculate on open.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def calculate_always?
        calculate_always ? true : false
      end

      # Returns whether the formula is shared.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def shared?
        [:shared, "shared"].include?(type)
      end

      # Returns whether the formula is an array formula.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def array?
        [:array, "array"].include?(type)
      end

      # Returns whether the formula is a data table.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def data_table?
        [:data_table, "data_table", :dataTable, "dataTable"].include?(type)
      end

      # Returns whether the always calculate array flag is set.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def always_calculate_array?
        aca ? true : false
      end

      # rubocop:disable Naming/PredicateMethod

      # @deprecated Use {#always_calculate_array?} instead.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def always_calculate_array
        warn "[Xlsxrb DEPRECATION] Formula#always_calculate_array is deprecated; use #always_calculate_array? instead.", uplevel: 1
        always_calculate_array?
      end

      # Returns whether the formula is assigned to a defined name.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def assign_to_name?
        bx ? true : false
      end

      # @deprecated Use {#assign_to_name?} instead.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def assign_to_name
        warn "[Xlsxrb DEPRECATION] Formula#assign_to_name is deprecated; use #assign_to_name? instead.", uplevel: 1
        assign_to_name?
      end

      # Returns whether this is a 2D data table reference.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def data_table_2d?
        dt2d ? true : false
      end

      # @deprecated Use {#data_table_2d?} instead.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def data_table_2d
        warn "[Xlsxrb DEPRECATION] Formula#data_table_2d is deprecated; use #data_table_2d? instead.", uplevel: 1
        data_table_2d?
      end

      # Returns whether this is a 1D data table row reference.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def data_table_row?
        dtr ? true : false
      end

      # @deprecated Use {#data_table_row?} instead.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def data_table_row
        warn "[Xlsxrb DEPRECATION] Formula#data_table_row is deprecated; use #data_table_row? instead.", uplevel: 1
        data_table_row?
      end

      # rubocop:enable Naming/PredicateMethod
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

    class CellError
      # Predefined frozen CellError constants
      #: Elements::CellError
      NULL = new(code: "#NULL!").freeze
      #: Elements::CellError
      DIV0 = new(code: "#DIV/0!").freeze
      #: Elements::CellError
      VALUE = new(code: "#VALUE!").freeze
      #: Elements::CellError
      REF = new(code: "#REF!").freeze
      #: Elements::CellError
      NAME = new(code: "#NAME?").freeze
      #: Elements::CellError
      NUM = new(code: "#NUM!").freeze
      #: Elements::CellError
      NA = new(code: "#N/A").freeze
      #: Elements::CellError
      GETTING_DATA = new(code: "#GETTING_DATA").freeze
    end

    # Represents a formatted text run within a rich text string.
    #
    # @example
    #   run = Elements::RichTextRun.new(text: "Hello", font: { bold: true })
    #
    # @api public
    RichTextRun = Data.define(:text, :font) do
      # @param text [String] Plain text for this run.
      # @param font [Hash, nil] Optional font styling properties.
      #: (?text: String, ?font: Hash[Symbol, untyped]?) -> void
      def initialize(text: "", font: nil)
        super
      end

      # Supports hash-like key access for backwards compatibility.
      #
      # @deprecated Use {#text} or {#font} instead.
      # @param key [Symbol]
      # @return [Object, nil]
      # @api public
      #: (Symbol key) -> untyped
      def [](key)
        warn "[Xlsxrb DEPRECATION] RichTextRun#[] is deprecated; use #text or #font instead.", uplevel: 1
        case key
        when :text then text
        when :font then font
        end
      end

      # Returns the font size in points, if specified.
      #
      # @return [Numeric, nil]
      # @api public
      #: () -> Numeric?
      def font_size
        return nil unless font

        sz = font[:size] || font[:sz]
        sz.is_a?(Numeric) ? sz : nil
      end

      # @deprecated Use {#font_size} instead. Will be removed in 0.2.0.
      #: () -> Numeric?
      def sz
        warn "[DEPRECATION] `sz` is deprecated; use `font_size` instead.", uplevel: 1
        font_size
      end

      # Returns the plain text of this run.
      #
      # @return [String]
      # @api public
      #: () -> String
      def to_s
        text
      end
    end

    # Represents a rich text string with multiple formatting runs.
    #
    # @example
    #   rt = Elements::RichText.new(runs: [{ text: "Hello ", font: { bold: true } }, { text: "World" }])
    #
    # @api public
    RichText = Data.define(:runs) do
      [Enumerable].each { |m| include m }

      # Iterates over each text run in this rich text string.
      #
      # @yield [run]
      # @yieldparam run [Elements::RichTextRun, Hash]
      # @return [Enumerator, void]
      # @api public
      #: () { (Elements::RichTextRun | Hash[Symbol, untyped]) -> void } -> void
      #: () -> Enumerator[Elements::RichTextRun | Hash[Symbol, untyped], void]
      def each(&block)
        return enum_for(:each) unless block

        runs.each(&block)
      end

      # Accesses a text run by index.
      #
      # @param index [Integer]
      # @return [Elements::RichTextRun, Hash, nil]
      # @api public
      #: (Integer index) -> (Elements::RichTextRun | Hash[Symbol, untyped])?
      def [](index)
        runs[index]
      end

      # Returns the number of text runs.
      #
      # @return [Integer]
      # @api public
      #: () -> Integer
      def size
        runs.size
      end

      # Returns whether this rich text contains no text runs.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def empty?
        runs.empty?
      end

      # Returns concatenated plain text of all runs.
      #
      # @return [String]
      # @api public
      #: () -> String
      def to_s
        runs.map { |r| r.is_a?(Hash) ? r[:text] : r.to_s }.join
      end
    end

    # Represents a worksheet hyperlink with an optional target URL, display text, tooltip (ScreenTip), and internal location.
    #
    # @example Create a hyperlink
    #   link = Elements::Hyperlink.new(url: "https://example.com", tooltip: "Visit website")
    #
    # @api public
    # rubocop:disable Lint/DataDefineOverride -- :display represents the hyperlink display text attribute in OOXML
    Hyperlink = Data.define(:url, :display, :tooltip, :location) do
      # rubocop:enable Lint/DataDefineOverride
      # @param url [String, nil] Target external URL or URI.
      # @param display [String, nil] Optional friendly display label.
      # @param tooltip [String, nil] Optional hover ScreenTip text.
      # @param location [String, nil] Optional internal sheet/cell reference.
      #: (?url: String?, ?display: String?, ?tooltip: String?, ?location: String?, **untyped) -> void
      def initialize(url: nil, display: nil, tooltip: nil, location: nil, **_extra)
        super(url: url, display: display, tooltip: tooltip, location: location)
      end

      # Supports hash-like key access for backwards compatibility.
      #
      # @deprecated Hash-like access via {#[]} is deprecated and will be removed in 0.2.0; use method calls instead.
      # @param key [Symbol, String]
      # @return [Object, nil]
      # @api public
      #: (Symbol | String key) -> untyped
      def [](key)
        warn "[DEPRECATION] Hash access on Elements::Hyperlink (`#[]`) is deprecated and will be removed in 0.2.0; use method calls instead.", uplevel: 1
        k = key.is_a?(Symbol) ? key : key.to_s.to_sym
        case k
        when :url then url
        when :display then display
        when :tooltip then tooltip
        when :location then location
        end
      end

      # Returns whether all attributes are nil or empty.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def empty?
        to_h.empty?
      end

      # Supports Hash conversion.
      #
      # @return [Hash{Symbol => String}]
      # @api public
      #: () -> Hash[Symbol, String]
      def to_h
        h = {}
        h[:url] = url if url
        h[:display] = display if display
        h[:tooltip] = tooltip if tooltip
        h[:location] = location if location
        h
      end

      alias_method :to_hash, :to_h

      # Checks equality against another instance or a Hash.
      #
      # @param other [Object]
      # @return [Boolean]
      # @api public
      #: (untyped other) -> bool
      def ==(other)
        if other.is_a?(Hash)
          warn "[Xlsxrb DEPRECATION] Comparing Hyperlink with Hash is deprecated; compare with Hyperlink instead.", uplevel: 1
          to_h == other.transform_keys(&:to_sym).compact
        else
          super
        end
      end

      # Returns the target URL or location string.
      #
      # @return [String, nil]
      # @api public
      #: () -> String?
      def to_s
        url || location
      end
    end
  end
end
