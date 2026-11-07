# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Elements
    # Represents a single cell in a worksheet.
    # All row and column indices are 0-based.
    #
    # @example Access cell properties
    #   cell = sheet["A1"]
    #   cell.value       # raw value
    #   cell.ref         # "A1"
    #   cell.to_i        # integer value
    #   cell.to_date     # Date value
    #
    # @api public
    class Cell
      attr_reader :row_index, :column_index, :value, :formula, :style, :style_index, :unmapped_data, :errors, :hyperlink, :comment, :raw_value, :format_code, :date1904, :threaded_comments

      # @param row_index [Integer] 0-based row index.
      # @param column_index [Integer] 0-based column index.
      # @param value [Object, nil] The cell's value.
      # @param formula [Elements::Formula, nil] Optional formula.
      # @param style [String, Symbol, Hash, nil] Style name or inline style properties.
      # @param style_index [Integer, nil] 0-based style index.
      # @param unmapped_data [Hash] Additional metadata.
      # @param errors [Array<String>, nil] Validation errors.
      # @param hyperlink [Hash, String, Elements::Hyperlink, nil] Optional hyperlink metadata or URL.
      # @param comment [Hash, String, nil] Optional cell comment or note.
      # @param raw_value [String, nil] Optional unparsed raw string value.
      # @param format_code [String, nil] Optional OpenXML number/date format code.
      # @param date1904 [Boolean] Whether the cell belongs to a 1904 date system workbook.
      # @param threaded_comments [Array<Hash>, nil] Optional threaded comments.
      #: (row_index: untyped, column_index: untyped, ?value: untyped, ?formula: (Elements::Formula | String)?, ?style: (String | Symbol | Hash[Symbol, untyped])?, ?style_index: (Integer | String | Symbol | Hash[Symbol, untyped])?, ?unmapped_data: Hash[untyped, untyped], ?errors: Array[String]?, ?hyperlink: (Hash[Symbol, untyped] | String | Elements::Hyperlink)?, ?comment: (Hash[Symbol, untyped] | String)?, ?raw_value: String?, ?format_code: String?, ?date1904: bool, ?threaded_comments: Array[Hash[Symbol, untyped]]?) -> void
      def initialize(row_index:, column_index:, value: nil, formula: nil, style: nil, style_index: nil, unmapped_data: EMPTY_HASH, errors: nil, hyperlink: nil, comment: nil, raw_value: nil, format_code: nil, date1904: false, threaded_comments: nil)
        if style_index && !style_index.is_a?(Integer)
          warn "[DEPRECATION] Elements::Cell.new passing non-Integer to style_index: is deprecated; use style: instead.", uplevel: 1
          style ||= style_index
          style_index = nil
        end
        @row_index = row_index
        @column_index = column_index
        @value = value
        @formula = formula
        @style = style || (style_index unless style_index.is_a?(Integer))
        @style_index = style_index || (style if style.is_a?(Integer))
        @unmapped_data = unmapped_data || EMPTY_HASH
        computed_errors = errors || self.class.validate(row_index, column_index, value)
        @errors = computed_errors.frozen? ? computed_errors : computed_errors.freeze
        @hyperlink = hyperlink
        @comment = comment
        @raw_value = raw_value
        @format_code = format_code
        @date1904 = date1904 ? true : false
        @threaded_comments = (threaded_comments || EMPTY_ARRAY).freeze
      end

      # rubocop:disable Style/OptionalBooleanParameter
      def self.fast_create(row_index, column_index, value, style_index = nil, formula = nil, raw_value = nil, format_code = nil, date1904 = false)
        inst = allocate
        inst.instance_variable_set(:@row_index, row_index)
        inst.instance_variable_set(:@column_index, column_index)
        inst.instance_variable_set(:@value, value)
        inst.instance_variable_set(:@formula, formula)
        inst.instance_variable_set(:@style, (style_index.is_a?(Integer) ? nil : style_index))
        inst.instance_variable_set(:@style_index, (style_index.is_a?(Integer) ? style_index : nil))
        inst.instance_variable_set(:@unmapped_data, EMPTY_HASH)
        inst.instance_variable_set(:@errors, EMPTY_ERRORS)
        inst.instance_variable_set(:@hyperlink, nil)
        inst.instance_variable_set(:@comment, nil)
        inst.instance_variable_set(:@raw_value, raw_value)
        inst.instance_variable_set(:@format_code, format_code)
        inst.instance_variable_set(:@date1904, date1904 ? true : false)
        inst.instance_variable_set(:@threaded_comments, EMPTY_ARRAY)
        inst
      end
      # rubocop:enable Style/OptionalBooleanParameter

      def ==(other)
        other.is_a?(Cell) &&
          row_index == other.row_index &&
          column_index == other.column_index &&
          value == other.value &&
          formula == other.formula &&
          style == other.style &&
          style_index == other.style_index &&
          unmapped_data == other.unmapped_data &&
          errors == other.errors &&
          hyperlink == other.hyperlink &&
          comment == other.comment &&
          raw_value == other.raw_value &&
          format_code == other.format_code &&
          threaded_comments == other.threaded_comments
      end
      alias eql? ==

      def hash
        [row_index, column_index, value, formula, style, style_index, unmapped_data, errors, hyperlink, comment, raw_value, format_code, threaded_comments].hash
      end

      def deconstruct
        [row_index, column_index, value, formula, (style_index || style), unmapped_data, errors, hyperlink, comment, raw_value, format_code, threaded_comments]
      end

      def deconstruct_keys(keys)
        h = { row_index: row_index, column_index: column_index, value: value, formula: formula,
              style_index: style_index, unmapped_data: unmapped_data, errors: errors, hyperlink: hyperlink,
              comment: comment, raw_value: raw_value, format_code: format_code, threaded_comments: threaded_comments }
        h[:style] = style if style
        keys ? h.slice(*keys) : h
      end

      def with(**changes)
        self.class.new(
          row_index: changes.fetch(:row_index, row_index),
          column_index: changes.fetch(:column_index, column_index),
          value: changes.fetch(:value, value),
          formula: changes.fetch(:formula, formula),
          style: changes.fetch(:style, style),
          style_index: changes.fetch(:style_index, style_index),
          unmapped_data: changes.fetch(:unmapped_data, unmapped_data),
          errors: changes.fetch(:errors, errors),
          hyperlink: changes.fetch(:hyperlink, hyperlink),
          comment: changes.fetch(:comment, comment),
          raw_value: changes.fetch(:raw_value, raw_value),
          format_code: changes.fetch(:format_code, format_code),
          date1904: changes.fetch(:date1904, @date1904 || false),
          threaded_comments: changes.fetch(:threaded_comments, threaded_comments)
        )
      end

      # Returns whether the cell is valid according to OOXML specifications.
      #
      # @return [Boolean]
      #: () -> bool
      def valid?
        errors.empty?
      end

      # Returns the Excel-style reference (e.g. "A1", "B2").
      #
      # @return [String]
      # @api public
      #: () -> String
      def ref
        "#{self.class._column_letter(column_index)}#{row_index + 1}"
      end

      # Returns whether the cell contains a formula.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def formula?
        if formula.is_a?(Formula)
          !formula.expression.nil? && !formula.expression.empty?
        else
          !formula.nil? && !formula.to_s.empty?
        end
      end

      # Returns the formula expression string without leading '=', or nil.
      #
      # @return [String, nil]
      # @api public
      #: () -> String?
      def formula_expression
        expr = formula.is_a?(Formula) ? formula.expression : formula
        return nil if expr.nil?

        str = expr.to_s
        str.empty? ? nil : str
      end

      # Returns whether the cell contains a hyperlink.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def link?
        case hyperlink
        when nil then false
        when Hash, String then !hyperlink.empty?
        else true
        end
      end
      alias hyperlink? link?

      # Returns the hyperlink target URL or location string, or nil.
      #
      # @return [String, nil]
      # @api public
      #: () -> String?
      def url
        case hyperlink
        when Hash
          u = hyperlink[:url] || hyperlink[:location]
          u&.to_s
        when String
          hyperlink.empty? ? nil : hyperlink
        when Elements::Hyperlink
          (hyperlink.url || hyperlink.location)&.to_s
        end
      end

      # Returns the hyperlink tooltip (ScreenTip) text, or nil.
      #
      # @return [String, nil]
      # @api public
      #: () -> String?
      def tooltip
        case hyperlink
        when Hash
          hyperlink[:tooltip]&.to_s
        when Elements::Hyperlink
          hyperlink.tooltip&.to_s
        end
      end

      # Returns whether the cell contains a comment or note.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def comment?
        case comment
        when nil then false
        when Hash, String then !comment.empty?
        else true
        end
      end

      # Returns the plain text of the cell comment or note, or nil.
      #
      # @return [String, nil]
      # @api public
      #: () -> String?
      def comment_text
        case comment
        when Hash
          txt = comment[:text]
          txt&.to_s
        when String
          comment.empty? ? nil : comment
        end
      end

      # Returns whether the cell contains any threaded comments.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def threaded_comments?
        !@threaded_comments.empty?
      end

      # Returns the root threaded comment for this cell, or nil.
      #
      # @return [Hash[Symbol, untyped], nil]
      # @api public
      #: () -> Hash[Symbol, untyped]?
      def threaded_comment
        @threaded_comments.first
      end

      # Returns the formatted string representation of the cell value according to its format code.
      #
      # @param custom_format_code [String, nil] Optional format code override.
      # @return [String, nil]
      # @api public
      #: (?String? custom_format_code) -> String?
      def formatted_value(custom_format_code = nil)
        return nil if value.nil?

        NumberFormatter.format(value, custom_format_code || format_code)
      end

      # Access cell attributes by Symbol key.
      #
      # @param key [Symbol] Attribute key (:value, :formula, :style_index, :ref, :column_index, :row_index, :type, :raw_value, :format_code, :formatted_value).
      # @return [Object, nil]
      # @api public
      #: (Symbol key) -> untyped
      def [](key)
        case key
        when :value then value
        when :formula
          formula.is_a?(Formula) ? formula.expression : formula
        when :style_index then style_index
        when :style then style || style_index
        when :ref then ref
        when :column_index then column_index
        when :row_index then row_index
        when :hyperlink then hyperlink
        when :url then url
        when :tooltip then tooltip
        when :comment then comment
        when :threaded_comments then threaded_comments
        when :threaded_comment then threaded_comment
        when :raw_value then raw_value
        when :format_code then format_code
        when :formatted_value then formatted_value
        when :format_type then format_type
        when :date_format then date_format?
        when :time_format then time_format?
        when :date1904 then date1904?
        when :type
          case value
          when String then "s"
          when true, false then "b"
          end
        end
      end

      # Returns the classified format category for this cell (:date, :datetime, :time, :number, :text, :general, or nil).
      #
      # @return [Symbol, nil]
      # @api public
      #: () -> Symbol?
      def format_type
        format_code ? NumberFormatter.format_type(format_code) : nil
      end

      # Returns whether the cell has a date, datetime, or time format.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def date_format?
        format_code ? NumberFormatter.date_format?(format_code) : false
      end

      # Returns whether the cell has a time-only format.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def time_format?
        format_code ? NumberFormatter.time_format?(format_code) : false
      end

      # Returns whether the cell's workbook uses the 1904 date system.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def date1904?
        @date1904 ? true : false
      end

      # Returns the cell value.
      #
      # @return [Object, nil]
      # @api public
      #: () -> untyped
      def content
        value
      end

      # Returns the string representation of the cell value.
      #
      # @return [String]
      # @api public
      #: () -> String
      def to_s
        value.to_s
      end

      # Converts the cell value to Integer.
      #
      # @return [Integer]
      # @api public
      #: () -> Integer
      def to_i
        value.to_i
      end

      # Converts the cell value to Float.
      #
      # @return [Float]
      # @api public
      #: () -> Float
      def to_f
        value.to_f
      end

      # Converts the cell value (numeric serial date or date string) to Date.
      #
      # @param date1904 [Boolean, nil] Whether to use the 1904 date system (defaults to cell's date1904 setting).
      # @return [Date, nil]
      # @api public
      #: (?date1904: bool?) -> Date?
      def to_date(date1904: nil)
        return value if value.is_a?(Date)

        d1904 = date1904.nil? ? date1904? : date1904
        if value.is_a?(Numeric)
          Ooxml::Utils.serial_to_date(value, date1904: d1904)
        else
          begin
            Date.parse(value.to_s)
          rescue StandardError
            nil
          end
        end
      end

      # Converts the cell value (numeric serial datetime or datetime string) to Time.
      #
      # @param date1904 [Boolean, nil] Whether to use the 1904 date system (defaults to cell's date1904 setting).
      # @return [Time, nil]
      # @api public
      #: (?date1904: bool?) -> Time?
      def to_time(date1904: nil)
        return value if value.is_a?(Time)

        d1904 = date1904.nil? ? date1904? : date1904
        if value.is_a?(Numeric)
          Ooxml::Utils.serial_to_datetime(value, date1904: d1904)
        else
          begin
            Time.parse(value.to_s)
          rescue StandardError
            nil
          end
        end
      end

      # Pure calculation: converts a 0-based column index to an Excel letter.
      #
      # @param index [Integer] 0-based column index.
      # @return [String] Excel column letter.
      # @api private
      #: (Integer index) -> String
      def self.calculate_column_letter(index)
        result = +""
        i = index
        loop do
          result.prepend(("A".ord + (i % 26)).chr)
          i = (i / 26) - 1
          break if i.negative?
        end
        result
      end

      # Cache column letters up to Excel's limit (16,384)
      @column_letters = (0...16_384).map { |index| calculate_column_letter(index).freeze }.freeze

      # @api private
      def self._column_letter(index)
        raise ArgumentError, "Column index must be a non-negative Integer, got #{index.inspect}" unless index.is_a?(Integer) && index >= 0

        @column_letters[index] || calculate_column_letter(index)
      end

      # Converts a 0-based column index to an Excel letter (0 -> "A", 25 -> "Z", 26 -> "AA").
      #
      # @example
      #   Cell.column_letter(0)  #=> "A"
      #   Cell.column_letter(26) #=> "AA"
      #
      # @deprecated Use {Xlsxrb.column_letter} instead. Will be removed in 0.2.0.
      # @param index [Integer] 0-based column index.
      # @return [String] Excel column letter.
      #: (untyped index) -> String
      def self.column_letter(index)
        warn "[DEPRECATION] Xlsxrb::Elements::Cell.column_letter is deprecated; use Xlsxrb.column_letter instead.", uplevel: 1
        _column_letter(index)
      end

      # Pure calculation: converts an uppercase column letter string to a 0-based index.
      #
      # @param letter [String] Uppercase column letter string (e.g. "A", "AA").
      # @return [Integer] 0-based column index.
      # @api private
      #: (String letter) -> Integer
      def self.calculate_column_index(letter)
        letter.bytes.reduce(0) { |acc, b| (acc * 26) + (b - 64) } - 1
      end

      # @api private
      def self._column_index(letter)
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

        calculate_column_index(str.upcase)
      end

      # Converts a column letter (e.g. "A", :AA) to a 0-based column index.
      #
      # @example
      #   Cell.column_index("A")  #=> 0
      #   Cell.column_index(:AA)  #=> 26
      #
      # @deprecated Use {Xlsxrb.column_index} instead. Will be removed in 0.2.0.
      # @param letter [String, Symbol, Integer] Column letter or integer index.
      # @return [Integer] 0-based column index.
      #: (String | Symbol | Integer letter) -> Integer
      def self.column_index(letter)
        warn "[DEPRECATION] Xlsxrb::Elements::Cell.column_index is deprecated; use Xlsxrb.column_index instead.", uplevel: 1
        _column_index(letter)
      end

      # @api private
      def self._parse_ref(ref)
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

        row_part = bytes.byteslice(i, len - i)
        return nil unless row_part.match?(/\A[1-9]\d*\z/)

        row = row_part.to_i - 1
        [row, col - 1]
      end

      # Parses an Excel-style reference to [row_index, col_index] (both 0-based).
      #
      # @example
      #   Cell.parse_ref("A1")  #=> [0, 0]
      #   Cell.parse_ref("C10") #=> [9, 2]
      #
      # @deprecated Use {Xlsxrb.parse_ref} instead. Will be removed in 0.2.0.
      # @param ref [String, nil] Excel cell reference.
      # @return [Array(Integer, Integer), nil] 0-based [row_index, col_index].
      #: (String? ref) -> [Integer, Integer]?
      def self.parse_ref(ref)
        warn "[DEPRECATION] Xlsxrb::Elements::Cell.parse_ref is deprecated; use Xlsxrb.parse_ref instead.", uplevel: 1
        _parse_ref(ref)
      end

      # Returns whether the coordinates are within OOXML limits (row: 0..1,048,575, col: 0..16,383).
      #
      # @param row_index [Integer] 0-based row index.
      # @param column_index [Integer] 0-based column index.
      # @return [Boolean]
      #: (untyped row_index, untyped column_index) -> bool
      def self.valid_coordinates?(row_index, column_index)
        row_index.instance_of?(Integer) && row_index >= 0 && row_index < 1_048_576 &&
          column_index.instance_of?(Integer) && column_index >= 0 && column_index < 16_384
      end

      # Returns whether the value type is supported by OOXML specifications.
      #
      # @param value [Object]
      # @return [Boolean]
      #: (untyped value) -> bool
      def self.valid_value?(value)
        case value
        when nil, String, Numeric, true, false, Date, Time, Formula, RichText, CellError
          true
        when Hash
          value.key?(:formula)
        else
          false
        end
      end

      # Validates cell coordinates and value type against OOXML specifications.
      #
      # @param row_index [Integer]
      # @param column_index [Integer]
      # @param value [Object]
      # @return [Array<String>] List of errors.
      #: (untyped row_index, untyped column_index, untyped value) -> Array[String]
      def self.validate(row_index, column_index, value)
        valid_val = valid_value?(value)
        return EMPTY_ERRORS if valid_coordinates?(row_index, column_index) && valid_val

        errs = []
        errs << "row_index must be a non-negative Integer (got #{row_index.inspect})" if !row_index.is_a?(Integer) || row_index.negative?
        errs << "column_index must be a non-negative Integer (got #{column_index.inspect})" if !column_index.is_a?(Integer) || column_index.negative?
        errs << "row_index must be < 1048576 (got #{row_index}, max row is 1048575)" if row_index.is_a?(Integer) && row_index >= 1_048_576
        errs << "column_index must be < 16384 (got #{column_index}, max column is XFD=16383)" if column_index.is_a?(Integer) && column_index >= 16_384
        errs << "unsupported value type: #{value.class} (#{value.inspect}) — supported types: String, Numeric, true/false, Date, Time, Formula, RichText, CellError, or nil" unless valid_val
        errs
      end
    end
  end
end
