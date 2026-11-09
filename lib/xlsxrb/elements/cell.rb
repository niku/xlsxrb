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
      #: Integer
      attr_reader :row_index
      #: Integer
      attr_reader :column_index
      #: untyped
      attr_reader :value
      #: Elements::Formula?
      attr_reader :formula
      #: (String | Symbol | Hash[Symbol, untyped])?
      attr_reader :style
      #: Integer?
      attr_reader :style_index
      #: Hash[untyped, untyped]
      attr_reader :unmapped_data
      #: Array[String]
      attr_reader :errors
      #: Elements::Hyperlink?
      attr_reader :hyperlink
      #: Elements::Comment?
      attr_reader :comment
      #: String?
      attr_reader :raw_value
      #: String?
      attr_reader :format_code
      #: Array[Elements::ThreadedComment]
      attr_reader :threaded_comments

      # @param row_index [Integer] 0-based row index.
      # @param column_index [Integer] 0-based column index.
      # @param value [Object, nil] The cell's value.
      # @param formula [Elements::Formula, nil] Optional formula.
      # @param style [String, Symbol, Hash, nil] Style name or inline style properties.
      # @param style_index [Integer, nil] 0-based style index.
      # @param unmapped_data [Hash] Additional metadata.
      # @param errors [Array<String>, nil] Validation errors.
      # @param hyperlink [Elements::Hyperlink, Hash, String, nil] Optional hyperlink metadata or URL.
      # @param comment [Elements::Comment, Hash, String, nil] Optional cell comment or note.
      # @param raw_value [String, nil] Optional unparsed raw string value.
      # @param format_code [String, nil] Optional OpenXML number/date format code.
      # @param date1904 [Boolean] Whether the cell belongs to a 1904 date system workbook.
      # @param threaded_comments [Array<Elements::ThreadedComment, Hash>, nil] Optional threaded comments.
      #: (row_index: untyped, column_index: untyped, ?value: untyped, ?formula: (Elements::Formula | String | Hash[Symbol, untyped])?, ?style: (String | Symbol | Hash[Symbol, untyped])?, ?style_index: Integer?, ?unmapped_data: Hash[untyped, untyped], ?errors: Array[String]?, ?hyperlink: (Elements::Hyperlink | Hash[Symbol, untyped] | String)?, ?comment: (Elements::Comment | Hash[Symbol, untyped] | String)?, ?raw_value: String?, ?format_code: String?, ?date1904: bool, ?threaded_comments: Array[Elements::ThreadedComment | Hash[Symbol, untyped]]?) -> void
      def initialize(row_index:, column_index:, value: nil, formula: nil, style: nil, style_index: nil, unmapped_data: EMPTY_HASH, errors: nil, hyperlink: nil, comment: nil, raw_value: nil, format_code: nil, date1904: false, threaded_comments: nil)
        @row_index = row_index
        @column_index = column_index
        @value = value
        @formula = self.class.normalize_formula(formula)
        @style = style
        @style_index = style_index || (style if style.is_a?(Integer))
        @unmapped_data = unmapped_data || EMPTY_HASH
        computed_errors = errors || self.class.validate(row_index, column_index, value)
        @errors = computed_errors.frozen? ? computed_errors : computed_errors.freeze
        @hyperlink = case hyperlink
                     when Hash
                       Elements::Hyperlink.new(**hyperlink)
                     when String
                       hyperlink.empty? ? nil : Elements::Hyperlink.new(url: hyperlink)
                     else
                       hyperlink
                     end
        @comment = case comment
                   when nil, Elements::Comment
                     comment
                   when String
                     comment.empty? ? nil : Elements::Comment.new(text: comment)
                   when Hash
                     Elements::Comment.new(**comment)
                   else
                     Elements::Comment.new(text: comment.to_s)
                   end
        @raw_value = raw_value
        @format_code = format_code
        @date1904 = date1904 ? true : false
        @threaded_comments = if threaded_comments
                               threaded_comments.map do |tc|
                                 tc.is_a?(Elements::ThreadedComment) ? tc : Elements::ThreadedComment.new(**tc)
                               end.freeze
                             else
                               EMPTY_ARRAY
                             end
        freeze
      end

      # @api private
      # rubocop:disable Style/OptionalBooleanParameter
      def self.fast_create(row_index, column_index, value, style_index = nil, formula = nil, raw_value = nil, format_code = nil, date1904 = false)
        allocate.__send__(:init_fast, row_index, column_index, value, style_index, formula, raw_value, format_code, date1904)
      end
      # rubocop:enable Style/OptionalBooleanParameter

      # Assigns instance variables directly, in the same order as #initialize, for
      # fast_create. One method call instead of one instance_variable_set per variable.
      #: (untyped row_index, untyped column_index, untyped value, untyped style_index, untyped formula, String? raw_value, String? format_code, bool date1904) -> Cell
      def init_fast(row_index, column_index, value, style_index, formula, raw_value, format_code, date1904)
        integer_style = style_index.is_a?(Integer)
        @row_index = row_index
        @column_index = column_index
        @value = value
        @formula = formula.nil? ? nil : self.class.normalize_formula(formula)
        @style = integer_style ? nil : style_index
        @style_index = integer_style ? style_index : nil
        @unmapped_data = EMPTY_HASH
        @errors = EMPTY_ERRORS
        @hyperlink = nil
        @comment = nil
        @raw_value = raw_value
        @format_code = format_code
        @date1904 = date1904 ? true : false
        @threaded_comments = EMPTY_ARRAY
        freeze
      end
      private :init_fast

      # @api private
      #: (untyped formula) -> Elements::Formula?
      def self.normalize_formula(formula)
        case formula
        when nil, Elements::Formula
          formula
        when String
          expr = formula.delete_prefix("=")
          expr.empty? ? nil : Elements::Formula.new(expression: expr)
        when Hash
          expr = (formula[:expression] || formula[:formula] || "").to_s.delete_prefix("=")
          is_shared = formula[:shared] || formula[:shared_index]
          return nil if expr.empty? && !is_shared

          Elements::Formula.new(
            expression: expr,
            cached_value: formula[:cached_value] || formula[:value],
            type: formula[:type],
            ref: formula[:ref],
            shared_index: formula[:shared_index]&.to_i,
            calculate_always: formula[:calculate_always],
            aca: formula[:aca],
            bx: formula[:bx],
            dt2d: formula[:dt2d],
            dtr: formula[:dtr],
            r1: formula[:r1],
            r2: formula[:r2]
          )
        else
          expr = formula.to_s.delete_prefix("=")
          expr.empty? ? nil : Elements::Formula.new(expression: expr)
        end
      end

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
        [row_index, column_index, value, formula, style_index || style, unmapped_data, errors, hyperlink, comment, raw_value, format_code, threaded_comments]
      end

      # Deconstructs cell attributes for pattern matching.
      #
      # @param _keys [Array<Symbol>, nil]
      # @return [Hash{Symbol => Object}]
      # @api public
      #: (Array[Symbol]? _keys) -> Hash[Symbol, untyped]
      def deconstruct_keys(_keys)
        {
          ref: ref,
          row_index: @row_index,
          column_index: @column_index,
          value: @value,
          formula: @formula,
          style_index: @style_index,
          format_code: @format_code,
          hyperlink: @hyperlink,
          comment: @comment,
          threaded_comments: @threaded_comments
        }
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
      def hyperlink?
        case hyperlink
        when nil then false
        when Hash, String, Elements::Hyperlink then !hyperlink.empty?
        else true
        end
      end

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
        when Hash, String, Elements::Comment then !comment.empty?
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
        when Elements::Comment
          comment.text&.to_s
        else
          comment.respond_to?(:text) ? comment.text&.to_s : nil
        end
      end

      # Returns whether the cell contains any threaded comments.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def threaded_comment?
        !@threaded_comments.empty?
      end

      # Returns the root threaded comment for this cell, or nil.
      #
      # @return [Elements::ThreadedComment, nil]
      # @api public
      #: () -> Elements::ThreadedComment?
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

      # Returns the classified format category for this cell (:date, :datetime, :time, :number, :text, :general, or nil).
      #
      # @return [Symbol, nil]
      # @api public
      #: () -> Symbol?
      def format_type
        format_code ? NumberFormatter.format_type(format_code) : nil
      end

      # Returns the number format code for this cell.
      #
      # @return [String, nil]
      # @api public
      #: () -> String?
      def number_format
        format_code
      end

      # Returns whether the cell has a date, datetime, or time format.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def date_format?
        format_code ? NumberFormatter.date_format?(format_code) : false
      end

      # Returns whether the cell has a datetime format (date with time).
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def datetime_format?
        format_code ? NumberFormatter.datetime_format?(format_code) : false
      end

      # Returns whether the cell has a date-only format.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def date_only_format?
        format_code ? NumberFormatter.date_only_format?(format_code) : false
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
      # @return [Date, nil]
      # @api public
      #: () -> Date?
      def to_date
        return value if value.is_a?(Date)

        if value.is_a?(Numeric)
          Ooxml::Utils.serial_to_date(value, date1904: date1904?)
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
      # @return [Time, nil]
      # @api public
      #: () -> Time?
      def to_time
        return value if value.is_a?(Time)

        if value.is_a?(Numeric)
          Ooxml::Utils.serial_to_datetime(value, date1904: date1904?)
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
      # @api private
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
