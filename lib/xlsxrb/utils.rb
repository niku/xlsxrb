# frozen_string_literal: true

# rbs_inline: enabled

require_relative "elements/cell"
require_relative "ooxml/utils"

module Xlsxrb
  # Public coordinate conversion and reference parsing utilities.
  #
  # @deprecated Use top-level {Xlsxrb} coordinate and date helpers instead.
  # @api public
  module Utils
    # Converts a cell reference string (e.g. "A1", "BC100") to 0-based [row, col] coordinates.
    #
    # @example
    #   Xlsxrb::Utils.ref_to_row_col("A1")    #=> [0, 0]
    #   Xlsxrb::Utils.ref_to_row_col("BC100") #=> [99, 54]
    #
    # @deprecated Use {Xlsxrb.parse_ref} instead. Will be removed in 0.2.0.
    # @param cell_ref [String, nil] Excel cell reference string.
    # @return [Array(Integer, Integer), nil] 0-based [row_idx, col_idx] tuple, or nil if invalid.
    #: (String? cell_ref) -> [Integer, Integer]?
    def self.ref_to_row_col(cell_ref)
      warn "[DEPRECATION] Xlsxrb::Utils.ref_to_row_col is deprecated; use Xlsxrb.parse_ref instead.", uplevel: 1
      Elements::Cell._parse_ref(cell_ref)
    end

    # Converts 0-based [row_idx, col_idx] coordinates to an A1-style reference string.
    #
    # @example
    #   Xlsxrb::Utils.row_col_to_ref(0, 0)   #=> "A1"
    #   Xlsxrb::Utils.row_col_to_ref(99, 54) #=> "BC100"
    #
    # @deprecated Use {Xlsxrb.cell_to_ref} instead. Will be removed in 0.2.0.
    # @param row_idx [Integer] 0-based row index.
    # @param col_idx [Integer] 0-based column index.
    # @return [String] A1-style reference string.
    #: (Integer row_idx, Integer col_idx) -> String
    def self.row_col_to_ref(row_idx, col_idx)
      warn "[DEPRECATION] Xlsxrb::Utils.row_col_to_ref is deprecated; use Xlsxrb.cell_to_ref instead.", uplevel: 1
      "#{Elements::Cell._column_letter(col_idx)}#{row_idx + 1}"
    end

    # Converts a cell range reference string (e.g. "A1:C10", "B2:D5", "A1") to 0-based
    # [[top_row, left_col], [bottom_row, right_col]] coordinate pairs.
    #
    # @example Range reference
    #   Xlsxrb::Utils.parse_range("A1:C10") #=> [[0, 0], [9, 2]]
    #
    # @example Single cell reference
    #   Xlsxrb::Utils.parse_range("B5")     #=> [[4, 1], [4, 1]]
    #
    # @deprecated Use {Xlsxrb.parse_range} instead. Will be removed in 0.2.0.
    # @param range_ref [String, nil] Cell range reference string.
    # @return [Array([Integer, Integer], [Integer, Integer]), nil] 0-based coordinate pairs, or nil if invalid.
    #: (String? range_ref) -> [[Integer, Integer], [Integer, Integer]]?
    def self.parse_range(range_ref)
      warn "[DEPRECATION] Xlsxrb::Utils.parse_range is deprecated; use Xlsxrb.parse_range instead.", uplevel: 1
      Xlsxrb.parse_range(range_ref)
    end

    # Converts 0-based bounding coordinates to an A1-style range reference string.
    #
    # @example Range coordinates
    #   Xlsxrb::Utils.range_to_ref(0, 0, 9, 2) #=> "A1:C10"
    #
    # @example Single cell coordinates
    #   Xlsxrb::Utils.range_to_ref(4, 1, 4, 1) #=> "B5"
    #
    # @deprecated Use {Xlsxrb.range_to_ref} instead. Will be removed in 0.2.0.
    # @param top_row [Integer] 0-based start row index.
    # @param left_col [Integer] 0-based start column index.
    # @param bottom_row [Integer] 0-based end row index.
    # @param right_col [Integer] 0-based end column index.
    # @return [String] A1-style range reference string.
    #: (Integer top_row, Integer left_col, Integer bottom_row, Integer right_col) -> String
    def self.range_to_ref(top_row, left_col, bottom_row, right_col)
      warn "[DEPRECATION] Xlsxrb::Utils.range_to_ref is deprecated; use Xlsxrb.range_to_ref instead.", uplevel: 1
      Xlsxrb.range_to_ref(top_row, left_col, bottom_row, right_col)
    end

    # Converts a column letter name (e.g. "A", "BC") to a 0-based integer index.
    #
    # @example
    #   Xlsxrb::Utils.col_name_to_index("A")  #=> 0
    #   Xlsxrb::Utils.col_name_to_index("BC") #=> 54
    #
    # @deprecated Use {Xlsxrb.column_index} instead. Will be removed in 0.2.0.
    # @param col_name [String, Symbol, Integer] Column letter or index.
    # @return [Integer] 0-based column index.
    #: (String | Symbol | Integer col_name) -> Integer
    def self.col_name_to_index(col_name)
      warn "[DEPRECATION] Xlsxrb::Utils.col_name_to_index is deprecated; use Xlsxrb.column_index instead.", uplevel: 1
      Elements::Cell._column_index(col_name)
    end

    # Converts a 0-based column integer index to a column letter name (e.g. 0 -> "A", 54 -> "BC").
    #
    # @example
    #   Xlsxrb::Utils.col_index_to_name(0)  #=> "A"
    #   Xlsxrb::Utils.col_index_to_name(54) #=> "BC"
    #
    # @deprecated Use {Xlsxrb.column_letter} instead. Will be removed in 0.2.0.
    # @param col_idx [Integer] 0-based column index.
    # @return [String] Column letter name.
    #: (Integer col_idx) -> String
    def self.col_index_to_name(col_idx)
      warn "[DEPRECATION] Xlsxrb::Utils.col_index_to_name is deprecated; use Xlsxrb.column_letter instead.", uplevel: 1
      Elements::Cell._column_letter(col_idx)
    end

    # Splits an A1-style cell reference into its column name and 1-based row number.
    #
    # @example
    #   Xlsxrb::Utils.split_coordinate("A1")    #=> ["A", 1]
    #   Xlsxrb::Utils.split_coordinate("BC100") #=> ["BC", 100]
    #
    # @deprecated Use {Xlsxrb.split_coordinate} instead. Will be removed in 0.2.0.
    # @param cell_ref [String, nil] Excel cell reference string.
    # @return [Array(String, Integer), nil] [col_name, row_number] tuple (1-based row), or nil if invalid.
    #: (String? cell_ref) -> [String, Integer]?
    def self.split_coordinate(cell_ref)
      warn "[DEPRECATION] Xlsxrb::Utils.split_coordinate is deprecated; use Xlsxrb.split_coordinate instead.", uplevel: 1
      Xlsxrb.split_coordinate(cell_ref)
    end

    # Converts a Date or Time to an Excel numeric serial number (1900 or 1904 system).
    #
    # @example
    #   Xlsxrb::Utils.date_to_serial(Date.new(2026, 1, 1)) #=> 46023
    #
    # @deprecated Use {Xlsxrb.date_to_serial} instead. Will be removed in 0.2.0.
    # @param date_or_time [Date, Time] Date or Time instance.
    # @param date1904 [Boolean] Whether to use the 1904 date system.
    # @return [Integer, Float] Serial number.
    #: (Date | Time date_or_time, ?date1904: bool) -> (Integer | Float)
    def self.date_to_serial(date_or_time, date1904: false)
      warn "[DEPRECATION] Xlsxrb::Utils.date_to_serial is deprecated; use Xlsxrb.date_to_serial instead.", uplevel: 1
      Xlsxrb.date_to_serial(date_or_time, date1904: date1904)
    end

    # Checks whether a serial number is a valid Excel date within the supported range.
    #
    # @example
    #   Xlsxrb::Utils.valid_excel_date?(46023) #=> true
    #   Xlsxrb::Utils.valid_excel_date?(-1)    #=> false
    #
    # @deprecated Use {Xlsxrb.valid_excel_date?} instead. Will be removed in 0.2.0.
    # @param serial [Object] Serial number candidate.
    # @param date1904 [Boolean] Whether to use the 1904 date system.
    # @return [Boolean]
    #: (untyped serial, ?date1904: bool) -> bool
    def self.valid_excel_date?(serial, date1904: false)
      warn "[DEPRECATION] Xlsxrb::Utils.valid_excel_date? is deprecated; use Xlsxrb.valid_excel_date? instead.", uplevel: 1
      Xlsxrb.valid_excel_date?(serial, date1904: date1904)
    end

    # Converts an Excel serial number to a Date.
    #
    # @example
    #   Xlsxrb::Utils.serial_to_date(46023) #=> #<Date: 2026-01-01>
    #
    # @deprecated Use {Xlsxrb.serial_to_date} instead. Will be removed in 0.2.0.
    # @param serial [Numeric] Excel serial number.
    # @param date1904 [Boolean] Whether to use the 1904 date system.
    # @return [Date]
    #: (Numeric serial, ?date1904: bool) -> Date
    def self.serial_to_date(serial, date1904: false)
      warn "[DEPRECATION] Xlsxrb::Utils.serial_to_date is deprecated; use Xlsxrb.serial_to_date instead.", uplevel: 1
      Xlsxrb.serial_to_date(serial, date1904: date1904)
    end

    # Converts a Time or Date to a fractional Excel serial number.
    #
    # @example
    #   Xlsxrb::Utils.datetime_to_serial(Time.utc(2026, 1, 1, 12, 0, 0)) #=> 46023.5
    #
    # @deprecated Use {Xlsxrb.time_to_serial} instead. Will be removed in 0.2.0.
    # @param time [Time, Date] Time or Date instance.
    # @param date1904 [Boolean] Whether to use the 1904 date system.
    # @return [Float] Fractional serial number.
    #: (Time | Date time, ?date1904: bool) -> Float
    def self.datetime_to_serial(time, date1904: false)
      warn "[DEPRECATION] Xlsxrb::Utils.datetime_to_serial is deprecated; use Xlsxrb.time_to_serial instead.", uplevel: 1
      Xlsxrb.time_to_serial(time, date1904: date1904)
    end

    # Converts a Time or Date to a fractional Excel serial number.
    #
    # @example
    #   Xlsxrb::Utils.time_to_serial(Time.utc(2026, 1, 1, 12, 0, 0)) #=> 46023.5
    #
    # @deprecated Use {Xlsxrb.time_to_serial} instead. Will be removed in 0.2.0.
    # @param time [Time, Date] Time or Date instance.
    # @param date1904 [Boolean] Whether to use the 1904 date system.
    # @return [Float] Fractional serial number.
    #: (Time | Date time, ?date1904: bool) -> Float
    def self.time_to_serial(time, date1904: false)
      warn "[DEPRECATION] Xlsxrb::Utils.time_to_serial is deprecated; use Xlsxrb.time_to_serial instead.", uplevel: 1
      Xlsxrb.time_to_serial(time, date1904: date1904)
    end

    # Converts an Excel serial number to a Time (UTC).
    #
    # @example
    #   Xlsxrb::Utils.serial_to_time(46023.5) #=> 2026-01-01 12:00:00 UTC
    #
    # @deprecated Use {Xlsxrb.serial_to_time} instead. Will be removed in 0.2.0.
    # @param serial [Numeric] Excel serial number.
    # @param date1904 [Boolean] Whether to use the 1904 date system.
    # @return [Time]
    #: (Numeric serial, ?date1904: bool) -> Time
    def self.serial_to_time(serial, date1904: false)
      warn "[DEPRECATION] Xlsxrb::Utils.serial_to_time is deprecated; use Xlsxrb.serial_to_time instead.", uplevel: 1
      Xlsxrb.serial_to_time(serial, date1904: date1904)
    end

    FORMULA_TOKEN_RE = %r{(?<str>"(?:[^"]|"")*")|(?<ref>(?<=\b|\A|[$!:,()+\-*/^<>=& \r\n])(?<col_abs>\$?)(?<col>[A-Za-z]{1,3})(?<row_abs>\$?)(?<row>\d+)(?!\w|\())}

    # Expands a formula by shifting relative cell references from a master cell to a target cell.
    #
    # @example Shift formula down one row
    #   Xlsxrb::Utils.expand_formula("A2*2", from_ref: "B2", to_ref: "B3") #=> "A3*2"
    #
    # @example Shift formula with absolute references
    #   Xlsxrb::Utils.expand_formula("$A$1+A2", from_ref: "B2", to_ref: "B4") #=> "$A$1+A4"
    #
    # @deprecated Use {Xlsxrb.expand_formula} instead. Will be removed in 0.2.0.
    # @param formula [String, nil] Formula expression to expand.
    # @param from_ref [String, nil] Master cell reference (e.g. "B2").
    # @param to_ref [String, nil] Target cell reference (e.g. "B3").
    # @param row_delta [Integer, nil] Row shift amount (positive = down, negative = up).
    # @param column_delta [Integer, nil] Column shift amount (positive = right, negative = left).
    # @param col_delta [Integer, nil] Deprecated alias for column_delta.
    # @return [String, nil] Expanded formula expression.
    #: (String? formula, ?from_ref: String?, ?to_ref: String?, ?row_delta: Integer?, ?col_delta: Integer?, ?column_delta: Integer?) -> String?
    def self.expand_formula(formula, from_ref: nil, to_ref: nil, row_delta: nil, col_delta: nil, column_delta: nil)
      warn "[DEPRECATION] Xlsxrb::Utils.expand_formula is deprecated; use Xlsxrb.expand_formula instead.", uplevel: 1
      warn "[DEPRECATION] Xlsxrb::Utils.expand_formula keyword :col_delta is deprecated; use :column_delta instead.", uplevel: 1 if col_delta && column_delta.nil?
      Xlsxrb.expand_formula(formula, from_ref: from_ref, to_ref: to_ref, row_delta: row_delta, column_delta: column_delta || col_delta)
    end

    # @deprecated Use {Xlsxrb.serial_to_time} instead. Will be removed in 0.2.0.
    #: (Numeric serial, ?date1904: bool) -> Time
    def self.serial_to_datetime(serial, date1904: false)
      warn "[DEPRECATION] Xlsxrb::Utils.serial_to_datetime is deprecated; use Xlsxrb.serial_to_time instead.", uplevel: 1
      Xlsxrb.serial_to_time(serial, date1904: date1904)
    end
  end
end
