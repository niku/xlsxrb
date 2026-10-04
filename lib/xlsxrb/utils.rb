# frozen_string_literal: true

# rbs_inline: enabled

require_relative "elements/cell"
require_relative "ooxml/utils"

module Xlsxrb
  # Public coordinate conversion and reference parsing utilities.
  #
  # @api public
  module Utils
    # Converts a cell reference string (e.g. "A1", "BC100") to 0-based [row, col] coordinates.
    #
    # @example
    #   Xlsxrb::Utils.ref_to_row_col("A1")    #=> [0, 0]
    #   Xlsxrb::Utils.ref_to_row_col("BC100") #=> [99, 54]
    #
    # @param cell_ref [String, nil] Excel cell reference string.
    # @return [Array(Integer, Integer), nil] 0-based [row_idx, col_idx] tuple, or nil if invalid.
    # @api public
    #: (String? cell_ref) -> [Integer, Integer]?
    def self.ref_to_row_col(cell_ref)
      Elements::Cell.parse_ref(cell_ref)
    end

    # Converts 0-based [row_idx, col_idx] coordinates to an A1-style reference string.
    #
    # @example
    #   Xlsxrb::Utils.row_col_to_ref(0, 0)   #=> "A1"
    #   Xlsxrb::Utils.row_col_to_ref(99, 54) #=> "BC100"
    #
    # @param row_idx [Integer] 0-based row index.
    # @param col_idx [Integer] 0-based column index.
    # @return [String] A1-style reference string.
    # @api public
    #: (Integer row_idx, Integer col_idx) -> String
    def self.row_col_to_ref(row_idx, col_idx)
      "#{Elements::Cell.column_letter(col_idx)}#{row_idx + 1}"
    end

    # Converts a column letter name (e.g. "A", "BC") to a 0-based integer index.
    #
    # @example
    #   Xlsxrb::Utils.col_name_to_index("A")  #=> 0
    #   Xlsxrb::Utils.col_name_to_index("BC") #=> 54
    #
    # @param col_name [String, Symbol, Integer] Column letter or index.
    # @return [Integer] 0-based column index.
    # @api public
    #: (String | Symbol | Integer col_name) -> Integer
    def self.col_name_to_index(col_name)
      Elements::Cell.column_index(col_name)
    end

    # Converts a 0-based column integer index to a column letter name (e.g. 0 -> "A", 54 -> "BC").
    #
    # @example
    #   Xlsxrb::Utils.col_index_to_name(0)  #=> "A"
    #   Xlsxrb::Utils.col_index_to_name(54) #=> "BC"
    #
    # @param col_idx [Integer] 0-based column index.
    # @return [String] Column letter name.
    # @api public
    #: (Integer col_idx) -> String
    def self.col_index_to_name(col_idx)
      Elements::Cell.column_letter(col_idx)
    end

    # Splits an A1-style cell reference into its column name and 1-based row number.
    #
    # @example
    #   Xlsxrb::Utils.split_coordinate("A1")    #=> ["A", 1]
    #   Xlsxrb::Utils.split_coordinate("BC100") #=> ["BC", 100]
    #
    # @param cell_ref [String, nil] Excel cell reference string.
    # @return [Array(String, Integer), nil] [col_name, row_number] tuple (1-based row), or nil if invalid.
    # @api public
    #: (String? cell_ref) -> [String, Integer]?
    def self.split_coordinate(cell_ref)
      return nil unless cell_ref

      coords = Elements::Cell.parse_ref(cell_ref)
      return nil unless coords

      row_idx, col_idx = coords
      [Elements::Cell.column_letter(col_idx), row_idx + 1]
    end

    # Converts a Date or Time to an Excel numeric serial number (1900 or 1904 system).
    #
    # @example
    #   Xlsxrb::Utils.date_to_serial(Date.new(2026, 1, 1)) #=> 46023
    #
    # @param date_or_time [Date, Time] Date or Time instance.
    # @param date1904 [Boolean] Whether to use the 1904 date system.
    # @return [Integer, Float] Serial number.
    # @api public
    #: (Date | Time date_or_time, ?date1904: bool) -> (Integer | Float)
    def self.date_to_serial(date_or_time, date1904: false)
      if date_or_time.is_a?(Time)
        Ooxml::Utils.datetime_to_serial(date_or_time, date1904: date1904)
      else
        Ooxml::Utils.date_to_serial(date_or_time, date1904: date1904)
      end
    end

    # Converts an Excel serial number to a Date.
    #
    # @example
    #   Xlsxrb::Utils.serial_to_date(46023) #=> #<Date: 2026-01-01>
    #
    # @param serial [Numeric] Excel serial number.
    # @param date1904 [Boolean] Whether to use the 1904 date system.
    # @return [Date]
    # @api public
    #: (Numeric serial, ?date1904: bool) -> Date
    def self.serial_to_date(serial, date1904: false)
      Ooxml::Utils.serial_to_date(serial, date1904: date1904)
    end

    # Converts a Time or Date to a fractional Excel serial number.
    #
    # @example
    #   Xlsxrb::Utils.datetime_to_serial(Time.utc(2026, 1, 1, 12, 0, 0)) #=> 46023.5
    #
    # @param time [Time, Date] Time or Date instance.
    # @param date1904 [Boolean] Whether to use the 1904 date system.
    # @return [Float] Fractional serial number.
    # @api public
    #: (Time | Date time, ?date1904: bool) -> Float
    def self.datetime_to_serial(time, date1904: false)
      if time.is_a?(Time)
        Ooxml::Utils.datetime_to_serial(time, date1904: date1904)
      else
        Ooxml::Utils.date_to_serial(time, date1904: date1904).to_f
      end
    end

    # Converts an Excel serial number to a Time (UTC).
    #
    # @example
    #   Xlsxrb::Utils.serial_to_time(46023.5) #=> 2026-01-01 12:00:00 UTC
    #
    # @param serial [Numeric] Excel serial number.
    # @param date1904 [Boolean] Whether to use the 1904 date system.
    # @return [Time]
    # @api public
    #: (Numeric serial, ?date1904: bool) -> Time
    def self.serial_to_time(serial, date1904: false)
      Ooxml::Utils.serial_to_datetime(serial, date1904: date1904)
    end

    class << self
      alias serial_to_datetime serial_to_time
    end
  end
end
