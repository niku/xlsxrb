# frozen_string_literal: true

require "date"
require_relative "xlsxrb/version"
require_relative "xlsxrb/zip_generator"
require_relative "xlsxrb/writer"
require_relative "xlsxrb/reader"

# Ruby XLSX read/write library.
module Xlsxrb
  class Error < StandardError; end

  # Represents a formula with an optional cached value.
  # Optional: type (:shared, :array), ref (range), shared_index (si for shared formulas)
  Formula = Data.define(:expression, :cached_value, :type, :ref, :shared_index) do
    def initialize(expression:, cached_value: nil, type: nil, ref: nil, shared_index: nil)
      super
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

  # Excel 1900 date system epoch.
  EPOCH_1900 = Date.new(1899, 12, 31) # serial 1 = Jan 1, 1900

  # Built-in number format codes defined by SpreadsheetML.
  BUILTIN_NUM_FMT_CODES = {
    0 => "General",
    1 => "0",
    2 => "0.00",
    3 => "#,##0",
    4 => "#,##0.00",
    9 => "0%",
    10 => "0.00%",
    11 => "0.00E+00",
    12 => "# ?/?",
    13 => "# ??/??",
    14 => "mm-dd-yy",
    15 => "d-mmm-yy",
    16 => "d-mmm",
    17 => "mmm-yy",
    18 => "h:mm AM/PM",
    19 => "h:mm:ss AM/PM",
    20 => "h:mm",
    21 => "h:mm:ss",
    22 => "m/d/yy h:mm",
    37 => "#,##0 ;(#,##0)",
    38 => "#,##0 ;[Red](#,##0)",
    39 => "#,##0.00;(#,##0.00)",
    40 => "#,##0.00;[Red](#,##0.00)",
    45 => "mm:ss",
    46 => "[h]:mm:ss",
    47 => "mmss.0",
    48 => "##0.0E+0",
    49 => "@"
  }.freeze

  # Built-in numFmtIds that represent date/time formats.
  BUILTIN_DATE_FMT_IDS = [14, 15, 16, 17, 18, 19, 20, 21, 22].freeze

  # Default date format code used by Writer for Date cells.
  DEFAULT_DATE_FORMAT = "yyyy\\-mm\\-dd"

  # Converts a Date to an Excel serial number (1900 system).
  def self.date_to_serial(date)
    serial = (date - EPOCH_1900).to_i
    # Lotus 1-2-3 bug: serial 60 = Feb 29, 1900 (doesn't exist).
    # Dates on or after Mar 1, 1900 (raw serial >= 60) need +1.
    serial += 1 if serial >= 60
    serial
  end

  # Converts an Excel serial number (1900 system) to a Date.
  def self.serial_to_date(serial)
    # Adjust for Lotus 1-2-3 bug.
    serial -= 1 if serial > 60
    EPOCH_1900 + serial
  end
end
