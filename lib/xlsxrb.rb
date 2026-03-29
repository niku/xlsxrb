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
  Formula = Data.define(:expression, :cached_value)

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

  # Built-in numFmtIds that represent date/time formats (14-22).
  BUILTIN_DATE_FMT_IDS = (14..22).to_a.freeze

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
