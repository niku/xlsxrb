# frozen_string_literal: true

require "date"
require "openssl"
require "securerandom"
require_relative "xlsxrb/version"
require_relative "xlsxrb/zip_generator"
require_relative "xlsxrb/writer"
require_relative "xlsxrb/reader"

# Ruby XLSX read/write library.
module Xlsxrb
  class Error < StandardError; end

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
      raise ArgumentError, "invalid error code: #{code.inspect} (must be one of #{Xlsxrb::VALID_ERROR_CODES.join(", ")})" unless Xlsxrb::VALID_ERROR_CODES.include?(code)

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

  # Default date-time format code used by Writer for Time cells.
  DEFAULT_DATETIME_FORMAT = "yyyy\\-mm\\-dd\\ hh:mm:ss"

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

  # Converts a Time to a fractional Excel serial number (1900 system).
  def self.datetime_to_serial(time)
    date = time.to_date
    day_serial = date_to_serial(date)
    # Fractional part: seconds since midnight / seconds per day
    seconds_since_midnight = (time.hour * 3600) + (time.min * 60) + time.sec
    day_serial + (seconds_since_midnight.to_f / 86_400)
  end

  # Converts a fractional Excel serial number to a Time (1900 system, UTC).
  def self.serial_to_datetime(serial)
    int_part = serial.to_i
    frac = serial - int_part
    date = serial_to_date(int_part)
    total_seconds = (frac * 86_400).round
    hours = total_seconds / 3600
    minutes = (total_seconds % 3600) / 60
    seconds = total_seconds % 60
    Time.utc(date.year, date.month, date.day, hours, minutes, seconds)
  end

  # Hashes a plain-text password for use with sheet/workbook protection.
  # Returns { algorithm_name:, hash_value:, salt_value:, spin_count: }.
  # Algorithm per ECMA-376 Part 4 §2.4.2.24.
  def self.hash_password(password, algorithm: "SHA-512", salt: nil, spin_count: 100_000)
    raise ArgumentError, "password must be a String" unless password.is_a?(String)
    raise ArgumentError, "spin_count must be a positive Integer" unless spin_count.is_a?(Integer) && spin_count.positive?

    salt_bytes = salt || SecureRandom.random_bytes(16)
    password_bytes = password.encode("UTF-8").bytes.pack("C*")

    digest_name = algorithm.tr("-", "")
    hash = OpenSSL::Digest.digest(digest_name, salt_bytes + password_bytes)

    spin_count.times do |i|
      iteration_bytes = [i].pack("V") # little-endian uint32
      hash = OpenSSL::Digest.digest(digest_name, iteration_bytes + hash)
    end

    {
      algorithm_name: algorithm,
      hash_value: [hash].pack("m0"),
      salt_value: [salt_bytes].pack("m0"),
      spin_count: spin_count
    }
  end
end
