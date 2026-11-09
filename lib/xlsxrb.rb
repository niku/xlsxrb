# frozen_string_literal: true

# rbs_inline: enabled

require "date"
require "time"
require "openssl"
require "securerandom"
require "tempfile"
require "fileutils"
begin
  require "bigdecimal"
rescue LoadError
  # simplecov:disable
  # Define a dummy class for environments without bigdecimal (e.g., ruby.wasm).
  # This serves only as a fallback to prevent NameError in `case` statements (`when BigDecimal`).
  # Impossible to cover in standard test environment because bigdecimal is present.
  Object.const_set(:BigDecimal, Class.new)
  # simplecov:enable
end
require_relative "xlsxrb/version"
require_relative "xlsxrb/ooxml/zip_generator"
require_relative "xlsxrb/ooxml/writer"
require_relative "xlsxrb/ooxml/reader"
require_relative "xlsxrb/ooxml/crypto"
require_relative "xlsxrb/ooxml"
require_relative "xlsxrb/elements"
require_relative "xlsxrb/stream_row"
require_relative "xlsxrb/stream_sheet"
require_relative "xlsxrb/colors"
require_relative "xlsxrb/style_builder"
require_relative "xlsxrb/chart_builder"
require_relative "xlsxrb/dsl_helpers"
require_relative "xlsxrb/worksheet_builder"
require_relative "xlsxrb/workbook_builder"
require_relative "xlsxrb/number_formatter"
require_relative "xlsxrb/stream_writer"

# Modern, streaming-capable, low-memory XLSX reading and writing library for Ruby.
#
# Provides high-level facade methods ({Xlsxrb.read}, {Xlsxrb.write}, {Xlsxrb.build}, {Xlsxrb.load})
# with full ECMA-376 OpenXML compliance, MS-OFFCRYPTO password protection, and zero core dependencies.
#
# @api public
module Xlsxrb
  # Maximum number of rows supported by Excel / OpenXML.
  MAX_ROW = 1_048_576

  # Maximum number of columns supported by Excel / OpenXML.
  MAX_COLUMN = 16_384

  # Maximum number of individual cell/range references allowed in a single data validation element by Excel.
  MAX_REFS_PER_VALIDATION = 32_767

  # ZIP deflate level used when no compression_level is given.
  # See docs/adr/0009-default-zip-compression-level.md.
  DEFAULT_COMPRESSION_LEVEL = 1

  # Base class for all exceptions raised by Xlsxrb.
  class Error < StandardError; end

  # Raised when parsing malformed or invalid XML structures.
  class ParseError < Error; end

  OMITTED_TARGET = Object.new.freeze
  VALID_SHEET_STATES = %i[visible hidden very_hidden].freeze
  OFFICE_DOC_REL_TYPES = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"
  ].freeze
  NON_WORKSHEET_REL_SUFFIXES = %w[
    /chartsheet
    /dialogsheet
    /xlMacrosheet
    /xlIntlMacrosheet
    /macrosheet
  ].freeze
  NON_WORKSHEET_TARGET_PREFIXES = %w[
    chartsheets/
    dialogsheets/
    macrosheets/
    xl/chartsheets/
    xl/dialogsheets/
    xl/macrosheets/
  ].freeze
  private_constant :OMITTED_TARGET, :VALID_SHEET_STATES, :OFFICE_DOC_REL_TYPES, :NON_WORKSHEET_REL_SUFFIXES, :NON_WORKSHEET_TARGET_PREFIXES

  # Raised when cell, row, or workbook parameters fail specification validation.
  class ValidationError < Error; end

  # Raised when ZIP decompression or entry resolution fails.
  class ZipError < Error; end

  # Raised when ZIP decompression exceeds size bounds (ZIP bomb defense).
  class ZipBombError < ZipError; end

  # Base exception for encrypted or password-protected XLSX documents.
  class EncryptedFileError < Error; end

  # Raised when a password is missing or invalid for an encrypted XLSX file.
  class InvalidPasswordError < EncryptedFileError; end

  # Raised when encrypted package stream decryption fails.
  class DecryptionError < EncryptedFileError; end

  # Helper to construct {Elements::RichText} objects with formatted text runs.
  #
  # @example Create multi-run rich text with hashes
  #   rt = Xlsxrb.rich_text({ text: "Total: ", font: { bold: true } }, { text: "$100" })
  #
  # @example Create rich text with text/font pairs
  #   rt = Xlsxrb.rich_text("Normal text, ", ["bold text, ", { bold: true }], ["red text", { color: "FF0000" }])
  #
  # @example Create rich text with format objects implementing to_font_hash
  #   rt = Xlsxrb.rich_text("Normal text, ", bold_format, "bold text")
  #
  # @example Create simple styled text
  #   rt = Xlsxrb.rich_text(text: "Important Notice", bold: true, color: "FF0000")
  #
  # @param runs [Array<Hash, String, Array, Elements::RichTextRun, Object>] Optional array of rich text run items.
  # @param text [String, nil] Plain text string (convenience parameter).
  # @param base_font [Hash, Object, nil] Base font properties to merge into all runs.
  # @param font_props [Hash] Inline font styling options (e.g. bold: true, color: "FF0000").
  # @return [Elements::RichText] The compiled rich text element.
  # @api public
  #: (*(Hash[Symbol, untyped] | Elements::RichTextRun | String | Array[untyped] | untyped) runs, ?text: String?, ?base_font: untyped, **untyped font_props) -> Elements::RichText
  def self.rich_text(*runs, text: nil, base_font: nil, **font_props)
    base = if base_font.respond_to?(:to_font_hash)
             base_font.to_font_hash
           elsif base_font.respond_to?(:to_h)
             base_font.to_h
           else
             base_font
           end
    if text
      f = font_props.empty? ? nil : font_props
      f = base.merge(f || {}) if base.is_a?(Hash)
      runs = [Elements::RichTextRun.new(text: text, font: f)]
    elsif runs.size == 1 && runs.first.is_a?(String) && !font_props.empty?
      f = base.is_a?(Hash) ? base.merge(font_props) : font_props
      runs = [Elements::RichTextRun.new(text: runs.first, font: f)]
    else
      normalized_runs = []
      pending_font = nil

      runs.each do |item|
        if item.respond_to?(:to_font_hash)
          pending_font = item.to_font_hash
        elsif item.is_a?(Hash) && !item.key?(:text)
          pending_font = item
        elsif item.is_a?(Elements::RichTextRun)
          normalized_runs << item
          pending_font = nil
        elsif item.is_a?(Hash)
          f = item[:font] || pending_font
          f = f.to_font_hash if f.respond_to?(:to_font_hash)
          normalized_runs << Elements::RichTextRun.new(text: item[:text].to_s, font: f)
          pending_font = nil
        elsif item.is_a?(Array)
          t = item[0].to_s
          f = item[1] || pending_font
          f = f.to_font_hash if f.respond_to?(:to_font_hash)
          f = f.to_h if f.respond_to?(:to_h) && !f.is_a?(Hash)
          normalized_runs << Elements::RichTextRun.new(text: t, font: f.is_a?(Hash) ? f : nil)
          pending_font = nil
        elsif pending_font
          normalized_runs << Elements::RichTextRun.new(text: item.to_s, font: pending_font)
          pending_font = nil
        else
          normalized_runs << Elements::RichTextRun.new(text: item.to_s)
        end
      end
      if base.is_a?(Hash) && !base.empty?
        normalized_runs = normalized_runs.map do |run|
          merged_font = run.font ? base.merge(run.font) : base.dup
          Elements::RichTextRun.new(text: run.text, font: merged_font)
        end
      end
      runs = normalized_runs
    end
    Elements::RichText.new(runs: runs)
  end

  # Creates an {Elements::Formula} object for use in cell values.
  #
  # @example Create a formula without precalculated value
  #   formula = Xlsxrb.formula("SUM(A1:A10)")
  #
  # @example Create a formula with precomputed cached value
  #   formula = Xlsxrb.formula("A1+B1", cached_value: 42)
  #
  # @param expression [String] The formula expression without leading '=' (e.g. "SUM(A1:A10)").
  # @param cached_value [Object, nil] Optional precomputed value for readers that do not evaluate formulas.
  # @param type [Symbol, String, nil] Formula type (:shared, :array, etc.).
  # @param ref [String, nil] Target cell or range reference (e.g. "A1:C1").
  # @param shared_index [Integer, nil] Shared formula group index (0-based).
  # @return [Elements::Formula]
  # @api public
  #: (String expression, ?cached_value: (String | Numeric | bool | Elements::CellError | nil), ?type: (Symbol | String)?, ?ref: String?, ?shared_index: Integer?) -> Elements::Formula
  def self.formula(expression, cached_value: nil, type: nil, ref: nil, shared_index: nil)
    Elements::Formula.new(
      expression: expression,
      cached_value: cached_value,
      type: type || (shared_index ? :shared : nil),
      ref: ref,
      shared_index: shared_index,
      calculate_always: cached_value.nil? || nil
    )
  end

  # Creates a shared formula {Elements::Formula} object for master or follower cell.
  #
  # @example Create a shared formula master
  #   formula = Xlsxrb.shared_formula("A2*2", shared_index: 0, ref: "B2:B10")
  #
  # @example Create a shared formula follower
  #   formula = Xlsxrb.shared_formula(shared_index: 0)
  #
  # @param expression [String, nil] The formula expression (nil for follower cells).
  # @param shared_index [Integer, nil] Shared formula group index (0-based).
  # @param ref [String, nil] Target range for master cell (e.g. "B2:B10").
  # @param cached_value [Object, nil] Optional precomputed value.
  # @return [Elements::Formula]
  # @api public
  #: (?String? expression, ?shared_index: Integer?, ?ref: String?, ?cached_value: (String | Numeric | bool | Elements::CellError | nil)) -> Elements::Formula
  def self.shared_formula(expression = nil, shared_index: nil, ref: nil, cached_value: nil)
    raise ArgumentError, "shared_index is required for shared formula" if shared_index.nil?

    Elements::Formula.new(
      expression: expression || "",
      cached_value: cached_value,
      type: :shared,
      ref: ref,
      shared_index: shared_index,
      calculate_always: cached_value.nil? || nil
    )
  end

  # Creates an array or dynamic array {Elements::Formula} object.
  #
  # @example Create a multi-cell array formula
  #   formula = Xlsxrb.array_formula("SUM(B1:B10*C1:C10)", ref: "A1:C1", cached_value: 0)
  #
  # @example Create a dynamic array formula
  #   formula = Xlsxrb.array_formula("SORT(UNIQUE(A1:A10))", ref: "D1")
  #
  # @param expression [String] The formula expression without leading '='.
  # @param ref [String, nil] Target cell or range reference (keyword).
  # @param cached_value [Object, nil] Optional precomputed value.
  # @return [Elements::Formula]
  # @api public
  #: (String expression, ?ref: String?, ?cached_value: (String | Numeric | bool | Elements::CellError | nil), **untyped attrs) -> Elements::Formula
  def self.array_formula(expression, ref: nil, cached_value: nil, **attrs)
    Elements::Formula.new(
      expression: expression,
      cached_value: cached_value,
      type: :array,
      ref: ref,
      calculate_always: cached_value.nil? || nil,
      **attrs
    )
  end

  # Creates an {Elements::Hyperlink} object for use in cell hyperlinks.
  #
  # @example Create an external hyperlink with a ScreenTip tooltip
  #   link = Xlsxrb.hyperlink("https://example.com", tooltip: "Visit Example")
  #
  # @example Create an internal worksheet link
  #   link = Xlsxrb.hyperlink(location: "Sheet2!A1", tooltip: "Go to Sheet 2")
  #
  # @param url [String, nil] External hyperlink URL.
  # @param display [String, nil] Optional display label.
  # @param tooltip [String, nil] Optional hover ScreenTip text.
  # @param location [String, nil] Optional internal sheet coordinate.
  # @return [Elements::Hyperlink]
  # @api public
  #: (?String? url, ?display: String?, ?tooltip: String?, ?location: String?) -> Elements::Hyperlink
  def self.hyperlink(url = nil, display: nil, tooltip: nil, location: nil)
    Elements::Hyperlink.new(url: url, display: display, tooltip: tooltip, location: location)
  end

  # Creates an {Elements::Comment} object for use in cell comments.
  #
  # @example Create a cell comment
  #   comment = Xlsxrb.comment("Review this value", author: "Auditor")
  #
  # @param text [String, nil] Comment body text.
  # @param author [String] Comment author.
  # @param visible [Boolean] Whether comment is visible by default.
  # @param ref [String, nil] Cell coordinate (e.g. "A1").
  # @return [Elements::Comment]
  # @api public
  #: (?String? text, ?author: String, ?visible: bool, ?ref: String?, **untyped extra) -> Elements::Comment
  def self.comment(text = nil, author: "Author", visible: false, ref: nil, **extra)
    Elements::Comment.new(text: text, author: author, visible: visible, ref: ref, **extra)
  end

  # Creates or returns an {Elements::CellError} object.
  #
  # @example
  #   err = Xlsxrb.cell_error("#VALUE!")
  #   err = Xlsxrb.cell_error(:value)
  #
  # @param code [String, Symbol, Elements::CellError] Error code or symbol.
  # @return [Elements::CellError]
  # @api public
  #: (String | Symbol | Elements::CellError code) -> Elements::CellError
  def self.cell_error(code)
    return code if code.is_a?(Elements::CellError)

    code_str = code.to_s
    error_map = {
      null: "#NULL!",
      div0: "#DIV/0!",
      value: "#VALUE!",
      ref: "#REF!",
      name: "#NAME?",
      num: "#NUM!",
      na: "#N/A",
      getting_data: "#GETTING_DATA"
    }
    resolved_code = error_map[code.to_sym] || code_str
    Elements::CellError.new(code: resolved_code)
  end

  # Formats a cell value according to an Excel format code string.
  #
  # @param value [Object, nil] Value to format.
  # @param format_code [String, nil] Excel format code string.
  # @param date1904 [Boolean] Whether the 1904 date system is active.
  # @return [String] Formatted string representation.
  # @api public
  #: (untyped value, ?String? format_code, ?date1904: bool) -> String
  def self.format_value(value, format_code = nil, date1904: false)
    NumberFormatter.format(value, format_code, date1904: date1904)
  end

  # Converts a Date or Time to an Excel numeric serial number (1900 or 1904 system).
  #
  # @example
  #   Xlsxrb.date_to_serial(Date.new(2026, 1, 1)) #=> 46023
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

  # Checks whether a serial number is a valid Excel date within the supported range.
  #
  # @example
  #   Xlsxrb.valid_excel_date?(46023) #=> true
  #   Xlsxrb.valid_excel_date?(-1)    #=> false
  #
  # @param serial [Object] Serial number candidate.
  # @param date1904 [Boolean] Whether to use the 1904 date system.
  # @return [Boolean]
  # @api public
  #: (untyped serial, ?date1904: bool) -> bool
  def self.valid_excel_date?(serial, date1904: false)
    Ooxml::Utils.valid_excel_date?(serial, date1904: date1904)
  end

  # Converts an Excel serial number to a Date.
  #
  # @example
  #   Xlsxrb.serial_to_date(46023) #=> #<Date: 2026-01-01>
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
  #   Xlsxrb.time_to_serial(Time.utc(2026, 1, 1, 12, 0, 0)) #=> 46023.5
  #
  # @param time [Time, Date] Time or Date instance.
  # @param date1904 [Boolean] Whether to use the 1904 date system.
  # @return [Float] Fractional serial number.
  # @api public
  #: (Time | Date time, ?date1904: bool) -> Float
  def self.time_to_serial(time, date1904: false)
    if time.is_a?(Time)
      Ooxml::Utils.datetime_to_serial(time, date1904: date1904)
    else
      Ooxml::Utils.date_to_serial(time, date1904: date1904).to_f
    end
  end

  # Converts an Excel serial number to a Time (UTC).
  #
  # @example
  #   Xlsxrb.serial_to_time(46023.5) #=> 2026-01-01 12:00:00 UTC
  #
  # @param serial [Numeric] Excel serial number.
  # @param date1904 [Boolean] Whether to use the 1904 date system.
  # @return [Time]
  # @api public
  #: (Numeric serial, ?date1904: bool) -> Time
  def self.serial_to_time(serial, date1904: false)
    Ooxml::Utils.serial_to_datetime(serial, date1904: date1904)
  end

  # Converts a cell reference string (e.g. "A1", "BC100") to 0-based [row, col] coordinates.
  #
  # @example
  #   Xlsxrb.parse_ref("A1")    #=> [0, 0]
  #   Xlsxrb.parse_ref("BC100") #=> [99, 54]
  #
  # @param ref [String, nil] Excel cell reference string.
  # @return [Array(Integer, Integer), nil] 0-based [row, col] tuple, or nil if invalid.
  # @api public
  #: (String? ref) -> [Integer, Integer]?
  def self.parse_ref(ref)
    Elements::Cell._parse_ref(ref)
  end

  # Converts 0-based [row, col] coordinates to an A1-style reference string.
  #
  # @example
  #   Xlsxrb.cell_to_ref(0, 0)   #=> "A1"
  #   Xlsxrb.cell_to_ref(99, 54) #=> "BC100"
  #
  # @param row [Integer] 0-based row index.
  # @param column [Integer] 0-based column index.
  # @return [String] A1-style reference string.
  # @api public
  #: (Integer row, Integer column) -> String
  def self.cell_to_ref(row, column)
    "#{Elements::Cell._column_letter(column)}#{row + 1}"
  end

  # Converts a column letter (e.g. "A", :AA) to a 0-based column index.
  #
  # @example
  #   Xlsxrb.column_index("A")  #=> 0
  #   Xlsxrb.column_index(:AA)  #=> 26
  #
  # @param col [String, Symbol, Integer] Column letter or integer index.
  # @return [Integer] 0-based column index.
  # @api public
  #: (String | Symbol | Integer col) -> Integer
  def self.column_index(col)
    Elements::Cell._column_index(col)
  end

  # Converts a 0-based column index to an Excel letter (0 -> "A", 25 -> "Z", 26 -> "AA").
  #
  # @example
  #   Xlsxrb.column_letter(0)  #=> "A"
  #   Xlsxrb.column_letter(26) #=> "AA"
  #
  # @param index [Integer] 0-based column index.
  # @return [String] Excel column letter.
  # @api public
  #: (Integer index) -> String
  def self.column_letter(index)
    Elements::Cell._column_letter(index)
  end

  # Converts a cell range reference string (e.g. "A1:C10", "A1") to 0-based [[top_row, left_col], [bottom_row, right_col]].
  #
  # @example
  #   Xlsxrb.parse_range("A1:C10") #=> [[0, 0], [9, 2]]
  #
  # @param range_ref [String, nil]
  # @return [Array([Integer, Integer], [Integer, Integer]), nil]
  # @api public
  #: (String? range_ref) -> [[Integer, Integer], [Integer, Integer]]?
  def self.parse_range(range_ref)
    return nil if range_ref.nil? || range_ref.empty?

    parts = range_ref.split(":", 2)
    first = Elements::Cell._parse_ref(parts[0])
    return nil unless first

    second = parts[1] ? Elements::Cell._parse_ref(parts[1]) : first
    return nil unless second

    top_row = [first[0], second[0]].min
    bottom_row = [first[0], second[0]].max
    left_col = [first[1], second[1]].min
    right_col = [first[1], second[1]].max
    [[top_row, left_col], [bottom_row, right_col]]
  end

  # Converts 0-based bounding coordinates to an A1-style range reference string.
  # Accepts either 4 flat integers (top_row, left_column, bottom_row, right_column)
  # or a nested pair ([[top_row, left_column], [bottom_row, right_column]]) as returned by parse_range.
  #
  # @example 4 flat integers
  #   Xlsxrb.range_to_ref(0, 0, 9, 2) #=> "A1:C10"
  #
  # @example Nested pair from parse_range
  #   Xlsxrb.range_to_ref([[0, 0], [9, 2]]) #=> "A1:C10"
  #
  # @param top_row_or_bounds [Integer, Array] 0-based start row index or nested bounding coordinates.
  # @param left_column [Integer, nil] 0-based start column index.
  # @param bottom_row [Integer, nil] 0-based end row index.
  # @param right_column [Integer, nil] 0-based end column index.
  # @return [String] A1-style range reference string.
  # @api public
  #: (Integer | [[Integer, Integer], [Integer, Integer]] top_row_or_bounds, ?Integer? left_column, ?Integer? bottom_row, ?Integer? right_column) -> String
  def self.range_to_ref(top_row_or_bounds, left_column = nil, bottom_row = nil, right_column = nil)
    if top_row_or_bounds.is_a?(Array)
      first, second = top_row_or_bounds
      top_row = first[0]
      left_col = first[1]
      bot_row = second[0]
      rgt_col = second[1]
    else
      top_row = top_row_or_bounds
      left_col = left_column || 0
      bot_row = bottom_row || top_row
      rgt_col = right_column || left_col
    end
    if top_row == bot_row && left_col == rgt_col
      cell_to_ref(top_row, left_col)
    else
      "#{cell_to_ref(top_row, left_col)}:#{cell_to_ref(bot_row, rgt_col)}"
    end
  end

  # Splits an A1-style cell reference into its column name and 1-based row number.
  #
  # @example
  #   Xlsxrb.split_coordinate("A1")    #=> ["A", 1]
  #   Xlsxrb.split_coordinate("BC100") #=> ["BC", 100]
  #
  # @param cell_ref [String, nil] Excel cell reference string.
  # @return [Array(String, Integer), nil] [col_name, row_number] tuple (1-based row), or nil if invalid.
  # @api public
  #: (String? cell_ref) -> [String, Integer]?
  def self.split_coordinate(cell_ref)
    return nil unless cell_ref

    coords = Elements::Cell._parse_ref(cell_ref)
    return nil unless coords

    row_idx, col_idx = coords
    [Elements::Cell._column_letter(col_idx), row_idx + 1]
  end

  FORMULA_TOKEN_RE = %r{(?<str>"(?:[^"]|"")*")|(?<ref>(?<=\b|\A|[$!:,()+\-*/^<>=& \r\n])(?<col_abs>\$?)(?<col>[A-Za-z]{1,3})(?<row_abs>\$?)(?<row>\d+)(?!\w|\())}
  private_constant :FORMULA_TOKEN_RE

  # Expands a formula by shifting relative cell references from a master cell to a target cell or by offset amounts.
  #
  # @example Shift formula down one row
  #   Xlsxrb.expand_formula("A2*2", row_delta: 1, column_delta: 0) #=> "A3*2"
  #   Xlsxrb.expand_formula("A2*2", from_ref: "B2", to_ref: "B3") #=> "A3*2"
  #
  # @example Shift formula with absolute references
  #   Xlsxrb.expand_formula("$A$1+A2", row_delta: 2, column_delta: 0) #=> "$A$1+A4"
  #
  # @param formula [String, nil] Formula expression to expand.
  # @param from_ref [String, nil] Master cell reference (e.g. "B2").
  # @param to_ref [String, nil] Target cell reference (e.g. "B3").
  # @param row_delta [Integer, nil] Row shift amount (positive = down, negative = up).
  # @param column_delta [Integer, nil] Column shift amount (positive = right, negative = left).
  # @return [String, nil] Expanded formula expression.
  # @api public
  #: (String? formula, ?from_ref: String?, ?to_ref: String?, ?row_delta: Integer?, ?column_delta: Integer?) -> String?
  def self.expand_formula(formula, from_ref: nil, to_ref: nil, row_delta: nil, column_delta: nil)
    return formula if formula.nil? || formula.empty?

    if from_ref && to_ref
      f_coords = Elements::Cell._parse_ref(from_ref)
      t_coords = Elements::Cell._parse_ref(to_ref)
      return formula unless f_coords && t_coords

      rd = t_coords[0] - f_coords[0]
      cd = t_coords[1] - f_coords[1]
    else
      rd = row_delta || 0
      cd = column_delta || 0
    end

    return formula if rd.zero? && cd.zero?

    formula.gsub(FORMULA_TOKEN_RE) do |match|
      m = Regexp.last_match
      if m[:str]
        match
      else
        c_abs = !m[:col_abs].empty?
        r_abs = !m[:row_abs].empty?
        col = m[:col]
        row = m[:row].to_i

        new_col = if c_abs || cd.zero?
                    col
                  else
                    ci = Elements::Cell._column_index(col) + cd
                    ci = ci.clamp(0, 16_383)
                    Elements::Cell._column_letter(ci)
                  end

        new_row = if r_abs || rd.zero?
                    row
                  else
                    ri = row + rd
                    ri.clamp(1, 1_048_576)
                  end

        "#{m[:col_abs]}#{new_col}#{m[:row_abs]}#{new_row}"
      end
    end
  end

  # Resolves an ECMA-376 indexed color palette index (0..63) to an RRGGBB hex string (or AARRGGBB if alpha: true).
  #
  # @example
  #   Xlsxrb.palette_color(12)               #=> "0000FF"
  #   Xlsxrb.palette_color(12, alpha: true)  #=> "FF0000FF"
  #
  # @param index [Object, nil] Palette index (0..63).
  # @param alpha [Boolean] Whether to prefix "FF" alpha channel.
  # @return [String, nil]
  # @api public
  #: (untyped index, ?alpha: bool) -> String?
  def self.palette_color(index, alpha: false)
    Colors.palette_color(index, alpha: alpha)
  end

  # Converts any supported color representation (palette index 0..63, CSS name, #RGB, #RRGGBB, #AARRGGBB, integer RGB)
  # to a normalized hex string.
  #
  # @example
  #   Xlsxrb.to_hex_color(12)                      #=> "0000FF"
  #   Xlsxrb.to_hex_color(:navy)                   #=> "000080"
  #   Xlsxrb.to_hex_color("#FF0000")               #=> "FF0000"
  #
  # @param color [Integer, String, Symbol, nil]
  # @param alpha [Boolean] Whether to include the alpha channel (default false -> "RRGGBB").
  # @return [String, nil]
  # @api public
  #: (Integer | String | Symbol | nil color, ?alpha: bool) -> String?
  def self.to_hex_color(color, alpha: false)
    Colors.to_hex(color, alpha: alpha)
  end

  # Sanitizes a worksheet name to ensure it conforms to Excel constraints within 31 characters.
  # Replaces forbidden characters (\ / ? * [ ] : and control characters),
  # replaces leading/trailing apostrophes, and replaces reserved 'History' with 'History_'.
  #
  # @example
  #   Xlsxrb.sanitize_sheet_name("Sales: 2026") #=> "Sales_ 2026"
  #   Xlsxrb.sanitize_sheet_name("'Report'")    #=> "_Report_"
  #
  # @param name [Object] Desired worksheet name.
  # @param replace_char [String] Replacement character for invalid characters.
  # @return [String] Excel-compliant worksheet name.
  # @api public
  #: (untyped name, ?replace_char: String) -> String
  def self.sanitize_sheet_name(name, replace_char: "_")
    DslHelpers.sanitize_sheet_name(name, replace_char: replace_char)
  end

  # Reads an XLSX file (streaming and lazy-loaded by default) from a file path, IO stream, or binary String.
  #
  # Sheets and rows are streamed lazily with O(1) constant memory. If a block is given,
  # yields each {StreamSheet} sequentially. Call {StreamSheet#load} to convert a sheet to
  # an in-memory representation for coordinate random access (`sheet["A1"]`).
  #
  # @example Streaming read across sheets and rows (O(1) constant memory)
  #   Xlsxrb.read("large.xlsx") do |sheet|
  #     puts "Sheet: #{sheet.name}"
  #     sheet.each_row do |row|
  #       row.each_cell { |cell| puts "#{cell.ref}: #{cell.value}" }
  #     end
  #   end
  #
  # @param source [String, IO, StringIO] File path, binary content string, or readable IO stream.
  # @param password [String, nil] Optional password to decrypt password-protected XLSX files.
  # @param trim_empty_rows [Boolean] Whether to omit trailing empty rows during sheet iteration.
  # @param pad_empty_rows [Boolean] Whether to yield empty rows for skipped row numbers.
  # @param pad_empty_cells [Boolean] Whether to pad missing cells within rows.
  # @param max_uncompressed_size [Integer, nil] Maximum allowed uncompressed size in bytes per entry (ZIP bomb defense).
  # @yield [sheet]
  # @yieldparam sheet [StreamSheet] The streaming worksheet object.
  # @return [Enumerator<StreamSheet>, void]
  # @raise [ZipBombError] If an uncompressed entry exceeds max_uncompressed_size.
  # @raise [EncryptedFileError] If the file is encrypted and no password was supplied.
  # @raise [InvalidPasswordError] If the supplied password is incorrect.
  # @raise [ParseError] If the spreadsheet XML structure is invalid.
  # @api public
  #: (untyped source, ?password: String?, ?trim_empty_rows: bool, ?pad_empty_rows: bool, ?pad_empty_cells: bool, ?max_uncompressed_size: Integer?) { (StreamSheet) -> void } -> void
  #: (untyped source, ?password: String?, ?trim_empty_rows: bool, ?pad_empty_rows: bool, ?pad_empty_cells: bool, ?max_uncompressed_size: Integer?) -> Enumerator[StreamSheet, void]
  def self.read(source, password: nil, trim_empty_rows: false, pad_empty_rows: false, pad_empty_cells: false, max_uncompressed_size: nil, &)
    prepared_source = prepare_source_io(source, password)
    return to_enum(:read, source, password: password, trim_empty_rows: trim_empty_rows, pad_empty_rows: pad_empty_rows, pad_empty_cells: pad_empty_cells, max_uncompressed_size: max_uncompressed_size) unless block_given?

    zip_options = {}
    zip_options[:max_uncompressed_size] = max_uncompressed_size if max_uncompressed_size
    zip_reader = Ooxml::ZipReader.open(prepared_source, **zip_options)

    begin
      workbook_path = resolve_workbook_entry(zip_reader)
      wb_dir = File.dirname(workbook_path)
      wb_rels_name = wb_dir == "." ? "_rels/#{File.basename(workbook_path)}.rels" : "#{wb_dir}/_rels/#{File.basename(workbook_path)}.rels"
      rels_xml = zip_reader.read_entry(wb_rels_name)
      rels, rel_types = parse_workbook_relationships(rels_xml)

      sst_entry = resolve_part_entry(zip_reader, rels, "sharedStrings", "sharedStrings.xml", wb_dir)
      shared_strings_xml = sst_entry ? zip_reader.read_entry(sst_entry) : nil
      shared_strings = shared_strings_xml ? Ooxml::SharedStringsParser.parse(shared_strings_xml) : []

      styles_entry = resolve_part_entry(zip_reader, rels, "styles", "styles.xml", wb_dir)
      styles_xml = styles_entry ? zip_reader.read_entry(styles_entry) : nil
      styles = styles_xml ? Ooxml::StylesParser.parse(styles_xml) : {}

      workbook_xml = zip_reader.read_entry(workbook_path)
      wb_parsed = Ooxml::WorkbookParser.parse_with_properties(workbook_xml)
      workbook_sheets = wb_parsed[:sheets]
      is_date1904 = wb_parsed[:date1904]

      ws_idx = 0
      workbook_sheets.each do |sheet_info|
        r_id = sheet_info[:r_id]
        target = rels[r_id]
        rel_type = rel_types[r_id]
        next if non_worksheet_relationship?(rel_type, target)

        sheet_path = resolve_sheet_entry(zip_reader, target, wb_dir, ws_idx)
        current_sheet_idx = ws_idx
        ws_idx += 1
        next unless sheet_path

        chunk_supplier = ->(&blk) { zip_reader.each_entry_chunk(sheet_path, &blk) }

        stream_sheet = StreamSheet.new(
          sheet_info[:name],
          chunk_supplier,
          shared_strings,
          styles,
          zip_reader: zip_reader,
          entry_name: sheet_path,
          state: sheet_info[:state] || :visible,
          date1904: is_date1904,
          trim_empty_rows: trim_empty_rows,
          pad_empty_rows: pad_empty_rows,
          pad_empty_cells: pad_empty_cells,
          defined_names: wb_parsed[:defined_names] || [],
          sheet_index: current_sheet_idx
        )
        yield stream_sheet
      end
      nil
    ensure
      zip_reader&.close
    end
  end

  # Loads all sheets into memory, returning an {Elements::Workbook} where every
  # worksheet is a fully-parsed {Elements::Worksheet} supporting coordinate random access.
  # The underlying zip resources are closed immediately.
  #
  # @example
  #   wb = Xlsxrb.load("report.xlsx")
  #   puts wb["Sheet1"]["A1"].value
  #
  # @param source [String, IO, StringIO] Path to the .xlsx file or an IO-like stream.
  # @param password [String, nil] Password for decrypting password-protected (ECMA-376 Standard/Agile) workbooks.
  # @param trim_empty_rows [Boolean] Whether to omit trailing empty rows from worksheet data.
  # @param pad_empty_rows [Boolean] Whether to pad missing rows with empty row objects.
  # @param pad_empty_cells [Boolean] Whether to pad missing cells within rows with nil cells.
  # @param max_uncompressed_size [Integer, nil] Maximum uncompressed file size threshold in bytes for zip bomb protection.
  # @return [Elements::Workbook] The fully loaded in-memory workbook.
  # @api public
  #: (untyped source, ?password: String?, ?trim_empty_rows: bool, ?pad_empty_rows: bool, ?pad_empty_cells: bool, ?max_uncompressed_size: Integer?) -> Elements::Workbook
  def self.load(source, password: nil, trim_empty_rows: false, pad_empty_rows: false, pad_empty_cells: false, max_uncompressed_size: nil)
    prepared_source = prepare_source_io(source, password)
    zip_options = {}
    zip_options[:max_uncompressed_size] = max_uncompressed_size if max_uncompressed_size
    zip_reader = Ooxml::ZipReader.open(prepared_source, **zip_options)

    begin
      workbook_path = resolve_workbook_entry(zip_reader)
      wb_dir = File.dirname(workbook_path)
      wb_rels_name = wb_dir == "." ? "_rels/#{File.basename(workbook_path)}.rels" : "#{wb_dir}/_rels/#{File.basename(workbook_path)}.rels"
      rels_xml = zip_reader.read_entry(wb_rels_name)
      rels, rel_types = parse_workbook_relationships(rels_xml)

      sst_entry = resolve_part_entry(zip_reader, rels, "sharedStrings", "sharedStrings.xml", wb_dir)
      shared_strings_xml = sst_entry ? zip_reader.read_entry(sst_entry) : nil
      shared_strings = shared_strings_xml ? Ooxml::SharedStringsParser.parse(shared_strings_xml) : []

      styles_entry = resolve_part_entry(zip_reader, rels, "styles", "styles.xml", wb_dir)
      styles_xml = styles_entry ? zip_reader.read_entry(styles_entry) : nil
      styles = styles_xml ? Ooxml::StylesParser.parse(styles_xml) : {}

      workbook_xml = zip_reader.read_entry(workbook_path)
      wb_parsed = Ooxml::WorkbookParser.parse_with_properties(workbook_xml)
      workbook_sheets = wb_parsed[:sheets]
      is_date1904 = wb_parsed[:date1904]
      defined_names = wb_parsed[:defined_names] || []

      non_worksheet_sheets = []
      ws_idx = 0
      sheets = workbook_sheets.filter_map do |sheet_info|
        r_id = sheet_info[:r_id]
        target = rels[r_id]
        rel_type = rel_types[r_id]
        if non_worksheet_relationship?(rel_type, target)
          non_worksheet_sheets << sheet_info.merge(target: target, type: rel_type)
          next
        end

        sheet_path = resolve_sheet_entry(zip_reader, target, wb_dir, ws_idx)
        current_sheet_idx = ws_idx
        ws_idx += 1
        next unless sheet_path

        sheet_xml = zip_reader.read_entry(sheet_path)
        sheet_name = sheet_info[:name]
        pa = extract_sheet_print_area(defined_names, sheet_name, current_sheet_idx)
        pt = extract_sheet_print_titles(defined_names, sheet_name, current_sheet_idx)

        build_worksheet(
          sheet_name,
          sheet_xml,
          shared_strings,
          styles,
          state: sheet_info[:state] || :visible,
          zip_reader: zip_reader,
          entry_name: sheet_path,
          date1904: is_date1904,
          trim_empty_rows: trim_empty_rows,
          pad_empty_rows: pad_empty_rows,
          pad_empty_cells: pad_empty_cells,
          print_area: pa,
          print_titles: pt
        )
      end

      wb_unmapped = {}
      wb_props = {}
      wb_props[:date1904] = true if is_date1904
      wb_props[:active_sheet] = wb_parsed[:active_tab] unless wb_parsed[:active_tab].nil?
      wb_props[:active_sheet_index] = wb_parsed[:active_tab] unless wb_parsed[:active_tab].nil?
      wb_props[:first_sheet] = wb_parsed[:first_sheet] unless wb_parsed[:first_sheet].nil?
      wb_props[:show_sheet_tabs] = wb_parsed[:show_sheet_tabs] unless wb_parsed[:show_sheet_tabs].nil?
      wb_props[:tab_ratio] = wb_parsed[:tab_ratio] unless wb_parsed[:tab_ratio].nil?
      wb_props[:full_calc_on_load] = wb_parsed[:full_calc_on_load] unless wb_parsed[:full_calc_on_load].nil?
      wb_props[:code_name] = wb_parsed[:code_name] if wb_parsed[:code_name]
      wb_unmapped[:workbook_properties] = wb_props unless wb_props.empty?
      wb_unmapped[:non_worksheet_sheets] = non_worksheet_sheets unless non_worksheet_sheets.empty?

      if workbook_xml&.include?("<workbookProtection")
        prot_match = workbook_xml.match(%r{<workbookProtection\b([^>]*)/>})
        if prot_match
          prot = {}
          prot[:lock_structure] = true if prot_match[1].include?('lockStructure="1"') || prot_match[1].include?('lockStructure="true"')
          prot[:lock_windows] = true if prot_match[1].include?('lockWindows="1"') || prot_match[1].include?('lockWindows="true"')
          wb_unmapped[:workbook_protection] = prot unless prot.empty?
        end
      end

      core_xml = zip_reader.read_entry("docProps/core.xml")
      if core_xml && !core_xml.empty?
        c_listener = Ooxml::Reader::CorePropertiesListener.new
        Ooxml::XmlParser.parse(core_xml, c_listener)
        wb_unmapped[:core_properties] = c_listener.properties unless c_listener.properties.empty?
      end

      app_xml = zip_reader.read_entry("docProps/app.xml")
      if app_xml && !app_xml.empty?
        a_listener = Ooxml::Reader::AppPropertiesListener.new
        Ooxml::XmlParser.parse(app_xml, a_listener)
        wb_unmapped[:app_properties] = a_listener.properties unless a_listener.properties.empty?
      end

      custom_xml = zip_reader.read_entry("docProps/custom.xml")
      if custom_xml && !custom_xml.empty?
        cu_listener = Ooxml::Reader::CustomPropertiesListener.new
        Ooxml::XmlParser.parse(custom_xml, cu_listener)
        wb_unmapped[:custom_properties] = cu_listener.properties unless cu_listener.properties.empty?
      end

      Elements::Workbook.new(
        sheets: sheets,
        shared_strings: shared_strings,
        styles: styles,
        unmapped_data: wb_unmapped,
        defined_names: defined_names
      )
    ensure
      zip_reader&.close
    end
  end

  # Writes an XLSX file or IO stream (streaming or in-memory), or returns a binary string.
  #
  # @overload write(?target, password: nil, encryption_mode: :standard, strict_excel_mode: true, &block)
  #   Streaming write: yields a {StreamWriter} context for high-speed, zero-allocation XLSX generation.
  #   If target is omitted, returns the generated XLSX as an in-memory binary String.
  #   @param target [String, IO, StringIO, nil] Optional destination file path or writable IO object.
  #   @param password [String, nil] Optional password to encrypt the generated XLSX file.
  #   @param encryption_mode [Symbol] Encryption algorithm (:standard or :agile).
  #   @param strict_excel_mode [Boolean] Whether to enforce Microsoft Excel specification limits.
  #   @yield [stream_writer]
  #   @yieldparam stream_writer [Xlsxrb::StreamWriter]
  #   @return [String, void] Binary data if target is omitted, or void if written to target.
  #
  # @overload write(workbook, password: nil, encryption_mode: :standard)
  #   In-memory write: exports the workbook to an in-memory binary String.
  #   @param workbook [Elements::Workbook] The workbook to write.
  #   @param password [String, nil] Optional password to encrypt the binary string.
  #   @param encryption_mode [Symbol] Encryption algorithm (:standard or :agile).
  #   @return [String] Binary data representing the XLSX file.
  #
  # @overload write(target, workbook, password: nil, encryption_mode: :standard)
  #   In-memory write: writes the workbook to a file path or IO stream.
  #   @param target [String, IO, StringIO] Destination file path or writable IO object.
  #   @param workbook [Elements::Workbook] The workbook to write.
  #   @param password [String, nil] Optional password to encrypt the output file.
  #   @param encryption_mode [Symbol] Encryption algorithm (:standard or :agile).
  #   @return [void]
  #
  # @example Streaming write to file
  #   Xlsxrb.write("output.xlsx") do |writer|
  #     writer.sheet("Sheet1") { |s| s.row(["Hello", "World"]) }
  #   end
  #
  # @example Streaming write to in-memory binary string (target omitted)
  #   binary_data = Xlsxrb.write do |writer|
  #     writer.sheet("Sheet1") { |s| s.row(["Hello", "World"]) }
  #   end
  #
  # @example Password-protected streaming write
  #   Xlsxrb.write("protected.xlsx", password: "SecretPassword123") do |writer|
  #     writer.sheet("Confidential") { |s| s.row(["Private Data", 100]) }
  #   end
  #
  # @example In-memory export to binary string (for Rails send_data & mailers)
  #   binary_data = Xlsxrb.write(workbook)
  #
  # @api public
  # @param shared_strings [Boolean] Whether to intern strings in a shared strings table (`xl/sharedStrings.xml`).
  #   When set to `false`, strings are emitted directly as inline strings (`<c t="inlineStr">`), providing flat
  #   $O(1)$ constant memory during streaming write. Disabling shared strings is recommended when streaming
  #   large datasets where cells contain unique text (e.g. log exports, transaction details).
  # @param compression_level [Integer, nil] ZIP deflate compression level (0..9, or Zlib::DEFAULT_COMPRESSION
  #   for zlib's level 6). Defaults to {DEFAULT_COMPRESSION_LEVEL}. Higher levels write smaller files more slowly;
  #   level 0 stores entries without compression. See docs/adr/0009-default-zip-compression-level.md.
  #: (?untyped target_or_workbook, ?Elements::Workbook | untyped workbook_or_nil, ?password: String?, ?encryption_mode: Symbol, ?strict_excel_mode: bool, ?escape_formulas: bool, ?default_font: untyped?, ?shared_strings: bool, ?compression_level: Integer | nil) ?{ (StreamWriter) -> void } -> untyped
  def self.write(target_or_workbook = OMITTED_TARGET, workbook_or_nil = nil, password: nil, encryption_mode: :standard, strict_excel_mode: true, escape_formulas: false, default_font: nil, shared_strings: true, compression_level: DEFAULT_COMPRESSION_LEVEL, &)
    if block_given?
      if target_or_workbook.equal?(OMITTED_TARGET)
        io = StringIO.new
        io.binmode
        write(io, password: password, encryption_mode: encryption_mode, strict_excel_mode: strict_excel_mode, escape_formulas: escape_formulas, default_font: default_font, shared_strings: shared_strings, compression_level: compression_level, &)
        return io.string.b
      end

      target = target_or_workbook
      raise Error, "target is required" if target.nil?

      if password && !password.empty?
        buf = StringIO.new
        buf.binmode
        stream_writer = StreamWriter.new(buf, strict_excel_mode: strict_excel_mode, escape_formulas: escape_formulas, default_font: default_font, shared_strings: shared_strings, compression_level: compression_level)
        begin
          yield stream_writer
          stream_writer.close
          plain_bytes = buf.string.b
          encrypted_bytes = Ooxml::Crypto.encrypt(plain_bytes, password, mode: encryption_mode)
          if target.is_a?(String)
            File.binwrite(target, encrypted_bytes)
          elsif target.respond_to?(:write)
            target.write(encrypted_bytes)
          end
        ensure
          stream_writer.cleanup!
        end
      else
        stream_writer = StreamWriter.new(target, strict_excel_mode: strict_excel_mode, escape_formulas: escape_formulas, default_font: default_font, shared_strings: shared_strings, compression_level: compression_level)
        begin
          yield stream_writer
          stream_writer.close
        ensure
          stream_writer.cleanup!
        end
      end
      return
    end

    if workbook_or_nil.nil?
      raise Error, "workbook must be an Elements::Workbook" if target_or_workbook.equal?(OMITTED_TARGET)

      wb = target_or_workbook
      raise Error, "workbook must be an Elements::Workbook" unless wb.is_a?(Elements::Workbook)

      io = StringIO.new
      io.binmode
      write(io, wb, password: password, encryption_mode: encryption_mode, shared_strings: shared_strings, compression_level: compression_level)
      return io.string.b
    end

    target = target_or_workbook
    workbook = workbook_or_nil
    raise Error, "target is required" if target.nil?
    raise Error, "workbook must be an Elements::Workbook" unless workbook.is_a?(Elements::Workbook)

    if shared_strings
      sst = []
      sst_index = {}

      if workbook.shared_strings && !workbook.shared_strings.empty?
        workbook.shared_strings.each do |val|
          if (val.is_a?(String) || val.is_a?(Elements::RichText)) && !sst_index.key?(val)
            sst << val
            sst_index[val] = sst.size - 1
          end
        end
      end
    else
      sst = nil
      sst_index = nil
    end

    # Collect shared strings and build index without allocating new Hashes
    sheet_data = workbook.sheets.map do |raw_ws|
      ws = raw_ws.respond_to?(:load) ? raw_ws.load : raw_ws
      if sst_index
        ws.rows.each do |row|
          row.cells.each do |cell|
            val = cell.value
            if (val.is_a?(String) || val.is_a?(Elements::RichText)) && !sst_index.key?(val)
              sst << val
              sst_index[val] = sst.size - 1
            end
          end
        end
      end
      columns = ws.columns.map do |col|
        # simplecov:disable
        # Edge case / untested delegation block
        style_idx = col.unmapped_data&.[](:style_index) || col.unmapped_data&.[](:style) || (col.respond_to?(:style_index) && col.style_index)
        min = col.unmapped_data&.[](:min)
        max = col.unmapped_data&.[](:max)
        { index: col.index, width: col.width, hidden: col.hidden?, custom_width: col.custom_width?, outline_level: col.outline_level, style_index: style_idx, min: min, max: max }
        # simplecov:enable
      end
      sd = { name: ws.name, rows: ws.rows, columns: columns }
      sd[:charts] = ws.charts unless ws.charts.empty?
      sd[:conditional_formats] = ws.conditional_formats unless ws.conditional_formats.empty?
      sd[:data_validations] = ws.data_validations unless ws.data_validations.empty?

      sd[:state] = ws.state if ws.respond_to?(:state) && ws.state && ws.state != :visible
      sd[:comments] = ws.comments if ws.respond_to?(:comments) && !ws.comments.empty?
      if ws.respond_to?(:hyperlinks) && !ws.hyperlinks.empty?
        sd[:hyperlinks] = if ws.hyperlinks.is_a?(Hash)
                            ws.hyperlinks.map do |cell_ref, link|
                              entry = if link.is_a?(Hash)
                                        link.dup
                                      else
                                        (link.respond_to?(:to_h) ? link.to_h : { url: link })
                                      end
                              entry[:ref] ||= cell_ref
                              entry
                            end
                          else
                            ws.hyperlinks
                          end
      end
      sd[:sparkline_groups] = ws.sparkline_groups if ws.respond_to?(:sparkline_groups) && !ws.sparkline_groups.empty?
      sd[:sheet_properties] = ws.sheet_properties if ws.respond_to?(:sheet_properties) && ws.sheet_properties && !ws.sheet_properties.empty?
      sd[:sheet_format] = ws.sheet_format if ws.respond_to?(:sheet_format) && ws.sheet_format && !ws.sheet_format.empty?
      sd[:protected_ranges] = ws.protected_ranges.map { |r| r.respond_to?(:to_h) ? r.to_h : r } if ws.respond_to?(:protected_ranges) && !ws.protected_ranges.empty?
      sd[:cell_watches] = ws.cell_watches.map { |w| w.respond_to?(:to_h) ? w.to_h : w } if ws.respond_to?(:cell_watches) && !ws.cell_watches.empty?
      sd[:ignored_errors] = ws.ignored_errors if ws.respond_to?(:ignored_errors) && !ws.ignored_errors.empty?
      sd[:phonetic_properties] = ws.phonetic_properties.respond_to?(:to_h) ? ws.phonetic_properties.to_h : ws.phonetic_properties if ws.respond_to?(:phonetic_properties) && ws.phonetic_properties
      sd[:sheet_calc_pr] = ws.sheet_calc_properties.respond_to?(:to_h) ? ws.sheet_calc_properties.to_h : ws.sheet_calc_properties if ws.respond_to?(:sheet_calc_properties) && ws.sheet_calc_properties
      sd[:freeze_pane] = ws.freeze_pane.respond_to?(:to_h) ? ws.freeze_pane.to_h : ws.freeze_pane if ws.respond_to?(:freeze_pane) && ws.freeze_pane
      sd[:split_pane] = ws.unmapped_data.dig(:facade, :split_pane) || (ws.split_pane.respond_to?(:to_h) ? ws.split_pane.to_h : ws.split_pane) if ws.respond_to?(:split_pane) && ws.split_pane
      sd[:sheet_view] = ws.sheet_view.respond_to?(:to_h) ? ws.sheet_view.to_h : ws.sheet_view if ws.respond_to?(:sheet_view) && ws.sheet_view && !ws.sheet_view.empty?
      if ws.respond_to?(:selection) && ws.selection
        sd[:selection] = if ws.selection.is_a?(Array)
                           ws.selection.map { |sel| sel.respond_to?(:to_h) ? sel.to_h : sel }
                         else
                           ws.selection.respond_to?(:to_h) ? ws.selection.to_h : ws.selection
                         end
      end
      sd[:selection] ||= ws.sheet_view[:selection] if ws.respond_to?(:sheet_view) && ws.sheet_view.is_a?(Hash) && ws.sheet_view[:selection]
      sd[:page_setup] = ws.page_setup.respond_to?(:to_h) ? ws.page_setup.to_h : ws.page_setup if ws.respond_to?(:page_setup) && ws.page_setup && !ws.page_setup.empty?
      sd[:page_margins] = ws.page_margins.respond_to?(:to_h) ? ws.page_margins.to_h : ws.page_margins if ws.respond_to?(:page_margins) && ws.page_margins && !ws.page_margins.empty?
      sd[:header_footer] = ws.header_footer.respond_to?(:to_h) ? ws.header_footer.to_h : ws.header_footer if ws.respond_to?(:header_footer) && ws.header_footer && !ws.header_footer.empty?
      sd[:print_options] = ws.print_options.respond_to?(:to_h) ? ws.print_options.to_h : ws.print_options if ws.respond_to?(:print_options) && ws.print_options && !ws.print_options.empty?
      sd[:sheet_protection] = ws.sheet_protection.respond_to?(:to_h) ? ws.sheet_protection.to_h : ws.sheet_protection if ws.respond_to?(:sheet_protection) && ws.sheet_protection && !ws.sheet_protection.empty?
      sd[:tables] = ws.tables.map { |t| t.respond_to?(:to_h) ? t.to_h : t } if ws.respond_to?(:tables) && ws.tables && !ws.tables.empty?
      sd[:merge_cells] = ws.merged_cells if ws.respond_to?(:merged_cells) && ws.merged_cells && !ws.merged_cells.empty?
      sd[:auto_filter] = ws.auto_filter if ws.respond_to?(:auto_filter) && ws.auto_filter
      sd[:filter_columns] = ws.unmapped_data.dig(:facade, :filter_columns) || (ws.filter_columns.transform_values { |fc| fc.respond_to?(:to_h) ? fc.to_h : fc } if ws.respond_to?(:filter_columns) && ws.filter_columns && !ws.filter_columns.empty?)
      sd[:sort_state] = ws.unmapped_data.dig(:facade, :sort_state) || (ws.sort_state.respond_to?(:to_h) ? ws.sort_state.to_h : ws.sort_state) if ws.respond_to?(:sort_state) && ws.sort_state
      sd[:pivot_tables] = ws.pivot_tables.map { |pt| pt.respond_to?(:to_h) ? pt.to_h : pt } if ws.respond_to?(:pivot_tables) && ws.pivot_tables && !ws.pivot_tables.empty?
      sd[:shapes] = ws.shapes.map { |s| s.respond_to?(:to_h) ? s.to_h : s } if ws.respond_to?(:shapes) && ws.shapes && !ws.shapes.empty?
      sd[:threaded_comments] = ws.threaded_comments if ws.respond_to?(:threaded_comments) && ws.threaded_comments && !ws.threaded_comments.empty?
      sd[:header_footer_images] = ws.header_footer_images if ws.respond_to?(:header_footer_images) && ws.header_footer_images && !ws.header_footer_images.empty?
      if ws.respond_to?(:row_breaks) && ws.row_breaks && !ws.row_breaks.empty?
        sd[:row_breaks] = ws.row_breaks.map { |r| r.respond_to?(:to_h) ? r.to_h : r }
      end
      if ws.respond_to?(:column_breaks) && ws.column_breaks && !ws.column_breaks.empty?
        sd[:column_breaks] = ws.column_breaks.map { |c| c.respond_to?(:to_h) ? c.to_h : c }
      end
      sd[:background_image] = ws.background_image if ws.respond_to?(:background_image) && ws.background_image
      if ws.respond_to?(:images) && ws.images && !ws.images.empty?
        sd[:images] = ws.images.map do |img|
          if img.is_a?(Elements::Image)
            ext = img.filename ? File.extname(img.filename).delete_prefix(".").downcase : "png"
            opts = img.unmapped_data ? img.unmapped_data.dup : {}
            opts[:name] ||= img.name
            opts[:file_data] = img.data
            opts[:ext] = ext
            opts[:from_column] = img.from_column
            opts[:from_row] = img.from_row
            opts[:to_column] = img.to_column
            opts[:to_row] = img.to_row
            opts
          else
            img
          end
        end
      end
      sd[:shapes] = ws.shapes if ws.respond_to?(:shapes) && ws.shapes && !ws.shapes.empty?

      # Extract facade metadata from unmapped_data as fallback
      facade = ws.unmapped_data[:facade] if ws.respond_to?(:unmapped_data) && ws.unmapped_data.is_a?(Hash)
      facade&.each { |key, val| sd[key] = val unless sd.key?(key) }

      sd
    end

    # Extract workbook-level facade metadata
    wb_facade = workbook.unmapped_data[:facade] || {}
    wb_properties = (workbook.unmapped_data[:workbook_properties] || wb_facade[:workbook_properties] || {}).merge(
      workbook.respond_to?(:workbook_properties) && workbook.workbook_properties ? workbook.workbook_properties.to_h : {}
    )
    wb_core = (workbook.respond_to?(:core_properties) && workbook.core_properties) || wb_facade[:core_properties]
    wb_app = (workbook.respond_to?(:app_properties) && workbook.app_properties) || wb_facade[:app_properties]
    wb_custom = (workbook.respond_to?(:custom_properties) && workbook.custom_properties) || wb_facade[:custom_properties]
    wb_prot = (workbook.respond_to?(:workbook_protection) && workbook.workbook_protection) || wb_facade[:workbook_protection]
    raw_dns = wb_facade[:defined_names] || (workbook.respond_to?(:defined_names) && !workbook.defined_names.empty? ? workbook.defined_names : nil)
    defined_names = if raw_dns
                      raw_dns.map { |dn| dn.is_a?(Elements::MetadataRecord) ? dn.to_h : dn.dup }
                    else
                      []
                    end

    # Auto-generate _xlnm.Print_Area and _xlnm.Print_Titles from worksheet definitions if not already present
    workbook.sheets.each_with_index do |ws, idx|
      sheet_name = ws.name || "Sheet#{idx + 1}"
      if ws.respond_to?(:print_area) && ws.print_area && defined_names.none? { |dn| dn[:name] == "_xlnm.Print_Area" && (dn[:local_sheet_id] == idx || dn[:local_sheet_name] == sheet_name) }
        val = "'#{sheet_name}'!#{DslHelpers.absolute_range(ws.print_area)}"
        defined_names << { name: "_xlnm.Print_Area", value: val, local_sheet_id: idx }
      end
      next unless ws.respond_to?(:print_titles) && ws.print_titles && defined_names.none? { |dn| dn[:name] == "_xlnm.Print_Titles" && (dn[:local_sheet_id] == idx || dn[:local_sheet_name] == sheet_name) }

      pt = ws.print_titles
      pt = pt.to_h if pt.is_a?(Elements::MetadataRecord)
      parts = []
      col_val = pt[:columns]
      parts << "'#{sheet_name}'!$#{col_val.sub(":", ":$")}" if col_val
      parts << "'#{sheet_name}'!$#{pt[:rows].sub(":", ":$")}" if pt[:rows]
      defined_names << { name: "_xlnm.Print_Titles", value: parts.join(","), local_sheet_id: idx } unless parts.empty?
    end
    defined_names = nil if defined_names.empty?

    if password && !password.empty?
      buf = StringIO.new
      buf.binmode
      Ooxml::WorkbookWriter.write(
        buf,
        sheets: sheet_data,
        shared_strings: sst,
        shared_strings_index: sst_index,
        styles: workbook.styles,
        defined_names: defined_names,
        core_properties: wb_core,
        app_properties: wb_app,
        custom_properties: wb_custom,
        workbook_protection: wb_prot,
        workbook_properties: wb_properties,
        compression_level: compression_level
      )
      encrypted_bytes = Ooxml::Crypto.encrypt(buf.string.b, password, mode: encryption_mode)
      if target.is_a?(String)
        File.binwrite(target, encrypted_bytes)
      elsif target.respond_to?(:write)
        target.write(encrypted_bytes)
      end
    else
      Ooxml::WorkbookWriter.write(
        target,
        sheets: sheet_data,
        shared_strings: sst,
        shared_strings_index: sst_index,
        styles: workbook.styles,
        defined_names: defined_names,
        core_properties: wb_core,
        app_properties: wb_app,
        custom_properties: wb_custom,
        workbook_protection: wb_prot,
        workbook_properties: wb_properties,
        compression_level: compression_level
      )
    end
  end

  # Builds an in-memory {Elements::Workbook} using a declarative DSL.
  #
  # @example Build an in-memory workbook
  #   workbook = Xlsxrb.build do |builder|
  #     builder.sheet("Overview") do |sheet|
  #       sheet.row(["Title", "Date"])
  #       sheet.row(["Report", Date.today])
  #     end
  #   end
  #
  # @param strict_excel_mode [Boolean] Whether to enforce Microsoft Excel specification limits.
  # @param escape_formulas [Boolean] Whether to treat strings starting with formula prefixes as plain text.
  # @param default_font [Hash, nil] Optional default font options.
  # @yield [builder]
  # @yieldparam builder [Xlsxrb::WorkbookBuilder]
  # @return [Elements::Workbook] The compiled in-memory workbook.
  # @api public
  #: (?strict_excel_mode: bool, ?escape_formulas: bool, ?default_font: untyped?) ?{ (WorkbookBuilder) -> void } -> Elements::Workbook
  def self.build(strict_excel_mode: true, escape_formulas: false, default_font: nil)
    raise Error, "block is required" unless block_given?

    builder = WorkbookBuilder.new(strict_excel_mode: strict_excel_mode, escape_formulas: escape_formulas, default_font: default_font)
    yield builder
    builder.build
  end

  class << self
    private

    def resolve_workbook_entry(zip_reader)
      rels_xml = zip_reader.read_entry("_rels/.rels")
      if rels_xml
        entries = Ooxml::RelationshipsParser.parse_entries(rels_xml)
        doc_rel = entries.find { |e| OFFICE_DOC_REL_TYPES.include?(e[:type]) }
        if doc_rel
          clean = doc_rel[:target].delete_prefix("/")
          return clean if zip_reader.entry?(clean)
        end
      end
      return "xl/workbook.xml" if zip_reader.entry?("xl/workbook.xml")
      return "workbook.xml" if zip_reader.entry?("workbook.xml")

      "xl/workbook.xml"
    end

    def parse_workbook_relationships(rels_xml)
      return [{}, {}] if rels_xml.nil? || rels_xml.empty?

      targets = Ooxml::RelationshipsParser.parse(rels_xml)
      types = {}
      Ooxml::RelationshipsParser.parse_entries(rels_xml).each do |entry|
        types[entry[:id]] = entry[:type] if entry[:id] && entry[:type]
      end
      [targets, types]
    end

    def non_worksheet_relationship?(rel_type, target_path)
      return true if rel_type && NON_WORKSHEET_REL_SUFFIXES.any? { |suffix| rel_type.end_with?(suffix) }

      if target_path
        clean_target = target_path.delete_prefix("/")
        return true if NON_WORKSHEET_TARGET_PREFIXES.any? { |prefix| clean_target.start_with?(prefix) }
      end

      false
    end

    def resolve_part_path(zip_reader, target_path, wb_dir)
      return nil unless target_path

      clean = target_path.delete_prefix("/")
      candidates = [
        wb_dir == "." ? clean : "#{wb_dir}/#{clean}",
        clean,
        "xl/#{clean}"
      ]
      candidates.find { |c| zip_reader.entry?(c) }
    end

    def resolve_sheet_entry(zip_reader, target_path, wb_dir, index)
      resolved = resolve_part_path(zip_reader, target_path, wb_dir)
      return resolved if resolved

      fallbacks = [
        wb_dir == "." ? "worksheets/sheet#{index + 1}.xml" : "#{wb_dir}/worksheets/sheet#{index + 1}.xml",
        wb_dir == "." ? "sheet#{index + 1}.xml" : "#{wb_dir}/sheet#{index + 1}.xml",
        "xl/worksheets/sheet#{index + 1}.xml",
        "xl/sheet#{index + 1}.xml",
        "worksheets/sheet#{index + 1}.xml",
        "sheet#{index + 1}.xml"
      ]
      fallbacks.find { |f| zip_reader.entry?(f) }
    end

    def resolve_part_entry(zip_reader, rels, rel_type_sub, fallback_name, wb_dir)
      target = rels.find { |_id, t| t.include?(rel_type_sub) }&.last
      resolved = resolve_part_path(zip_reader, target, wb_dir)
      return resolved if resolved

      fallbacks = [
        wb_dir == "." ? fallback_name : "#{wb_dir}/#{fallback_name}",
        "xl/#{fallback_name}",
        fallback_name
      ]
      fallbacks.find { |f| zip_reader.entry?(f) }
    end

    #: (String name, String? sheet_xml, Array[String] shared_strings, untyped styles, ?state: Symbol, ?zip_reader: Ooxml::ZipReader?, ?entry_name: String?, ?hyperlinks: Hash[String, untyped]?, ?comments: Array[Hash[Symbol, untyped]]?, ?images: Array[Elements::Image]?, ?date1904: bool, ?dimension: String?, ?trim_empty_rows: bool, ?pad_empty_rows: bool, ?pad_empty_cells: bool, ?merged_cells: Array[String]?, ?auto_filter: String?, ?tables: Array[Hash[Symbol, untyped]]?, ?shapes: Array[Hash[Symbol, untyped]]?, ?freeze_pane: Hash[Symbol, untyped]?, ?page_setup: Hash[Symbol, untyped]?, ?page_margins: Hash[Symbol, untyped]?, ?header_footer: Hash[Symbol, untyped]?, ?print_options: Hash[Symbol, untyped]?, ?sheet_protection: Hash[Symbol, untyped]?, ?sheet_view: Hash[Symbol, untyped]?, ?sheet_properties: Hash[Symbol, untyped]?, ?sheet_format: Hash[Symbol, untyped]?, ?row_breaks: Array[untyped]?, ?column_breaks: Array[untyped]?, ?col_breaks: Array[untyped]?, ?protected_ranges: Array[Hash[Symbol, untyped]]?, ?cell_watches: Array[String]?, ?ignored_errors: Array[Hash[Symbol, untyped]]?, ?phonetic_properties: Hash[Symbol, untyped]?, ?selection: (Hash[Symbol, untyped] | Array[Hash[Symbol, untyped]])?, ?sheet_calc_properties: Hash[Symbol, untyped]?, ?background_image: Hash[Symbol, untyped]?, ?threaded_comments: Array[Hash[Symbol, untyped]]?, ?columns: Array[Elements::Column]?, ?pivot_tables: Array[Hash[Symbol, untyped]]?, ?split_pane: Hash[Symbol, untyped]?, ?filter_columns: Hash[Integer, Hash[Symbol, untyped]]?, ?sort_state: Hash[Symbol, untyped]?, ?header_footer_images: (Array[Elements::HeaderFooterImage] | Array[Hash[Symbol, untyped]])?, ?charts: (Array[Elements::Chart] | Array[Hash[Symbol, untyped]])?, ?sparkline_groups: (Array[Elements::SparklineGroup] | Array[Hash[Symbol, untyped]])?, ?print_area: String?, ?print_titles: (Elements::PrintTitles | Hash[Symbol, untyped])?) -> Elements::Worksheet
    def build_worksheet(name, sheet_xml, shared_strings, styles, state: :visible, zip_reader: nil, entry_name: nil, hyperlinks: nil, comments: nil, images: nil, date1904: false, dimension: nil, trim_empty_rows: false, pad_empty_rows: false, pad_empty_cells: false, merged_cells: nil, auto_filter: nil, tables: nil, shapes: nil, freeze_pane: nil, page_setup: nil, page_margins: nil, header_footer: nil, print_options: nil, sheet_protection: nil, sheet_view: nil, sheet_properties: nil, sheet_format: nil, row_breaks: nil, column_breaks: nil, protected_ranges: nil, cell_watches: nil, ignored_errors: nil, phonetic_properties: nil, selection: nil, sheet_calc_properties: nil, background_image: nil, threaded_comments: nil, columns: nil, pivot_tables: nil, split_pane: nil, filter_columns: nil, sort_state: nil, header_footer_images: nil, charts: nil, sparkline_groups: nil, print_area: nil, print_titles: nil)
      sheet_dimension = dimension || sheet_xml&.slice(/<(?:[a-zA-Z0-9_]+:)?dimension\b[^>]*\bref=["']([^"']+)["']/, 1)
      if sheet_xml.nil? || sheet_xml.empty?
        return Elements::Worksheet.new(
          name: name, state: state, styles: styles, images: images || [], date1904: date1904,
          dimension: sheet_dimension, trim_empty_rows: trim_empty_rows, pad_empty_rows: pad_empty_rows, pad_empty_cells: pad_empty_cells,
          columns: columns || [],
          charts: charts || [],
          sparkline_groups: sparkline_groups || [],
          print_area: print_area,
          print_titles: print_titles,
          merged_cells: merged_cells || [], auto_filter: auto_filter,
          tables: tables || [], shapes: shapes || [], freeze_pane: freeze_pane,
          split_pane: split_pane, pivot_tables: pivot_tables || [],
          filter_columns: filter_columns || {}, sort_state: sort_state,
          header_footer_images: header_footer_images || [],
          page_setup: page_setup || {}, page_margins: page_margins, header_footer: header_footer || {},
          print_options: print_options || {}, sheet_protection: sheet_protection, sheet_view: sheet_view || {},
          sheet_properties: sheet_properties || {}, sheet_format: sheet_format || {}, row_breaks: row_breaks || [], column_breaks: column_breaks || [],
          protected_ranges: protected_ranges || [], cell_watches: cell_watches || [], ignored_errors: ignored_errors || [],
          phonetic_properties: phonetic_properties, selection: selection, sheet_calc_properties: sheet_calc_properties,
          background_image: background_image,
          threaded_comments: threaded_comments || []
        )
      end

      sheet_hyperlinks = hyperlinks || resolve_hyperlinks(sheet_xml, zip_reader: zip_reader, entry_name: entry_name)
      sheet_comments = comments || resolve_comments(zip_reader: zip_reader, entry_name: entry_name)
      sheet_threaded_comments = threaded_comments || resolve_threaded_comments(zip_reader: zip_reader, entry_name: entry_name)
      sheet_images = images || resolve_images(zip_reader: zip_reader, entry_name: entry_name)
      raw_rows = Ooxml::WorksheetParser.parse(sheet_xml, shared_strings: shared_strings, styles: styles)
      raw_columns = Ooxml::WorksheetParser.parse_columns(sheet_xml)

      rows = raw_rows.map { |rr| build_row_from_raw(rr, sheet_hyperlinks, sheet_comments, styles, date1904: date1904, threaded_comments: sheet_threaded_comments) }
      sheet_columns = columns || raw_columns.map do |rc|
        # Columns from OOXML are 1-based min/max ranges; convert to 0-based
        col_unmapped = {}
        col_unmapped[:style_index] = rc[:style_index] if rc[:style_index]
        col_unmapped[:min] = rc[:min] if rc[:min]
        col_unmapped[:max] = rc[:max] if rc[:max]
        Elements::Column.new(
          index: (rc[:min] || 1) - 1,
          width: rc[:width],
          hidden: rc[:hidden] || false,
          custom_width: rc[:custom_width] || false,
          outline_level: rc[:outline_level],
          style_index: rc[:style_index],
          unmapped_data: col_unmapped
        )
      end

      cfs = if sheet_xml.include?("<conditionalFormatting")
              listener = Ooxml::Reader::ConditionalFormattingListener.new
              Ooxml::XmlParser.parse(sheet_xml, listener)
              listener.rules
            else
              []
            end

      dvs = if sheet_xml.include?("<dataValidation")
              listener = Ooxml::Reader::DataValidationsListener.new
              Ooxml::XmlParser.parse(sheet_xml, listener)
              listener.validations
            else
              []
            end

      sheet_merged_cells = merged_cells || if sheet_xml.include?("<mergeCell")
                                             listener = Ooxml::Reader::MergeCellsListener.new
                                             Ooxml::XmlParser.parse(sheet_xml, listener)
                                             listener.ranges
                                           else
                                             []
                                           end

      sheet_auto_filter = auto_filter
      sheet_filter_columns = filter_columns
      if sheet_auto_filter.nil? && sheet_xml.include?("<autoFilter")
        af_listener = Ooxml::Reader::AutoFilterListener.new
        Ooxml::XmlParser.parse(sheet_xml, af_listener)
        sheet_auto_filter = af_listener.ref
        sheet_filter_columns ||= af_listener.filter_columns
      end
      sheet_filter_columns ||= {}

      sheet_sort_state = sort_state || if sheet_xml.include?("<sortState")
                                         ss_listener = Ooxml::Reader::SortStateListener.new
                                         Ooxml::XmlParser.parse(sheet_xml, ss_listener)
                                         ss_listener.sort_state
                                       end

      sheet_tables = tables || resolve_tables(zip_reader: zip_reader, entry_name: entry_name)
      sheet_pivot_tables = pivot_tables || resolve_pivot_tables(zip_reader: zip_reader, entry_name: entry_name)
      sheet_shapes = shapes || resolve_shapes(zip_reader: zip_reader, entry_name: entry_name)

      sheet_view_data = if freeze_pane.nil? || split_pane.nil? || sheet_view.nil? || selection.nil?
                          if sheet_xml.include?("<sheetView")
                            v_listener = Ooxml::Reader::SheetViewListener.new
                            Ooxml::XmlParser.parse(sheet_xml, v_listener)
                            { view: v_listener.view, pane: v_listener.pane, selection: v_listener.selection }
                          else
                            { view: {}, pane: nil, selection: nil }
                          end
                        else
                          { view: sheet_view, pane: freeze_pane || split_pane, selection: selection }
                        end
      sheet_freeze_pane = freeze_pane || (sheet_view_data[:pane] if sheet_view_data[:pane] && sheet_view_data[:pane][:state] == :frozen)
      sheet_split_pane = split_pane || (sheet_view_data[:pane] if sheet_view_data[:pane] && sheet_view_data[:pane][:state] != :frozen)
      sheet_view_val = sheet_view || sheet_view_data[:view] || {}
      sheet_selection = selection || sheet_view_data[:selection]

      sheet_prot = sheet_protection || if sheet_xml.include?("<sheetProtection")
                                         prot_listener = Ooxml::Reader::SheetProtectionListener.new
                                         Ooxml::XmlParser.parse(sheet_xml, prot_listener)
                                         prot_listener.protection
                                       end

      sheet_props = sheet_properties || if sheet_xml.include?("<sheetPr") || sheet_xml.include?("<sheetCalcPr")
                                          pr_listener = Ooxml::Reader::SheetPropertiesListener.new
                                          Ooxml::XmlParser.parse(sheet_xml, pr_listener)
                                          pr_listener.properties
                                        else
                                          {}
                                        end

      sheet_fmt = sheet_format || if sheet_xml.include?("<sheetFormatPr")
                                    fmt_listener = Ooxml::Reader::SheetFormatListener.new
                                    Ooxml::XmlParser.parse(sheet_xml, fmt_listener)
                                    fmt_listener.properties
                                  else
                                    {}
                                  end

      print_data = if page_setup.nil? || page_margins.nil? || header_footer.nil? || print_options.nil? || row_breaks.nil? || column_breaks.nil?
                     if sheet_xml.match?(/<(?:pageSetup|pageMargins|headerFooter|printOptions|rowBreaks|colBreaks)\b/)
                       pg_listener = Ooxml::Reader::PrintPageListener.new
                       Ooxml::XmlParser.parse(sheet_xml, pg_listener)
                       {
                         page_margins: pg_listener.page_margins,
                         page_setup: pg_listener.page_setup,
                         header_footer: pg_listener.header_footer,
                         print_options: pg_listener.print_options,
                         row_breaks: pg_listener.row_breaks,
                         column_breaks: pg_listener.column_breaks
                       }
                     else
                       { page_margins: nil, page_setup: {}, header_footer: {}, print_options: {}, row_breaks: [], column_breaks: [] }
                     end
                   else
                     { page_margins: page_margins, page_setup: page_setup, header_footer: header_footer, print_options: print_options, row_breaks: row_breaks, column_breaks: column_breaks }
                   end

      sheet_prot_ranges = protected_ranges || if sheet_xml.include?("<protectedRanges")
                                                pr_listener = Ooxml::Reader::ProtectedRangesListener.new
                                                Ooxml::XmlParser.parse(sheet_xml, pr_listener)
                                                pr_listener.ranges
                                              else
                                                []
                                              end

      sheet_watches = cell_watches || if sheet_xml.include?("<cellWatches")
                                        cw_listener = Ooxml::Reader::CellWatchesListener.new
                                        Ooxml::XmlParser.parse(sheet_xml, cw_listener)
                                        cw_listener.watches
                                      else
                                        []
                                      end

      sheet_ignored_errors = ignored_errors || if sheet_xml.include?("<ignoredErrors")
                                                 ie_listener = Ooxml::Reader::IgnoredErrorsListener.new
                                                 Ooxml::XmlParser.parse(sheet_xml, ie_listener)
                                                 ie_listener.errors
                                               else
                                                 []
                                               end

      sheet_phonetic = phonetic_properties || if sheet_xml.include?("<phoneticPr")
                                                pp_listener = Ooxml::Reader::PhoneticPrListener.new
                                                Ooxml::XmlParser.parse(sheet_xml, pp_listener)
                                                pp_listener.phonetic_pr
                                              end

      sheet_charts = charts || resolve_charts(zip_reader: zip_reader, entry_name: entry_name)
      sheet_sparkline_groups = sparkline_groups || (sheet_xml.include?("sparklineGroup") ? resolve_sparkline_groups(sheet_xml) : [])
      sheet_hf_images = header_footer_images || (sheet_xml.include?("legacyDrawingHF") ? resolve_header_footer_images(sheet_xml, zip_reader: zip_reader, entry_name: entry_name) : [])

      Elements::Worksheet.new(
        name: name,
        rows: rows,
        columns: sheet_columns,
        charts: sheet_charts,
        sparkline_groups: sheet_sparkline_groups,
        print_area: print_area,
        print_titles: print_titles,
        conditional_formats: cfs,
        data_validations: dvs,
        state: state,
        hyperlinks: sheet_hyperlinks,
        comments: sheet_comments,
        images: sheet_images,
        styles: styles,
        date1904: date1904,
        dimension: sheet_dimension,
        trim_empty_rows: trim_empty_rows,
        pad_empty_rows: pad_empty_rows,
        pad_empty_cells: pad_empty_cells,
        merged_cells: sheet_merged_cells,
        auto_filter: sheet_auto_filter,
        tables: sheet_tables,
        pivot_tables: sheet_pivot_tables,
        shapes: sheet_shapes,
        freeze_pane: sheet_freeze_pane,
        split_pane: sheet_split_pane,
        filter_columns: sheet_filter_columns,
        sort_state: sheet_sort_state,
        header_footer_images: sheet_hf_images,
        page_setup: page_setup || print_data[:page_setup] || {},
        page_margins: page_margins || print_data[:page_margins],
        header_footer: header_footer || print_data[:header_footer] || {},
        print_options: print_options || print_data[:print_options] || {},
        sheet_protection: sheet_prot,
        sheet_view: sheet_view_val,
        sheet_properties: sheet_props,
        sheet_format: sheet_fmt,
        row_breaks: row_breaks || print_data[:row_breaks] || [],
        column_breaks: column_breaks || print_data[:column_breaks] || [],
        protected_ranges: sheet_prot_ranges,
        cell_watches: sheet_watches,
        ignored_errors: sheet_ignored_errors,
        phonetic_properties: sheet_phonetic,
        selection: sheet_selection,
        sheet_calc_properties: sheet_calc_properties,
        background_image: background_image || (sheet_xml.include?("<picture") ? resolve_background_image(sheet_xml, zip_reader: zip_reader, entry_name: entry_name) : nil),
        threaded_comments: sheet_threaded_comments
      )
    end

    #: (String? sheet_xml, ?zip_reader: Ooxml::ZipReader?, ?entry_name: String?, ?rels_xml: String?) -> Hash[String, Hash[Symbol, untyped]]
    def resolve_hyperlinks(sheet_xml, zip_reader: nil, entry_name: nil, rels_xml: nil)
      return {} if sheet_xml.nil? || !sheet_xml.include?("hyperlink")

      links = Ooxml::WorksheetParser.parse_hyperlinks(sheet_xml)
      return {} if links.empty?

      rid_to_url = {}
      if rels_xml && !rels_xml.empty?
        rid_to_url = Ooxml::RelationshipsParser.parse(rels_xml)
      elsif zip_reader && entry_name
        rels_path = entry_name.sub(%r{([^/]+)$}, '_rels/\1.rels')
        rels_content = zip_reader.read_entry(rels_path)
        rid_to_url = Ooxml::RelationshipsParser.parse(rels_content) if rels_content && !rels_content.empty?
      end

      result = {}
      links.each do |link|
        entry = {}
        if link[:rid]
          url = rid_to_url[link[:rid]]
          entry[:url] = url if url
        end
        entry[:display] = link[:display] if link[:display]
        entry[:tooltip] = link[:tooltip] if link[:tooltip]
        entry[:location] = link[:location] if link[:location]
        next if entry.empty?

        result[link[:ref]] = Elements::Hyperlink.new(
          url: entry[:url],
          location: entry[:location],
          tooltip: entry[:tooltip],
          display: entry[:display]
        )
      end
      result
    end

    #: (String base_dir, String target) -> String
    def resolve_relative_path(base_dir, target)
      if target.start_with?("/")
        target.delete_prefix("/")
      elsif target.start_with?("..")
        parts = base_dir.split("/") + target.split("/")
        resolved = []
        parts.each { |p| p == ".." ? resolved.pop : resolved << p }
        resolved.join("/")
      else
        "#{base_dir}/#{target}"
      end
    end

    #: (?zip_reader: Ooxml::ZipReader?, ?entry_name: String?) -> Array[Hash[Symbol, untyped]]
    def resolve_comments(zip_reader: nil, entry_name: nil)
      return [] unless zip_reader && entry_name

      rels_path = entry_name.sub(%r{([^/]+)$}, '_rels/\1.rels')
      rels_xml = zip_reader.read_entry(rels_path)
      return [] if rels_xml.nil? || rels_xml.empty?

      parser = REXML::Parsers::SAX2Parser.new(rels_xml)
      listener = Ooxml::Reader::RelsListener.new
      parser.listen(listener)
      parser.parse

      rel = listener.relationships.find { |r| r[:type]&.end_with?("/comments") }
      return [] unless rel

      base_dir = File.dirname(entry_name)
      comments_path = resolve_relative_path(base_dir, rel[:target])

      xml = zip_reader.read_entry(comments_path)
      return [] if xml.nil? || xml.empty?

      c_parser = REXML::Parsers::SAX2Parser.new(xml)
      c_listener = Ooxml::Reader::CommentsListener.new
      c_parser.listen(c_listener)
      c_parser.parse
      c_listener.comments
    end

    #: (?zip_reader: Ooxml::ZipReader?, ?entry_name: String?) -> Array[Hash[Symbol, untyped]]
    def resolve_threaded_comments(zip_reader: nil, entry_name: nil)
      return [] unless zip_reader && entry_name

      rels_path = entry_name.sub(%r{([^/]+)$}, '_rels/\1.rels')
      rels_xml = zip_reader.read_entry(rels_path)
      return [] if rels_xml.nil? || rels_xml.empty?

      parser = REXML::Parsers::SAX2Parser.new(rels_xml)
      listener = Ooxml::Reader::RelsListener.new
      parser.listen(listener)
      parser.parse

      rel = listener.relationships.find { |r| r[:type]&.end_with?("/threadedComment") }
      return [] unless rel

      base_dir = File.dirname(entry_name)
      tc_path = resolve_relative_path(base_dir, rel[:target])

      xml = zip_reader.read_entry(tc_path)
      return [] if xml.nil? || xml.empty?

      tc_parser = REXML::Parsers::SAX2Parser.new(xml)
      tc_listener = Ooxml::Reader::ThreadedCommentsListener.new
      tc_parser.listen(tc_listener)
      tc_parser.parse

      persons_map = resolve_persons_map(zip_reader: zip_reader)
      tc_listener.threaded_comments.map do |tc|
        tc_author = persons_map[tc[:person_id]]
        tc[:author] = tc_author if tc_author && !tc[:author]
        tc
      end
    end

    #: (?zip_reader: Ooxml::ZipReader?) -> Hash[String, String]
    def resolve_persons_map(zip_reader: nil)
      return {} unless zip_reader

      xml = zip_reader.read_entry("xl/persons/person.xml")
      if xml.nil? || xml.empty?
        wb_rels_xml = zip_reader.read_entry("xl/_rels/workbook.xml.rels")
        if wb_rels_xml && !wb_rels_xml.empty?
          parser = REXML::Parsers::SAX2Parser.new(wb_rels_xml)
          listener = Ooxml::Reader::RelsListener.new
          parser.listen(listener)
          parser.parse
          p_rel = listener.relationships.find { |r| r[:type]&.end_with?("/person") }
          if p_rel
            p_path = resolve_relative_path("xl", p_rel[:target])
            xml = zip_reader.read_entry(p_path)
          end
        end
      end
      return {} if xml.nil? || xml.empty?

      p_parser = REXML::Parsers::SAX2Parser.new(xml)
      p_listener = Ooxml::Reader::PersonListListener.new
      p_parser.listen(p_listener)
      p_parser.parse
      p_listener.persons.to_h { |p| [p[:id], p[:display_name]] }
    end

    #: (?zip_reader: Ooxml::ZipReader?, ?entry_name: String?) -> Array[Elements::Image]
    def resolve_images(zip_reader: nil, entry_name: nil)
      return [] unless zip_reader && entry_name

      rels_path = entry_name.sub(%r{([^/]+)$}, '_rels/\1.rels')
      rels_xml = zip_reader.read_entry(rels_path)
      return [] if rels_xml.nil? || rels_xml.empty?

      parser = REXML::Parsers::SAX2Parser.new(rels_xml)
      listener = Ooxml::Reader::RelsListener.new
      parser.listen(listener)
      parser.parse

      drawing_rel = listener.relationships.find { |r| r[:type]&.end_with?("/drawing") }
      return [] unless drawing_rel

      base_dir = File.dirname(entry_name)
      target = drawing_rel[:target]
      drawing_path = resolve_relative_path(base_dir, target)
      drawing_xml = zip_reader.read_entry(drawing_path)
      return [] if drawing_xml.nil? || drawing_xml.empty?

      drawing_rels_path = drawing_path.sub(%r{([^/]+)$}, '_rels/\1.rels')
      drawing_rels_xml = zip_reader.read_entry(drawing_rels_path)
      drawing_rels = {}
      if drawing_rels_xml && !drawing_rels_xml.empty?
        d_parser = REXML::Parsers::SAX2Parser.new(drawing_rels_xml)
        d_listener = Ooxml::Reader::RelsListener.new
        d_parser.listen(d_listener)
        d_parser.parse
        d_listener.relationships.each do |r|
          drawing_rels[r[:id]] = r[:target]
        end
      end

      p_parser = REXML::Parsers::SAX2Parser.new(drawing_xml)
      img_listener = Ooxml::Reader::DrawingImagesListener.new
      p_parser.listen(img_listener)
      p_parser.parse

      drawing_dir = File.dirname(drawing_path)
      img_listener.images.map do |raw_img|
        rid = raw_img[:embed_rid]
        media_target = drawing_rels[rid]
        img_data = nil
        filename = nil
        if media_target
          media_path = resolve_relative_path(drawing_dir, media_target)
          filename = File.basename(media_path)
          img_data = zip_reader.read_entry(media_path)
        end

        from_col = raw_img[:from_column] || raw_img[:from_col] || 0
        from_row = raw_img[:from_row] || 0
        to_col = raw_img[:to_column] || raw_img[:to_col]
        to_row = raw_img[:to_row]
        cx = raw_img[:cx]
        cy = raw_img[:cy]
        name = raw_img[:name]

        Elements::Image.new(
          filename: filename || "image.png",
          name: name,
          data: img_data,
          from_column: from_col,
          from_row: from_row,
          to_column: to_col,
          to_row: to_row,
          cx: cx,
          cy: cy,
          unmapped_data: raw_img
        )
      end
    end

    #: (String? sheet_xml, ?zip_reader: Ooxml::ZipReader?, ?entry_name: String?) -> Hash[Symbol, untyped]?
    def resolve_background_image(sheet_xml, zip_reader: nil, entry_name: nil)
      return nil unless sheet_xml && zip_reader && entry_name
      return nil unless sheet_xml.include?("<picture")

      m = sheet_xml.match(/<(?:[a-zA-Z0-9_]+:)?picture\b[^>]*\br:id=["']([^"']+)["']/)
      return nil unless m

      picture_rid = m[1]
      rels_path = entry_name.sub(%r{([^/]+)$}, '_rels/\1.rels')
      rels_xml = zip_reader.read_entry(rels_path)
      return nil if rels_xml.nil? || rels_xml.empty?

      parser = REXML::Parsers::SAX2Parser.new(rels_xml)
      listener = Ooxml::Reader::RelsListener.new
      parser.listen(listener)
      parser.parse

      pic_rel = listener.relationships.find { |r| r[:id] == picture_rid }
      return nil unless pic_rel

      base_dir = File.dirname(entry_name)
      target_path = resolve_relative_path(base_dir, pic_rel[:target])
      data = zip_reader.read_entry(target_path)
      return nil unless data

      ext = File.extname(target_path).delete_prefix(".").downcase
      mime = case ext
             when "png" then "image/png"
             when "jpg", "jpeg" then "image/jpeg"
             when "gif" then "image/gif"
             when "bmp" then "image/bmp"
             when "emf" then "image/x-emf"
             when "wmf" then "image/x-wmf"
             else "application/octet-stream"
             end

      {
        r_id: picture_rid,
        path: target_path,
        extension: ext,
        data: data.b,
        content_type: mime
      }.freeze
    end

    #: (?zip_reader: Ooxml::ZipReader?, ?entry_name: String?) -> Array[Hash[Symbol, untyped]]
    def resolve_tables(zip_reader: nil, entry_name: nil)
      return [] unless zip_reader && entry_name

      rels_path = entry_name.sub(%r{([^/]+)$}, '_rels/\1.rels')
      rels_xml = zip_reader.read_entry(rels_path)
      return [] if rels_xml.nil? || rels_xml.empty?

      parser = REXML::Parsers::SAX2Parser.new(rels_xml)
      listener = Ooxml::Reader::RelsListener.new
      parser.listen(listener)
      parser.parse

      table_rels = listener.relationships.select { |r| r[:type]&.end_with?("/table") }
      return [] if table_rels.empty?

      base_dir = File.dirname(entry_name)
      table_rels.filter_map do |rel|
        tbl_path = resolve_relative_path(base_dir, rel[:target])
        tbl_xml = zip_reader.read_entry(tbl_path)
        next if tbl_xml.nil? || tbl_xml.empty?

        t_parser = REXML::Parsers::SAX2Parser.new(tbl_xml)
        t_listener = Ooxml::Reader::TableListener.new
        t_parser.listen(t_listener)
        t_parser.parse
        t_listener.table
      end
    end

    #: (?zip_reader: Ooxml::ZipReader?, ?entry_name: String?) -> Array[Hash[Symbol, untyped]]
    def resolve_shapes(zip_reader: nil, entry_name: nil)
      return [] unless zip_reader && entry_name

      rels_path = entry_name.sub(%r{([^/]+)$}, '_rels/\1.rels')
      rels_xml = zip_reader.read_entry(rels_path)
      return [] if rels_xml.nil? || rels_xml.empty?

      parser = REXML::Parsers::SAX2Parser.new(rels_xml)
      listener = Ooxml::Reader::RelsListener.new
      parser.listen(listener)
      parser.parse

      drawing_rel = listener.relationships.find { |r| r[:type]&.end_with?("/drawing") }
      return [] unless drawing_rel

      base_dir = File.dirname(entry_name)
      drawing_path = resolve_relative_path(base_dir, drawing_rel[:target])
      drawing_xml = zip_reader.read_entry(drawing_path)
      return [] if drawing_xml.nil? || drawing_xml.empty?

      d_parser = REXML::Parsers::SAX2Parser.new(drawing_xml)
      d_listener = Ooxml::Reader::DrawingShapesListener.new
      d_parser.listen(d_listener)
      d_parser.parse
      d_listener.shapes
    end

    #: (?zip_reader: Ooxml::ZipReader?, ?entry_name: String?) -> Array[Hash[Symbol, untyped]]
    def resolve_pivot_tables(zip_reader: nil, entry_name: nil)
      return [] unless zip_reader && entry_name

      rels_path = entry_name.sub(%r{([^/]+)$}, '_rels/\1.rels')
      rels_xml = zip_reader.read_entry(rels_path)
      return [] if rels_xml.nil? || rels_xml.empty?

      parser = REXML::Parsers::SAX2Parser.new(rels_xml)
      listener = Ooxml::Reader::RelsListener.new
      parser.listen(listener)
      parser.parse

      pivot_rels = listener.relationships.select { |r| r[:type]&.end_with?("/pivotTable") }
      return [] if pivot_rels.empty?

      base_dir = File.dirname(entry_name)
      pivot_rels.filter_map do |rel|
        pt_path = resolve_relative_path(base_dir, rel[:target])
        pt_xml = zip_reader.read_entry(pt_path)
        next if pt_xml.nil? || pt_xml.empty?

        p_parser = REXML::Parsers::SAX2Parser.new(pt_xml)
        p_listener = Ooxml::Reader::PivotTableListener.new
        p_parser.listen(p_listener)
        p_parser.parse
        pt = p_listener.pivot_table
        next unless pt

        pt_rels_path = pt_path.sub(%r{([^/]+)$}, '_rels/\1.rels')
        pt_rels_xml = zip_reader.read_entry(pt_rels_path)
        if pt_rels_xml && !pt_rels_xml.empty?
          c_parser = REXML::Parsers::SAX2Parser.new(pt_rels_xml)
          c_listener = Ooxml::Reader::RelsListener.new
          c_parser.listen(c_listener)
          c_parser.parse
          cache_rel = c_listener.relationships.find { |r| r[:type]&.end_with?("/pivotCacheDefinition") }
          if cache_rel
            c_path = resolve_relative_path(File.dirname(pt_path), cache_rel[:target])
            c_xml = zip_reader.read_entry(c_path)
            if c_xml && !c_xml.empty?
              cd_parser = REXML::Parsers::SAX2Parser.new(c_xml)
              cd_listener = Ooxml::Reader::PivotCacheDefinitionListener.new
              cd_parser.listen(cd_listener)
              cd_parser.parse
              cache_info = cd_listener.cache_definition
              if cache_info
                cd_rels_path = c_path.sub(%r{([^/]+)$}, '_rels/\1.rels')
                cd_rels_xml = zip_reader.read_entry(cd_rels_path)
                if cd_rels_xml && !cd_rels_xml.empty?
                  rec_parser = REXML::Parsers::SAX2Parser.new(cd_rels_xml)
                  rec_rels_listener = Ooxml::Reader::RelsListener.new
                  rec_parser.listen(rec_rels_listener)
                  rec_parser.parse
                  rec_rel = rec_rels_listener.relationships.find { |r| r[:type]&.end_with?("/pivotCacheRecords") }
                  if rec_rel
                    rec_path = resolve_relative_path(File.dirname(c_path), rec_rel[:target])
                    rec_xml = zip_reader.read_entry(rec_path)
                    if rec_xml && !rec_xml.empty?
                      r_parser = REXML::Parsers::SAX2Parser.new(rec_xml)
                      r_listener = Ooxml::Reader::PivotCacheRecordsListener.new
                      r_parser.listen(r_listener)
                      r_parser.parse
                      cache_info[:records] = r_listener.records unless r_listener.records.empty?
                    end
                  end
                end
                pt[:cache] = cache_info
              end
            end
          end
        end
        pt
      end
    end

    #: (?zip_reader: Ooxml::ZipReader?, ?entry_name: String?) -> Array[Elements::Chart]
    def resolve_charts(zip_reader: nil, entry_name: nil)
      return [] unless zip_reader && entry_name

      rels_path = entry_name.sub(%r{([^/]+)$}, '_rels/\1.rels')
      rels_xml = zip_reader.read_entry(rels_path)
      return [] if rels_xml.nil? || rels_xml.empty?

      parser = REXML::Parsers::SAX2Parser.new(rels_xml)
      listener = Ooxml::Reader::RelsListener.new
      parser.listen(listener)
      parser.parse

      drawing_rel = listener.relationships.find { |r| r[:type]&.end_with?("/drawing") }
      return [] unless drawing_rel

      base_dir = File.dirname(entry_name)
      drawing_path = resolve_relative_path(base_dir, drawing_rel[:target])
      drawing_xml = zip_reader.read_entry(drawing_path)
      return [] if drawing_xml.nil? || drawing_xml.empty?

      drawing_rels_path = drawing_path.sub(%r{([^/]+)$}, '_rels/\1.rels')
      drawing_rels_xml = zip_reader.read_entry(drawing_rels_path)
      drawing_rels = {}
      if drawing_rels_xml && !drawing_rels_xml.empty?
        d_parser = REXML::Parsers::SAX2Parser.new(drawing_rels_xml)
        d_listener = Ooxml::Reader::RelsListener.new
        d_parser.listen(d_listener)
        d_parser.parse
        d_listener.relationships.each do |r|
          drawing_rels[r[:id]] = r[:target]
        end
      end

      d_parser = REXML::Parsers::SAX2Parser.new(drawing_xml)
      dc_listener = Ooxml::Reader::DrawingChartsListener.new
      d_parser.listen(dc_listener)
      d_parser.parse

      drawing_dir = File.dirname(drawing_path)
      dc_listener.charts.map do |raw_chart|
        chart = raw_chart.dup
        target = drawing_rels[chart[:rid]]
        if target
          chart[:target] = target
          chart_path = resolve_relative_path(drawing_dir, target)
          chart_xml = zip_reader.read_entry(chart_path)
          if chart_xml && !chart_xml.empty?
            cp = REXML::Parsers::SAX2Parser.new(chart_xml)
            cl = Ooxml::Reader::ChartTypeListener.new
            cp.listen(cl)
            cp.parse
            chart[:chart_type] = cl.chart_type
            chart[:type] = cl.chart_type
            chart[:title] = cl.title
            chart[:title_overlay] = cl.title_overlay unless cl.title_overlay.nil?
            chart[:title_font] = cl.title_font if cl.title_font
            chart[:title_fill_color] = cl.title_fill_color if cl.title_fill_color
            chart[:title_no_fill] = cl.title_no_fill if cl.title_no_fill
            chart[:title_line_color] = cl.title_line_color if cl.title_line_color
            chart[:title_line_width] = cl.title_line_width if cl.title_line_width
            chart[:title_line_dash] = cl.title_line_dash if cl.title_line_dash
            chart[:series] = cl.series unless cl.series.empty?
            chart[:legend] = cl.legend unless cl.legend.empty?
            if cl.legend_font
              chart[:legend] ||= {}
              chart[:legend][:font] = cl.legend_font
            end
            chart[:data_labels] = cl.data_labels unless cl.data_labels.empty?
            chart[:cat_axis_title] = cl.cat_axis_title if cl.cat_axis_title
            chart[:val_axis_title] = cl.val_axis_title if cl.val_axis_title
            chart[:cat_axis_title_font] = cl.cat_axis_title_font if cl.cat_axis_title_font
            chart[:val_axis_title_font] = cl.val_axis_title_font if cl.val_axis_title_font
            chart[:cat_axis_title_fill] = cl.cat_axis_title_fill if cl.cat_axis_title_fill
            chart[:cat_axis_title_no_fill] = cl.cat_axis_title_no_fill if cl.cat_axis_title_no_fill
            chart[:cat_axis_title_line_color] = cl.cat_axis_title_line_color if cl.cat_axis_title_line_color
            chart[:cat_axis_title_line_width] = cl.cat_axis_title_line_width if cl.cat_axis_title_line_width
            chart[:cat_axis_title_line_dash] = cl.cat_axis_title_line_dash if cl.cat_axis_title_line_dash
            chart[:val_axis_title_fill] = cl.val_axis_title_fill if cl.val_axis_title_fill
            chart[:val_axis_title_no_fill] = cl.val_axis_title_no_fill if cl.val_axis_title_no_fill
            chart[:val_axis_title_line_color] = cl.val_axis_title_line_color if cl.val_axis_title_line_color
            chart[:val_axis_title_line_width] = cl.val_axis_title_line_width if cl.val_axis_title_line_width
            chart[:val_axis_title_line_dash] = cl.val_axis_title_line_dash if cl.val_axis_title_line_dash
            chart[:title_layout] = cl.title_layout if cl.title_layout
            chart[:cat_axis_title_layout] = cl.cat_axis_title_layout if cl.cat_axis_title_layout
            chart[:val_axis_title_layout] = cl.val_axis_title_layout if cl.val_axis_title_layout
            chart[:title_rotation] = cl.title_rotation if cl.title_rotation
            chart[:cat_axis_title_rotation] = cl.cat_axis_title_rotation if cl.cat_axis_title_rotation
            chart[:val_axis_title_rotation] = cl.val_axis_title_rotation if cl.val_axis_title_rotation
            chart[:grouping] = cl.grouping if cl.grouping
            chart[:bar_dir] = cl.bar_dir if cl.bar_dir
            chart[:vary_colors] = cl.vary_colors unless cl.vary_colors.nil?
            chart[:plot_vis_only] = cl.plot_vis_only unless cl.plot_vis_only.nil?
            chart[:disp_blanks_as] = cl.disp_blanks_as if cl.disp_blanks_as
            chart[:style] = cl.style if cl.style
            chart[:auto_title_deleted] = cl.auto_title_deleted unless cl.auto_title_deleted.nil?
            chart[:rounded_corners] = cl.rounded_corners unless cl.rounded_corners.nil?
            chart[:cat_axis_tick_lbl_pos] = cl.cat_axis_tick_lbl_pos if cl.cat_axis_tick_lbl_pos
            chart[:val_axis_tick_lbl_pos] = cl.val_axis_tick_lbl_pos if cl.val_axis_tick_lbl_pos
            chart[:cat_axis_major_gridlines] = cl.cat_axis_major_gridlines if cl.cat_axis_major_gridlines
            chart[:val_axis_major_gridlines] = cl.val_axis_major_gridlines if cl.val_axis_major_gridlines
            chart[:cat_axis_minor_gridlines] = cl.cat_axis_minor_gridlines if cl.cat_axis_minor_gridlines
            chart[:val_axis_minor_gridlines] = cl.val_axis_minor_gridlines if cl.val_axis_minor_gridlines
            chart[:show_d_lbls_over_max] = cl.show_d_lbls_over_max unless cl.show_d_lbls_over_max.nil?
            chart[:cat_axis_delete] = cl.cat_axis_delete unless cl.cat_axis_delete.nil?
            chart[:val_axis_delete] = cl.val_axis_delete unless cl.val_axis_delete.nil?
            chart[:cat_axis_orientation] = cl.cat_axis_orientation if cl.cat_axis_orientation
            chart[:val_axis_orientation] = cl.val_axis_orientation if cl.val_axis_orientation
            chart[:gap_width] = cl.gap_width if cl.gap_width
            chart[:overlap] = cl.overlap if cl.overlap
            chart[:gap_depth] = cl.gap_depth if cl.gap_depth
            chart[:bar_shape] = cl.bar_shape if cl.bar_shape
            chart[:bubble_3d] = cl.bubble_3d unless cl.bubble_3d.nil?
            chart[:bubble_scale] = cl.bubble_scale if cl.bubble_scale
            chart[:show_neg_bubbles] = cl.show_neg_bubbles unless cl.show_neg_bubbles.nil?
            chart[:size_represents] = cl.size_represents if cl.size_represents
            chart[:view_3d] = cl.view_3d if cl.view_3d
            chart[:cat_axis_num_fmt] = cl.cat_axis_num_fmt if cl.cat_axis_num_fmt
            chart[:val_axis_num_fmt] = cl.val_axis_num_fmt if cl.val_axis_num_fmt
            chart[:cat_axis_major_tick_mark] = cl.cat_axis_major_tick_mark if cl.cat_axis_major_tick_mark
            chart[:cat_axis_minor_tick_mark] = cl.cat_axis_minor_tick_mark if cl.cat_axis_minor_tick_mark
            chart[:val_axis_major_tick_mark] = cl.val_axis_major_tick_mark if cl.val_axis_major_tick_mark
            chart[:val_axis_minor_tick_mark] = cl.val_axis_minor_tick_mark if cl.val_axis_minor_tick_mark
            chart[:cat_axis_crosses] = cl.cat_axis_crosses if cl.cat_axis_crosses
            chart[:val_axis_crosses] = cl.val_axis_crosses if cl.val_axis_crosses
            chart[:cat_axis_crosses_at] = cl.cat_axis_crosses_at if cl.cat_axis_crosses_at
            chart[:val_axis_crosses_at] = cl.val_axis_crosses_at if cl.val_axis_crosses_at
            chart[:cat_axis_tick_lbl_skip] = cl.cat_axis_tick_lbl_skip if cl.cat_axis_tick_lbl_skip
            chart[:cat_axis_tick_mark_skip] = cl.cat_axis_tick_mark_skip if cl.cat_axis_tick_mark_skip
            chart[:cat_axis_lbl_offset] = cl.cat_axis_lbl_offset if cl.cat_axis_lbl_offset
            chart[:cat_axis_auto] = cl.cat_axis_auto unless cl.cat_axis_auto.nil?
            chart[:cat_axis_lbl_algn] = cl.cat_axis_lbl_algn if cl.cat_axis_lbl_algn
            chart[:cat_axis_no_multi_lvl_lbl] = cl.cat_axis_no_multi_lvl_lbl unless cl.cat_axis_no_multi_lvl_lbl.nil?
            chart[:val_axis_cross_between] = cl.val_axis_cross_between if cl.val_axis_cross_between
            chart[:val_axis_major_unit] = cl.val_axis_major_unit if cl.val_axis_major_unit
            chart[:val_axis_minor_unit] = cl.val_axis_minor_unit if cl.val_axis_minor_unit
            chart[:val_axis_disp_units] = cl.val_axis_disp_units if cl.val_axis_disp_units
            chart[:cat_axis_scaling_max] = cl.cat_axis_scaling_max if cl.cat_axis_scaling_max
            chart[:cat_axis_scaling_min] = cl.cat_axis_scaling_min if cl.cat_axis_scaling_min
            chart[:val_axis_scaling_max] = cl.val_axis_scaling_max if cl.val_axis_scaling_max
            chart[:val_axis_scaling_min] = cl.val_axis_scaling_min if cl.val_axis_scaling_min
            chart[:cat_axis_log_base] = cl.cat_axis_log_base if cl.cat_axis_log_base
            chart[:val_axis_log_base] = cl.val_axis_log_base if cl.val_axis_log_base
            chart[:first_slice_ang] = cl.first_slice_ang if cl.first_slice_ang
            chart[:hole_size] = cl.hole_size if cl.hole_size
            chart[:smooth] = cl.smooth unless cl.smooth.nil?
            chart[:marker] = cl.marker unless cl.marker.nil?
            chart[:drop_lines] = cl.drop_lines unless cl.drop_lines.nil?
            chart[:hi_low_lines] = cl.hi_low_lines unless cl.hi_low_lines.nil?
            chart[:ser_lines] = cl.ser_lines unless cl.ser_lines.nil?
            chart[:up_down_bars] = cl.up_down_bars if cl.up_down_bars
            chart[:scatter_style] = cl.scatter_style if cl.scatter_style
            chart[:radar_style] = cl.radar_style if cl.radar_style
            chart[:cat_axis_pos] = cl.cat_axis_pos if cl.cat_axis_pos
            chart[:val_axis_pos] = cl.val_axis_pos if cl.val_axis_pos
            chart[:wireframe] = cl.wireframe unless cl.wireframe.nil?
            chart[:band_fmts] = cl.band_fmts if cl.band_fmts
            chart[:of_pie_type] = cl.of_pie_type if cl.of_pie_type
            chart[:split_type] = cl.split_type if cl.split_type
            chart[:split_pos] = cl.split_pos if cl.split_pos
            chart[:cust_split] = cl.cust_split if cl.cust_split&.any?
            chart[:second_pie_size] = cl.second_pie_size if cl.second_pie_size
            chart[:data_table] = cl.data_table if cl.data_table
            chart[:plot_area_fill] = cl.plot_area_fill if cl.plot_area_fill
            chart[:plot_area_line_color] = cl.plot_area_line_color if cl.plot_area_line_color
            chart[:plot_area_line_width] = cl.plot_area_line_width if cl.plot_area_line_width
            chart[:plot_area_line_dash] = cl.plot_area_line_dash if cl.plot_area_line_dash
            chart[:plot_area_no_fill] = cl.plot_area_no_fill if cl.plot_area_no_fill
            chart[:plot_area_layout] = cl.plot_area_layout if cl.plot_area_layout
            chart[:cat_axis_label_rotation] = cl.cat_axis_label_rotation if cl.cat_axis_label_rotation
            chart[:val_axis_label_rotation] = cl.val_axis_label_rotation if cl.val_axis_label_rotation
            chart[:cat_axis_font] = cl.cat_axis_font if cl.cat_axis_font
            chart[:val_axis_font] = cl.val_axis_font if cl.val_axis_font
            chart[:cat_axis_fill] = cl.cat_axis_fill if cl.cat_axis_fill
            chart[:cat_axis_no_fill] = cl.cat_axis_no_fill if cl.cat_axis_no_fill
            chart[:val_axis_fill] = cl.val_axis_fill if cl.val_axis_fill
            chart[:val_axis_no_fill] = cl.val_axis_no_fill if cl.val_axis_no_fill
            chart[:cat_axis_line_color] = cl.cat_axis_line_color if cl.cat_axis_line_color
            chart[:cat_axis_line_width] = cl.cat_axis_line_width if cl.cat_axis_line_width
            chart[:cat_axis_line_dash] = cl.cat_axis_line_dash if cl.cat_axis_line_dash
            chart[:val_axis_line_color] = cl.val_axis_line_color if cl.val_axis_line_color
            chart[:val_axis_line_width] = cl.val_axis_line_width if cl.val_axis_line_width
            chart[:val_axis_line_dash] = cl.val_axis_line_dash if cl.val_axis_line_dash
            chart[:floor] = cl.floor if cl.floor
            chart[:side_wall] = cl.side_wall if cl.side_wall
            chart[:back_wall] = cl.back_wall if cl.back_wall
            chart[:cat_axis_type] = cl.cat_axis_type if cl.cat_axis_type
            chart[:cat_axis_base_time_unit] = cl.cat_axis_base_time_unit if cl.cat_axis_base_time_unit
            chart[:cat_axis_major_time_unit] = cl.cat_axis_major_time_unit if cl.cat_axis_major_time_unit
            chart[:cat_axis_minor_time_unit] = cl.cat_axis_minor_time_unit if cl.cat_axis_minor_time_unit
            chart[:cat_axis_major_unit] = cl.cat_axis_major_unit if cl.cat_axis_major_unit
            chart[:cat_axis_minor_unit] = cl.cat_axis_minor_unit if cl.cat_axis_minor_unit
            chart[:chart_fill] = cl.chart_fill if cl.chart_fill
            chart[:chart_no_fill] = cl.chart_no_fill if cl.chart_no_fill
            chart[:chart_line_color] = cl.chart_line_color if cl.chart_line_color
            chart[:chart_line_width] = cl.chart_line_width if cl.chart_line_width
            chart[:chart_line_dash] = cl.chart_line_dash if cl.chart_line_dash
            chart[:protection] = cl.protection if cl.protection
            chart[:print_settings] = cl.print_settings if cl.print_settings
            chart[:chart_font] = cl.chart_font if cl.chart_font
            ca = {
              title: cl.cat_axis_title,
              major_gridlines: cl.cat_axis_major_gridlines,
              minor_gridlines: cl.cat_axis_minor_gridlines,
              num_fmt: cl.cat_axis_num_fmt,
              pos: cl.cat_axis_pos,
              max: cl.cat_axis_scaling_max,
              min: cl.cat_axis_scaling_min,
              major_unit: cl.cat_axis_major_unit,
              minor_unit: cl.cat_axis_minor_unit,
              delete: cl.cat_axis_delete
            }.compact
            chart[:category_axis] = ca unless ca.empty?

            va = {
              title: cl.val_axis_title,
              major_gridlines: cl.val_axis_major_gridlines,
              minor_gridlines: cl.val_axis_minor_gridlines,
              num_fmt: cl.val_axis_num_fmt,
              pos: cl.val_axis_pos,
              max: cl.val_axis_scaling_max,
              min: cl.val_axis_scaling_min,
              major_unit: cl.val_axis_major_unit,
              minor_unit: cl.val_axis_minor_unit,
              delete: cl.val_axis_delete
            }.compact
            chart[:value_axis] = va unless va.empty?

            chart[:sec_val_axis_title] = cl.sec_val_axis_title if cl.sec_val_axis_title
            chart[:sec_val_axis_major_gridlines] = cl.sec_val_axis_major_gridlines if cl.sec_val_axis_major_gridlines
            chart[:sec_val_axis_minor_gridlines] = cl.sec_val_axis_minor_gridlines if cl.sec_val_axis_minor_gridlines
            chart[:sec_val_axis_num_fmt] = cl.sec_val_axis_num_fmt if cl.sec_val_axis_num_fmt
            chart[:sec_val_axis_pos] = cl.sec_val_axis_pos if cl.sec_val_axis_pos
            chart[:sec_val_axis_scaling_max] = cl.sec_val_axis_scaling_max if cl.sec_val_axis_scaling_max
            chart[:sec_val_axis_scaling_min] = cl.sec_val_axis_scaling_min if cl.sec_val_axis_scaling_min
            chart[:sec_val_axis_major_unit] = cl.sec_val_axis_major_unit if cl.sec_val_axis_major_unit
            chart[:sec_val_axis_minor_unit] = cl.sec_val_axis_minor_unit if cl.sec_val_axis_minor_unit
            if cl.sec_val_axis_title || cl.sec_val_axis_major_gridlines || cl.sec_val_axis_scaling_max ||
               cl.sec_val_axis_scaling_min || cl.sec_val_axis_num_fmt || cl.sec_val_axis_pos
              sva = {
                title: cl.sec_val_axis_title,
                major_gridlines: cl.sec_val_axis_major_gridlines,
                minor_gridlines: cl.sec_val_axis_minor_gridlines,
                num_fmt: cl.sec_val_axis_num_fmt,
                pos: cl.sec_val_axis_pos,
                max: cl.sec_val_axis_scaling_max,
                min: cl.sec_val_axis_scaling_min,
                major_unit: cl.sec_val_axis_major_unit,
                minor_unit: cl.sec_val_axis_minor_unit
              }.compact
              chart[:secondary_value_axis] = sva unless sva.empty?
            end
          end
        end
        from_col = chart.delete(:from_col)
        to_col = chart.delete(:to_col)
        chart[:from_column] ||= from_col if from_col
        chart[:to_column] ||= to_col if to_col
        rid = chart.delete(:rid)
        chart[:r_id] ||= rid if rid
        ctype = chart.delete(:chart_type)
        chart[:type] ||= ctype if ctype
        Elements::Chart.new(**chart)
      end
    end

    #: (String? sheet_xml) -> Array[Elements::SparklineGroup]
    def resolve_sparkline_groups(sheet_xml)
      return [] if sheet_xml.nil? || !sheet_xml.include?("sparklineGroup")

      listener = Ooxml::Reader::SparklineGroupsListener.new
      parser = REXML::Parsers::SAX2Parser.new(sheet_xml)
      parser.listen(listener)
      parser.parse
      listener.groups.map do |grp|
        grp.is_a?(Elements::SparklineGroup) ? grp : Elements::SparklineGroup.new(**grp)
      end
    end

    #: (String? sheet_xml, ?zip_reader: Ooxml::ZipReader?, ?entry_name: String?) -> Array[Elements::HeaderFooterImage]
    def resolve_header_footer_images(sheet_xml, zip_reader: nil, entry_name: nil)
      return [] unless sheet_xml && zip_reader && entry_name
      return [] unless sheet_xml.include?("legacyDrawingHF")

      m = sheet_xml.match(/<(?:[a-zA-Z0-9_]+:)?legacyDrawingHF\b[^>]*\br:id=["']([^"']+)["']/)
      return [] unless m

      hf_rid = m[1]
      rels_path = entry_name.sub(%r{([^/]+)$}, '_rels/\1.rels')
      rels_xml = zip_reader.read_entry(rels_path)
      return [] if rels_xml.nil? || rels_xml.empty?

      parser = REXML::Parsers::SAX2Parser.new(rels_xml)
      listener = Ooxml::Reader::RelsListener.new
      parser.listen(listener)
      parser.parse

      vml_rel = listener.relationships.find { |r| r[:id] == hf_rid }
      return [] unless vml_rel

      base_dir = File.dirname(entry_name)
      vml_path = resolve_relative_path(base_dir, vml_rel[:target])
      vml_xml = zip_reader.read_entry(vml_path)
      return [] if vml_xml.nil? || vml_xml.empty?

      vml_rels_path = vml_path.sub(%r{([^/]+)$}, '_rels/\1.rels')
      vml_rels_xml = zip_reader.read_entry(vml_rels_path)
      vml_rels = {}
      if vml_rels_xml && !vml_rels_xml.empty?
        vr_parser = REXML::Parsers::SAX2Parser.new(vml_rels_xml)
        vr_listener = Ooxml::Reader::RelsListener.new
        vr_parser.listen(vr_listener)
        vr_parser.parse
        vr_listener.relationships.each do |r|
          vml_rels[r[:id]] = r[:target]
        end
      end

      vml_parser = REXML::Parsers::SAX2Parser.new(vml_xml)
      hf_listener = Ooxml::Reader::VmlHeaderFooterListener.new
      vml_parser.listen(hf_listener)
      vml_parser.parse

      vml_dir = File.dirname(vml_path)
      hf_listener.shapes.filter_map do |shape|
        rel_id = shape[:rel_id]
        media_target = vml_rels[rel_id]
        next unless media_target

        media_path = resolve_relative_path(vml_dir, media_target)
        img_data = zip_reader.read_entry(media_path)
        next unless img_data

        img_data = img_data.b
        ext = File.extname(media_path).delete_prefix(".").downcase
        content_type = case ext
                       when "png" then "image/png"
                       when "jpg", "jpeg" then "image/jpeg"
                       when "gif" then "image/gif"
                       when "emf" then "image/x-emf"
                       when "wmf" then "image/x-wmf"
                       else "application/octet-stream"
                       end

        Elements::HeaderFooterImage.new(
          position: shape[:id],
          r_id: rel_id,
          path: media_path,
          extension: ext,
          data: img_data,
          content_type: content_type
        )
      end
    end

    #: (Array[untyped]? defined_names, String? sheet_name, Integer? sheet_index) -> String?
    def extract_sheet_print_area(defined_names, sheet_name, sheet_index)
      return nil unless defined_names && !defined_names.empty?

      dn = defined_names.find do |d|
        d_name = d.respond_to?(:name) ? d.name : (d[:name] || d["name"])
        d_name == "_xlnm.Print_Area" && match_sheet_defined_name?(d, sheet_name, sheet_index)
      end
      return nil unless dn

      val = dn.respond_to?(:value) ? dn.value : (dn[:value] || dn["value"])
      return nil unless val && !val.empty?

      val.split(",").map { |part| part.sub(/\A(?:'(?:''|[^'])*'|[^!]+)!/, "").delete("$").strip }.join(",")
    end

    #: (Array[untyped]? defined_names, String? sheet_name, Integer? sheet_index) -> Elements::PrintTitles?
    def extract_sheet_print_titles(defined_names, sheet_name, sheet_index)
      return nil unless defined_names && !defined_names.empty?

      dn = defined_names.find do |d|
        d_name = d.respond_to?(:name) ? d.name : (d[:name] || d["name"])
        d_name == "_xlnm.Print_Titles" && match_sheet_defined_name?(d, sheet_name, sheet_index)
      end
      return nil unless dn

      val = dn.respond_to?(:value) ? dn.value : (dn[:value] || dn["value"])
      return nil unless val && !val.empty?

      rows = nil
      cols = nil
      val.split(",").each do |raw_part|
        part = raw_part.sub(/\A(?:'(?:''|[^'])*'|[^!]+)!/, "").delete("$").strip
        if part =~ /\A\d+(?::\d+)?\z/
          rows = part
        elsif part =~ /\A[A-Za-z]+(?::[A-Za-z]+)?\z/
          cols = part
        end
      end

      return nil if rows.nil? && cols.nil?

      Elements::PrintTitles.new(rows: rows, columns: cols)
    end

    #: (untyped defined_name, String? sheet_name, Integer? sheet_index) -> bool
    def match_sheet_defined_name?(defined_name, sheet_name, sheet_index)
      d_id = if defined_name.respond_to?(:local_sheet_id)
               defined_name.local_sheet_id
             else
               defined_name[:local_sheet_id] || defined_name["local_sheet_id"]
             end
      d_name = if defined_name.respond_to?(:local_sheet_name)
                 defined_name.local_sheet_name
               else
                 defined_name[:local_sheet_name] || defined_name["local_sheet_name"] ||
                   (defined_name.respond_to?(:sheet) ? defined_name.sheet : defined_name[:sheet])
               end
      d_val = defined_name.respond_to?(:value) ? defined_name.value : (defined_name[:value] || defined_name["value"])

      return true if !d_id.nil? && !sheet_index.nil? && d_id.to_i == sheet_index
      return true if d_name && sheet_name && d_name.to_s == sheet_name.to_s

      if d_id.nil? && sheet_name && d_val
        quoted = Regexp.escape(sheet_name)
        return true if d_val =~ /\A(?:'#{quoted}'|#{quoted})!/
      end
      false
    end

    #: (untyped raw_row, ?Hash[String, untyped] hyperlinks, ?(Array[untyped] | Hash[String, untyped]) comments, ?Hash[untyped, untyped]? styles, ?date1904: bool, ?threaded_comments: (Array[untyped] | Hash[String, untyped])) -> (Elements::Row | untyped)
    def build_row_from_raw(raw_row, hyperlinks = {}, comments = [], styles = nil, date1904: false, threaded_comments: [])
      return raw_row if raw_row.is_a?(Elements::Row)

      comments_map = if comments.is_a?(Hash)
                       comments.each_with_object({}) do |(k, c), acc|
                         acc[k.to_s.upcase] = c.is_a?(Elements::Comment) ? c : Elements::Comment.new(**c)
                       end
                     else
                       comments.each_with_object({}) do |c, acc|
                         obj = c.is_a?(Elements::Comment) ? c : Elements::Comment.new(**c)
                         r = obj.ref
                         acc[r.to_s.upcase] = obj if r
                       end
                     end

      threaded_comments_map = if threaded_comments.is_a?(Hash)
                                threaded_comments.each_with_object({}) do |(k, tcs), acc|
                                  arr = Array(tcs).map do |tc|
                                    tc.is_a?(Elements::ThreadedComment) ? tc : Elements::ThreadedComment.new(**tc)
                                  end
                                  acc[k.to_s.upcase] = arr
                                end
                              else
                                threaded_comments.each_with_object({}) do |c, acc|
                                  obj = c.is_a?(Elements::ThreadedComment) ? c : Elements::ThreadedComment.new(**c)
                                  r = obj.ref
                                  (acc[r.to_s.upcase] ||= []) << obj if r
                                end
                              end

      hyperlinks_map = if hyperlinks && !hyperlinks.empty?
                         hyperlinks.each_with_object({}) do |(k, hl), acc|
                           acc[k.to_s.upcase] = hl.is_a?(Elements::Hyperlink) ? hl : Elements::Hyperlink.new(**hl)
                         end
                       else
                         {}
                       end

      raw_cells = raw_row.respond_to?(:cells) ? raw_row.cells : raw_row[:cells]
      cells = raw_cells.map do |rc|
        if rc.is_a?(Elements::Cell)
          hl = hyperlinks_map[rc.ref]
          cm = comments_map[rc.ref]
          tcs = threaded_comments_map[rc.ref]
          fmt = rc.format_code || (NumberFormatter.format_code_for(rc.style_index, styles) if styles && rc.style_index)
          if hl || cm || tcs || (fmt && rc.format_code != fmt) || (date1904 && !rc.date1904?)
            rc.with(
              hyperlink: hl || rc.hyperlink,
              comment: cm || rc.comment,
              format_code: fmt || rc.format_code,
              date1904: date1904 || rc.date1904?,
              threaded_comments: (tcs ? (rc.threaded_comments + tcs) : rc.threaded_comments)
            )
          else
            rc
          end
        else
          parsed = Elements::Cell._parse_ref(rc[:ref]) if rc[:ref]
          row_idx = if parsed
                      parsed[0]
                    else
                      (raw_row.respond_to?(:index) ? raw_row.index : raw_row[:index])
                    end
          col_idx = parsed ? parsed[1] : 0
          ref = rc[:ref] || "#{Elements::Cell._column_letter(col_idx)}#{row_idx + 1}"
          ref_key = ref.upcase
          hl = hyperlinks_map[ref_key]
          cm = comments_map[ref_key]
          tcs = threaded_comments_map[ref_key]
          fmt = rc[:format_code] || (NumberFormatter.format_code_for(rc[:style_index], styles) if styles && rc[:style_index])

          val = rc[:value]
          raw_val = rc[:raw_value] || val&.to_s
          cell_errors = Elements::Cell.validate(row_idx, col_idx, val)
          if !cell_errors.empty? && rc[:source]
            cell_errors = cell_errors.map do |err|
              "#{err} (at #{rc[:source][:part]} row #{rc[:source][:row] + 1} cell #{rc[:ref] || "unknown"})"
            end
          end

          Elements::Cell.new(
            row_index: row_idx,
            column_index: col_idx,
            value: val,
            formula: rc[:formula],
            style_index: rc[:style_index],
            hyperlink: hl,
            comment: cm,
            raw_value: raw_val,
            format_code: fmt,
            errors: cell_errors,
            date1904: date1904,
            threaded_comments: tcs
          )
        end
      end
      attrs = raw_row.respond_to?(:attributes) ? raw_row.attributes : (raw_row[:attrs] || {})
      row_idx = raw_row.respond_to?(:index) ? raw_row.index : raw_row[:index]
      row_source = if raw_row.respond_to?(:source)
                     raw_row.source
                   else
                     (raw_row.is_a?(Hash) ? raw_row[:source] : nil)
                   end
      row_errors = Elements::Row.validate(row_idx, cells)
      if !row_errors.empty? && row_source
        # simplecov:disable
        # Edge case / untested delegation block
        row_errors = row_errors.map do |err|
          "#{err} (at #{row_source[:part]} row #{row_source[:row] + 1})"
          # simplecov:enable
        end
      end
      Elements::Row.new(
        index: row_idx,
        cells: cells,
        height: attrs[:height],
        hidden: attrs[:hidden] || false,
        custom_height: attrs[:custom_height] || false,
        collapsed: attrs[:collapsed] || false,
        outline_level: attrs[:outline_level],
        style_index: attrs[:style_index],
        unmapped_data: Elements::EMPTY_HASH,
        errors: row_errors
      )
    end

    #: (untyped source, String? password) -> untyped
    def prepare_source_io(source, password)
      if source.is_a?(String)
        bin_source = source.encoding == Encoding::BINARY ? source : source.b
        if bin_source.start_with?("PK\x03\x04".b) || bin_source.include?("\x00".b) || Ooxml::Cfb::Reader.cfb?(bin_source)
          if Ooxml::Cfb::Reader.cfb?(bin_source)
            decrypted_zip = Ooxml::Crypto.decrypt(bin_source, password)
            StringIO.new(decrypted_zip)
          else
            StringIO.new(bin_source)
          end
        elsif File.file?(source)
          first_bytes = begin
            File.binread(source, 8)
          rescue StandardError
            nil
          end
          if Ooxml::Cfb::Reader.cfb?(first_bytes)
            encrypted_data = File.binread(source)
            decrypted_zip = Ooxml::Crypto.decrypt(encrypted_data, password)
            StringIO.new(decrypted_zip)
          else
            source
          end
        else
          source
        end
      elsif source.respond_to?(:read) && source.respond_to?(:pos) && source.respond_to?(:seek)
        begin
          cur_pos = source.pos
          first_bytes = source.read(8)
          source.seek(cur_pos)
          if Ooxml::Cfb::Reader.cfb?(first_bytes)
            full_data = source.read
            decrypted_zip = Ooxml::Crypto.decrypt(full_data, password)
            StringIO.new(decrypted_zip)
          else
            source
          end
        rescue StandardError
          source
        end
      else
        source
      end
    end
  end
end
