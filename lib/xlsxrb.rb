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
require_relative "xlsxrb/utils"

# Modern, streaming-capable, low-memory XLSX reading and writing library for Ruby.
#
# Provides high-level facade methods ({Xlsxrb.read}, {Xlsxrb.write}, {Xlsxrb.build}, {Xlsxrb.modify})
# with full ECMA-376 OpenXML compliance, MS-OFFCRYPTO password protection, and zero core dependencies.
#
# @api public
module Xlsxrb
  # Base class for all exceptions raised by Xlsxrb.
  # Predefined cell error constants are exposed on this class for convenience.
  class Error < StandardError
    DIV0  = Elements::CellError.new(code: "#DIV/0!")
    NA    = Elements::CellError.new(code: "#N/A")
    NAME  = Elements::CellError.new(code: "#NAME?")
    NULL  = Elements::CellError.new(code: "#NULL!")
    NUM   = Elements::CellError.new(code: "#NUM!")
    REF   = Elements::CellError.new(code: "#REF!")
    VALUE = Elements::CellError.new(code: "#VALUE!")
  end

  # Raised when parsing malformed or invalid XML structures.
  class ParseError < Error; end

  OMITTED_TARGET = Object.new.freeze
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
  private_constant :OMITTED_TARGET, :OFFICE_DOC_REL_TYPES, :NON_WORKSHEET_REL_SUFFIXES, :NON_WORKSHEET_TARGET_PREFIXES

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
  # @param font_props [Hash] Inline font styling options (e.g. bold: true, color: "FF0000").
  # @return [Elements::RichText] The compiled rich text element.
  # @api public
  #: (*(Hash[Symbol, untyped] | Elements::RichTextRun | String | Array[untyped] | untyped) runs, ?text: String?, **untyped font_props) -> Elements::RichText
  def self.rich_text(*runs, text: nil, **font_props)
    if text
      runs = [Elements::RichTextRun.new(text: text, font: font_props.empty? ? nil : font_props)]
    elsif runs.size == 1 && runs.first.is_a?(String) && !font_props.empty?
      runs = [Elements::RichTextRun.new(text: runs.first, font: font_props)]
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
  # @return [Elements::Formula]
  # @api public
  #: (String expression, ?cached_value: (String | Numeric | bool | Elements::CellError | nil), ?type: (Symbol | String)?, ?ref: String?) -> Elements::Formula
  def self.formula(expression, cached_value: nil, type: nil, ref: nil)
    Elements::Formula.new(
      expression: expression,
      cached_value: cached_value,
      type: type,
      ref: ref,
      calculate_always: cached_value.nil? || nil
    )
  end

  # Creates an array or dynamic array {Elements::Formula} object.
  #
  # @example Create a multi-cell array formula
  #   formula = Xlsxrb.array_formula("SUM(B1:B10*C1:C10)", "A1:C1", cached_value: 0)
  #
  # @example Create a dynamic array formula
  #   formula = Xlsxrb.array_formula("SORT(UNIQUE(A1:A10))", "D1")
  #
  # @param expression [String] The formula expression without leading '='.
  # @param ref [String, nil] Target cell or range reference (e.g. "A1:C1" or "D1").
  # @param cached_value [Object, nil] Optional precomputed value.
  # @return [Elements::Formula]
  # @api public
  #: (String expression, ?String? ref, ?cached_value: (String | Numeric | bool | nil)) -> Elements::Formula
  def self.array_formula(expression, ref = nil, cached_value: nil)
    Elements::Formula.new(
      expression: expression,
      cached_value: cached_value,
      type: :array,
      ref: ref,
      calculate_always: cached_value.nil? || nil
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

  # Formats a cell value according to an Excel format code string.
  #
  # @param value [Object, nil] Value to format.
  # @param format_code [String, nil] Excel format code string.
  # @param date1904 [Boolean] Whether the 1904 date system is active.
  # @return [String] Formatted string representation.
  # @api public
  #: (untyped value, ?String? format_code, ?date1904: bool) -> String
  def self.format(value, format_code = nil, date1904: false)
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
    Utils.date_to_serial(date_or_time, date1904: date1904)
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
    Utils.valid_excel_date?(serial, date1904: date1904)
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
    Utils.serial_to_date(serial, date1904: date1904)
  end

  # Converts a Time or Date to a fractional Excel serial number.
  #
  # @example
  #   Xlsxrb.datetime_to_serial(Time.utc(2026, 1, 1, 12, 0, 0)) #=> 46023.5
  #
  # @param time [Time, Date] Time or Date instance.
  # @param date1904 [Boolean] Whether to use the 1904 date system.
  # @return [Float] Fractional serial number.
  # @api public
  #: (Time | Date time, ?date1904: bool) -> Float
  def self.datetime_to_serial(time, date1904: false)
    Utils.datetime_to_serial(time, date1904: date1904)
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
    Utils.serial_to_time(serial, date1904: date1904)
  end

  class << self
    alias serial_to_datetime serial_to_time
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
    Utils.parse_range(range_ref)
  end

  # Converts 0-based bounding coordinates to an A1-style range reference string.
  #
  # @example
  #   Xlsxrb.range_to_ref(0, 0, 9, 2) #=> "A1:C10"
  #
  # @param top_row [Integer]
  # @param left_col [Integer]
  # @param bottom_row [Integer]
  # @param right_col [Integer]
  # @return [String]
  # @api public
  #: (Integer top_row, Integer left_col, Integer bottom_row, Integer right_col) -> String
  def self.range_to_ref(top_row, left_col, bottom_row, right_col)
    Utils.range_to_ref(top_row, left_col, bottom_row, right_col)
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

  class << self
    alias color_to_hex to_hex_color
  end

  # Reads an XLSX file (streaming and lazy-loaded by default) from a file path, IO stream, or binary String.
  #
  # Sheets and rows are streamed lazily with O(1) constant memory. If a block is given,
  # yields each {StreamSheet} sequentially. Call {#load} on the returned Workbook or Sheet
  # to convert to an in-memory representation for coordinate random access (`sheet["A1"]`).
  #
  # @overload read(source, password: nil, trim_empty_rows: false, pad_empty_rows: false, pad_empty_cells: false, &block)
  #   Yields each {StreamSheet} sequentially in streaming mode.
  #   @param source [String, IO, StringIO] File path, binary content string, or readable IO stream.
  #   @param password [String, nil] Optional password to decrypt password-protected XLSX files.
  #   @param trim_empty_rows [Boolean] Whether to omit trailing empty rows during sheet iteration.
  #   @param pad_empty_rows [Boolean] Whether to yield empty rows for skipped row numbers.
  #   @param pad_empty_cells [Boolean] Whether to pad missing cells within rows.
  #   @yield [sheet]
  #   @yieldparam sheet [StreamSheet] The streaming worksheet object.
  #   @return [void]
  #
  # @overload read(source, password: nil, trim_empty_rows: false, pad_empty_rows: false, pad_empty_cells: false)
  #   Returns a lazy {Elements::Workbook} object.
  #   @param source [String, IO, StringIO] File path, binary content string, or readable IO stream.
  #   @param password [String, nil] Optional password to decrypt password-protected XLSX files.
  #   @param trim_empty_rows [Boolean] Whether to omit trailing empty rows during sheet iteration.
  #   @param pad_empty_rows [Boolean] Whether to yield empty rows for skipped row numbers.
  #   @param pad_empty_cells [Boolean] Whether to pad missing cells within rows.
  #   @return [Elements::Workbook]
  #
  # @example Streaming read across sheets and rows (O(1) constant memory)
  #   Xlsxrb.read("large.xlsx") do |sheet|
  #     puts "Sheet: #{sheet.name}"
  #     sheet.each_row do |row|
  #       row.each_cell { |cell| puts "#{cell.ref}: #{cell.value}" }
  #     end
  #   end
  #
  # @example Explicit in-memory loading for coordinate random access
  #   wb = Xlsxrb.read("data.xlsx")
  #   sheet = wb.sheets.first
  #   doc_sheet = sheet.load         # explicitly load into memory
  #   puts doc_sheet["A1"].value     # coordinate random access
  #
  # @param max_uncompressed_size [Integer, nil] Maximum allowed uncompressed size in bytes per entry (ZIP bomb defense).
  # @raise [ZipBombError] If an uncompressed entry exceeds max_uncompressed_size.
  # @raise [EncryptedFileError] If the file is encrypted and no password was supplied.
  # @raise [InvalidPasswordError] If the supplied password is incorrect.
  # @raise [ParseError] If the spreadsheet XML structure is invalid.
  # @api public
  #: (untyped source, ?password: String?, ?trim_empty_rows: bool, ?pad_empty_rows: bool, ?pad_empty_cells: bool, ?max_uncompressed_size: Integer?) { (StreamSheet) -> void } -> void
  #: (untyped source, ?password: String?, ?trim_empty_rows: bool, ?pad_empty_rows: bool, ?pad_empty_cells: bool, ?max_uncompressed_size: Integer?) -> Elements::Workbook
  def self.read(source, password: nil, trim_empty_rows: false, pad_empty_rows: false, pad_empty_cells: false, max_uncompressed_size: nil, &)
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
        ws_idx += 1
        next unless sheet_path

        chunk_supplier = ->(&blk) { zip_reader.each_entry_chunk(sheet_path, &blk) }

        StreamSheet.new(
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
          pad_empty_cells: pad_empty_cells
        )
      end

      wb_unmapped = {}
      wb_props = {}
      wb_props[:date1904] = true if is_date1904
      wb_props[:active_tab] = wb_parsed[:active_tab] unless wb_parsed[:active_tab].nil?
      wb_props[:full_calc_on_load] = wb_parsed[:full_calc_on_load] unless wb_parsed[:full_calc_on_load].nil?
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
      defined_names = wb_parsed[:defined_names] || []
      wb = Elements::Workbook.new(
        sheets: sheets,
        shared_strings: shared_strings,
        styles: styles,
        unmapped_data: wb_unmapped,
        defined_names: defined_names
      )

      if block_given?
        sheets.each(&)
        nil
      else
        wb
      end
    ensure
      zip_reader.close if block_given?
    end
  end

  class << self
    alias stream read
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
  #: (Elements::Workbook workbook, ?password: String?, ?encryption_mode: Symbol) -> String
  #: (untyped target, Elements::Workbook | untyped workbook, ?password: String?, ?encryption_mode: Symbol) -> void
  #: (?untyped target_or_workbook, ?Elements::Workbook | untyped workbook_or_nil, ?password: String?, ?encryption_mode: Symbol, ?strict_excel_mode: bool, ?escape_formulas: bool) ?{ (StreamWriter) -> void } -> untyped
  def self.write(target_or_workbook = OMITTED_TARGET, workbook_or_nil = nil, password: nil, encryption_mode: :standard, strict_excel_mode: true, escape_formulas: false, &)
    if block_given?
      if target_or_workbook.equal?(OMITTED_TARGET)
        io = StringIO.new
        io.binmode
        write(io, password: password, encryption_mode: encryption_mode, strict_excel_mode: strict_excel_mode, escape_formulas: escape_formulas, &)
        return io.string.b
      end

      target = target_or_workbook
      raise Error, "target is required" if target.nil?

      if password && !password.empty?
        buf = StringIO.new
        buf.binmode
        stream_writer = StreamWriter.new(buf, strict_excel_mode: strict_excel_mode, escape_formulas: escape_formulas)
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
        stream_writer = StreamWriter.new(target, strict_excel_mode: strict_excel_mode, escape_formulas: escape_formulas)
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
      write(io, wb, password: password, encryption_mode: encryption_mode)
      return io.string.b
    end

    target = target_or_workbook
    workbook = workbook_or_nil
    raise Error, "target is required" if target.nil?
    raise Error, "workbook must be an Elements::Workbook" unless workbook.is_a?(Elements::Workbook)

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

    # Collect shared strings and build index without allocating new Hashes
    sheet_data = workbook.sheets.map do |raw_ws|
      ws = raw_ws.respond_to?(:load) ? raw_ws.load : raw_ws
      ws.rows.each do |row|
        row.cells.each do |cell|
          val = cell.value
          if (val.is_a?(String) || val.is_a?(Elements::RichText)) && !sst_index.key?(val)
            sst << val
            sst_index[val] = sst.size - 1
          end
        end
      end
      columns = ws.columns.map do |col|
        # simplecov:disable
        # Edge case / untested delegation block
        style_idx = col.unmapped_data&.[](:style_index) || col.unmapped_data&.[](:style) || (col.respond_to?(:style_index) && col.style_index)
        { index: col.index, width: col.width, hidden: col.hidden, custom_width: col.custom_width, outline_level: col.outline_level, style_index: style_idx }
        # simplecov:enable
      end
      sd = { name: ws.name, rows: ws.rows, columns: columns }
      sd[:charts] = ws.charts unless ws.charts.empty?
      sd[:conditional_formats] = ws.conditional_formatting unless ws.conditional_formatting.empty?
      sd[:data_validations] = ws.data_validations unless ws.data_validations.empty?

      # Extract facade metadata from unmapped_data
      facade = ws.unmapped_data[:facade]
      facade&.each { |key, val| sd[key] ||= val }

      sd[:state] ||= ws.state if ws.respond_to?(:state) && ws.state && ws.state != :visible
      sd[:comments] ||= ws.comments if ws.respond_to?(:comments) && !ws.comments.empty?
      if ws.respond_to?(:hyperlinks) && !ws.hyperlinks.empty?
        sd[:hyperlinks] ||= if ws.hyperlinks.is_a?(Hash)
                              ws.hyperlinks.map do |cell_ref, link|
                                entry = link.is_a?(Hash) ? link.dup : { url: link }
                                entry[:ref] ||= cell_ref
                                entry
                              end
                            else
                              ws.hyperlinks
                            end
      end
      sd[:sparkline_groups] ||= ws.sparkline_groups if ws.respond_to?(:sparkline_groups) && !ws.sparkline_groups.empty?
      sd[:sheet_properties] ||= ws.sheet_properties if ws.respond_to?(:sheet_properties) && !ws.sheet_properties.empty?
      sd[:sheet_format] ||= ws.sheet_format if ws.respond_to?(:sheet_format) && !ws.sheet_format.empty?
      sd[:protected_ranges] ||= ws.protected_ranges if ws.respond_to?(:protected_ranges) && !ws.protected_ranges.empty?
      sd[:cell_watches] ||= ws.cell_watches if ws.respond_to?(:cell_watches) && !ws.cell_watches.empty?
      sd[:freeze_pane] ||= ws.freeze_pane if ws.respond_to?(:freeze_pane) && ws.freeze_pane
      sd[:split_pane] ||= ws.split_pane if ws.respond_to?(:split_pane) && ws.split_pane
      sd[:sheet_view] ||= ws.sheet_view if ws.respond_to?(:sheet_view) && ws.sheet_view && !ws.sheet_view.empty?
      sd[:page_setup] ||= ws.page_setup if ws.respond_to?(:page_setup) && ws.page_setup && !ws.page_setup.empty?
      sd[:page_margins] ||= ws.page_margins if ws.respond_to?(:page_margins) && ws.page_margins && !ws.page_margins.empty?
      sd[:header_footer] ||= ws.header_footer if ws.respond_to?(:header_footer) && ws.header_footer && !ws.header_footer.empty?
      sd[:print_options] ||= ws.print_options if ws.respond_to?(:print_options) && ws.print_options && !ws.print_options.empty?
      sd[:sheet_protection] ||= ws.sheet_protection if ws.respond_to?(:sheet_protection) && ws.sheet_protection && !ws.sheet_protection.empty?
      sd[:tables] ||= ws.tables if ws.respond_to?(:tables) && ws.tables && !ws.tables.empty?
      sd[:merge_cells] ||= ws.merged_cells if ws.respond_to?(:merged_cells) && ws.merged_cells && !ws.merged_cells.empty?
      sd[:auto_filter] ||= ws.auto_filter if ws.respond_to?(:auto_filter) && ws.auto_filter
      sd[:row_breaks] ||= ws.row_breaks if ws.respond_to?(:row_breaks) && ws.row_breaks && !ws.row_breaks.empty?
      sd[:col_breaks] ||= ws.col_breaks if ws.respond_to?(:col_breaks) && ws.col_breaks && !ws.col_breaks.empty?
      if ws.respond_to?(:images) && ws.images && !ws.images.empty?
        sd[:images] ||= ws.images.map do |img|
          if img.is_a?(Elements::Image)
            ext = img.filename ? File.extname(img.filename).delete_prefix(".").downcase : "png"
            opts = img.unmapped_data ? img.unmapped_data.dup : {}
            opts[:name] ||= img.name
            opts[:file_data] = img.data
            opts[:ext] = ext
            opts[:from_col] = img.from_col
            opts[:from_row] = img.from_row
            opts[:to_col] = img.to_col
            opts[:to_row] = img.to_row
            opts
          else
            img
          end
        end
      end
      sd[:shapes] ||= ws.shapes if ws.respond_to?(:shapes) && ws.shapes && !ws.shapes.empty?

      sd
    end

    # Extract workbook-level facade metadata
    wb_facade = workbook.unmapped_data[:facade] || {}
    wb_properties = workbook.unmapped_data[:workbook_properties] || wb_facade[:workbook_properties]
    wb_core = (workbook.respond_to?(:core_properties) && workbook.core_properties) || wb_facade[:core_properties]
    wb_app = (workbook.respond_to?(:app_properties) && workbook.app_properties) || wb_facade[:app_properties]
    wb_custom = (workbook.respond_to?(:custom_properties) && workbook.custom_properties) || wb_facade[:custom_properties]
    wb_prot = (workbook.respond_to?(:workbook_protection) && workbook.workbook_protection) || wb_facade[:workbook_protection]
    raw_dns = wb_facade[:defined_names] || (workbook.respond_to?(:defined_names) && !workbook.defined_names.empty? ? workbook.defined_names : nil)
    defined_names = raw_dns ? raw_dns.map(&:dup) : []

    # Auto-generate _xlnm.Print_Area and _xlnm.Print_Titles from worksheet definitions if not already present
    workbook.sheets.each_with_index do |ws, idx|
      sheet_name = ws.name || "Sheet#{idx + 1}"
      if ws.respond_to?(:print_area) && ws.print_area && defined_names.none? { |dn| dn[:name] == "_xlnm.Print_Area" && (dn[:local_sheet_id] == idx || dn[:local_sheet_name] == sheet_name) }
        val = "'#{sheet_name}'!#{DslHelpers.absolute_range(ws.print_area)}"
        defined_names << { name: "_xlnm.Print_Area", value: val, local_sheet_id: idx }
      end
      next unless ws.respond_to?(:print_titles) && ws.print_titles && defined_names.none? { |dn| dn[:name] == "_xlnm.Print_Titles" && (dn[:local_sheet_id] == idx || dn[:local_sheet_name] == sheet_name) }

      pt = ws.print_titles
      parts = []
      parts << "'#{sheet_name}'!$#{pt[:cols].sub(":", ":$")}" if pt[:cols]
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
        workbook_properties: wb_properties
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
        workbook_properties: wb_properties
      )
    end
  end

  # Modifies an existing XLSX file in-memory using an immutable transformation block.
  #
  # Reads the workbook, passes it to the block, and writes the resulting workbook.
  # If no target is specified, the source file is overwritten in-place.
  #
  # @example Modify a template and save to new file
  #   Xlsxrb.modify("template.xlsx", "output.xlsx") do |workbook|
  #     workbook.update_sheet("Sheet1") do |sheet|
  #       sheet.update_cell("B1", value: "Updated Title")
  #            .update_cell("B2", value: 100)
  #     end
  #   end
  #
  # @param source [String, IO, StringIO] The source file path or IO object.
  # @param target [String, IO, StringIO, nil] Target file path or IO object (overwrites source if nil).
  # @param password [String, nil] Optional password for encrypted files.
  # @yield [workbook]
  # @yieldparam workbook [Elements::Workbook] The loaded workbook.
  # @yieldreturn [Elements::Workbook] The modified workbook.
  # @return [void]
  # @api public
  #: (untyped source, ?untyped target, ?password: String?) ?{ (Elements::Workbook) -> untyped } -> void
  def self.modify(source, target = nil, password: nil)
    raise Error, "source is required" if source.nil?
    raise Error, "block is required" unless block_given?

    prepared_source = prepare_source_io(source, password)
    in_place = target.nil? || target == source || (source.is_a?(String) && target.is_a?(String) && File.file?(target) && File.identical?(source, target))
    temp_target_path = nil

    begin
      write_dest = if in_place && source.is_a?(String)
                     temp_target = Tempfile.new(["xlsxrb_mod", ".xlsx"], File.dirname(source))
                     temp_target_path = temp_target.path
                     temp_target.close
                     temp_target_path
                   elsif in_place && source.is_a?(StringIO)
                     StringIO.new
                   else
                     target || source
                   end

      replaced_entries = {}
      structural_change = false

      Ooxml::ZipReader.open(prepared_source) do |zip_reader|
        workbook_path = resolve_workbook_entry(zip_reader)
        wb_dir = File.dirname(workbook_path)
        wb_rels_name = wb_dir == "." ? "_rels/#{File.basename(workbook_path)}.rels" : "#{wb_dir}/_rels/#{File.basename(workbook_path)}.rels"
        rels_xml = zip_reader.read_entry(wb_rels_name)
        rels, rel_types = parse_workbook_relationships(rels_xml)

        sst_entry = resolve_part_entry(zip_reader, rels, "sharedStrings", "sharedStrings.xml", wb_dir)
        shared_strings_xml = sst_entry ? zip_reader.read_entry(sst_entry) : nil
        shared_strings = shared_strings_xml ? Ooxml::SharedStringsParser.parse(shared_strings_xml) : []
        sst_index = {}
        shared_strings.each_with_index { |s, idx| sst_index[s] ||= idx }
        initial_sst_size = shared_strings.size

        styles_entry = resolve_part_entry(zip_reader, rels, "styles", "styles.xml", wb_dir)
        styles_xml = styles_entry ? zip_reader.read_entry(styles_entry) : nil
        styles = styles_xml ? Ooxml::StylesParser.parse(styles_xml) : {}

        workbook_xml = zip_reader.read_entry(workbook_path)
        wb_parsed = workbook_xml ? Ooxml::WorkbookParser.parse_with_properties(workbook_xml) : { sheets: [], date1904: false, defined_names: [] }
        workbook_sheets = wb_parsed[:sheets]

        sheet_entry_map = {}
        ws_idx = 0
        loaded_sheets = workbook_sheets.filter_map do |sheet_info|
          r_id = sheet_info[:r_id]
          target_path = rels[r_id]
          next if non_worksheet_relationship?(rel_types[r_id], target_path)

          sheet_path = resolve_sheet_entry(zip_reader, target_path, wb_dir, ws_idx)
          ws_idx += 1
          next unless sheet_path

          sheet_entry_map[sheet_info[:name]] = sheet_path
          sheet_xml = zip_reader.read_entry(sheet_path) || ""
          build_worksheet(sheet_info[:name], sheet_xml, shared_strings, styles, state: sheet_info[:state] || :visible, zip_reader: zip_reader, entry_name: sheet_path)
        end

        mod_unmapped = {}
        mod_props = {}
        mod_props[:date1904] = true if wb_parsed[:date1904]
        mod_props[:active_tab] = wb_parsed[:active_tab] unless wb_parsed[:active_tab].nil?
        mod_props[:full_calc_on_load] = wb_parsed[:full_calc_on_load] unless wb_parsed[:full_calc_on_load].nil?
        mod_unmapped[:workbook_properties] = mod_props unless mod_props.empty?

        original_workbook = Elements::Workbook.new(
          sheets: loaded_sheets,
          shared_strings: shared_strings,
          styles: styles,
          unmapped_data: mod_unmapped,
          defined_names: wb_parsed[:defined_names] || []
        )
        mutable_shared_strings = shared_strings.dup

        result_workbook = yield original_workbook
        result_workbook = original_workbook unless result_workbook.is_a?(Elements::Workbook)

        structural_change = result_workbook.sheets.size != original_workbook.sheets.size ||
                            result_workbook.sheets.map(&:name) != original_workbook.sheets.map(&:name)

        if structural_change
          write_target = target || source
          write(write_target, result_workbook, password: password)
        else
          result_workbook.sheets.each_with_index do |new_sheet, idx|
            orig_sheet = original_workbook.sheets[idx]
            next if new_sheet.equal?(orig_sheet) || new_sheet == orig_sheet

            entry_name = sheet_entry_map[new_sheet.name] || "xl/worksheets/sheet#{idx + 1}.xml"
            orig_sheet_xml = zip_reader.read_entry(entry_name) || ""

            new_sheet_data = serialize_sheet_data_xml(
              new_sheet.rows,
              mutable_shared_strings,
              sst_index,
              use_sst: !shared_strings_xml.nil?
            )
            updated_sheet_xml = replace_sheet_data(orig_sheet_xml, new_sheet_data)
            updated_sheet_xml = update_dimension(updated_sheet_xml, new_sheet.rows)

            all_hyperlinks = new_sheet.hyperlinks.dup
            new_sheet.rows.each do |row|
              row.cells.each do |cell|
                all_hyperlinks[cell.ref] ||= cell.hyperlink if cell.hyperlink
              end
            end

            if all_hyperlinks != orig_sheet.hyperlinks
              rels_path = entry_name.sub(%r{([^/]+)$}, '_rels/\1.rels')
              orig_rels_xml = replaced_entries[rels_path] || zip_reader.read_entry(rels_path)
              updated_sheet_xml, updated_rels_xml = update_sheet_hyperlinks(
                updated_sheet_xml,
                orig_rels_xml,
                all_hyperlinks
              )
              replaced_entries[rels_path] = updated_rels_xml if updated_rels_xml
            end

            all_comments = new_sheet.comments.dup
            new_sheet.rows.each do |row|
              row.cells.each do |cell|
                if cell.comment && all_comments.none? { |c| (c[:ref] || c[:cell])&.to_s&.upcase == cell.ref.upcase }
                  entry = cell.comment.is_a?(Hash) ? cell.comment.merge(ref: cell.ref) : { ref: cell.ref, text: cell.comment }
                  all_comments << entry
                end
              end
            end

            if all_comments != orig_sheet.comments
              rels_path = entry_name.sub(%r{([^/]+)$}, '_rels/\1.rels')
              orig_rels_xml = replaced_entries[rels_path] || zip_reader.read_entry(rels_path)
              updated_sheet_xml, updated_rels_xml = update_sheet_comments(
                updated_sheet_xml,
                rels_path,
                orig_rels_xml,
                all_comments,
                zip_reader,
                idx,
                entry_name,
                replaced_entries
              )
              replaced_entries[rels_path] = updated_rels_xml if updated_rels_xml
            end

            replaced_entries[entry_name] = updated_sheet_xml
          end

          if mutable_shared_strings.size > initial_sst_size && shared_strings_xml
            new_strings = mutable_shared_strings[initial_sst_size..]
            sst_path = sst_entry || (wb_dir == "." ? "sharedStrings.xml" : "#{wb_dir}/sharedStrings.xml")
            replaced_entries[sst_path] = append_to_shared_strings_xml(
              shared_strings_xml,
              new_strings
            )
          end

          perform_modify_write(
            write_dest,
            zip_reader,
            replaced_entries,
            password: password
          )
        end
      end

      if in_place && !structural_change
        if source.is_a?(String) && temp_target_path
          FileUtils.mv(temp_target_path, source)
          temp_target_path = nil
        elsif source.is_a?(StringIO) && write_dest.is_a?(StringIO)
          source.string = write_dest.string
        elsif source.respond_to?(:write) && source.respond_to?(:seek) && write_dest.is_a?(StringIO)
          source.seek(0)
          source.write(write_dest.string)
          source.truncate(source.pos) if source.respond_to?(:truncate)
          source.rewind
        end
      end
    ensure
      if temp_target_path && File.exist?(temp_target_path)
        begin
          File.unlink(temp_target_path)
        rescue SystemCallError
          nil
        end
      end
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
  # @yield [builder]
  # @yieldparam builder [Xlsxrb::WorkbookBuilder]
  # @return [Elements::Workbook] The compiled in-memory workbook.
  # @api public
  #: (?strict_excel_mode: bool) ?{ (WorkbookBuilder) -> void } -> Elements::Workbook
  def self.build(strict_excel_mode: true)
    raise Error, "block is required" unless block_given?

    builder = WorkbookBuilder.new(strict_excel_mode: strict_excel_mode)
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

    #: (String name, String? sheet_xml, Array[String] shared_strings, untyped styles, ?state: Symbol, ?zip_reader: Ooxml::ZipReader?, ?entry_name: String?, ?hyperlinks: Hash[String, untyped]?, ?comments: Array[Hash[Symbol, untyped]]?, ?images: Array[Elements::Image]?, ?date1904: bool, ?dimension: String?, ?trim_empty_rows: bool, ?pad_empty_rows: bool, ?pad_empty_cells: bool, ?merged_cells: Array[String]?, ?auto_filter: String?, ?tables: Array[Hash[Symbol, untyped]]?, ?shapes: Array[Hash[Symbol, untyped]]?, ?freeze_pane: Hash[Symbol, untyped]?, ?page_setup: Hash[Symbol, untyped]?, ?page_margins: Hash[Symbol, untyped]?, ?header_footer: Hash[Symbol, untyped]?, ?print_options: Hash[Symbol, untyped]?, ?sheet_protection: Hash[Symbol, untyped]?, ?sheet_view: Hash[Symbol, untyped]?, ?sheet_properties: Hash[Symbol, untyped]?, ?sheet_format: Hash[Symbol, untyped]?, ?row_breaks: Array[untyped]?, ?col_breaks: Array[untyped]?, ?protected_ranges: Array[Hash[Symbol, untyped]]?, ?cell_watches: Array[String]?) -> Elements::Worksheet
    def build_worksheet(name, sheet_xml, shared_strings, styles, state: :visible, zip_reader: nil, entry_name: nil, hyperlinks: nil, comments: nil, images: nil, date1904: false, dimension: nil, trim_empty_rows: false, pad_empty_rows: false, pad_empty_cells: false, merged_cells: nil, auto_filter: nil, tables: nil, shapes: nil, freeze_pane: nil, page_setup: nil, page_margins: nil, header_footer: nil, print_options: nil, sheet_protection: nil, sheet_view: nil, sheet_properties: nil, sheet_format: nil, row_breaks: nil, col_breaks: nil, protected_ranges: nil, cell_watches: nil)
      sheet_dimension = dimension || sheet_xml&.slice(/<(?:[a-zA-Z0-9_]+:)?dimension\b[^>]*\bref=["']([^"']+)["']/, 1)
      if sheet_xml.nil? || sheet_xml.empty?
        return Elements::Worksheet.new(
          name: name, state: state, styles: styles, images: images || [], date1904: date1904,
          dimension: sheet_dimension, trim_empty_rows: trim_empty_rows, pad_empty_rows: pad_empty_rows, pad_empty_cells: pad_empty_cells,
          merged_cells: merged_cells || [], auto_filter: auto_filter,
          tables: tables || [], shapes: shapes || [], freeze_pane: freeze_pane,
          page_setup: page_setup || {}, page_margins: page_margins, header_footer: header_footer || {},
          print_options: print_options || {}, sheet_protection: sheet_protection, sheet_view: sheet_view || {},
          sheet_properties: sheet_properties || {}, sheet_format: sheet_format || {}, row_breaks: row_breaks || [], col_breaks: col_breaks || [],
          protected_ranges: protected_ranges || [], cell_watches: cell_watches || []
        )
      end

      sheet_hyperlinks = hyperlinks || resolve_hyperlinks(sheet_xml, zip_reader: zip_reader, entry_name: entry_name)
      sheet_comments = comments || resolve_comments(zip_reader: zip_reader, entry_name: entry_name)
      sheet_images = images || resolve_images(zip_reader: zip_reader, entry_name: entry_name)
      raw_rows = Ooxml::WorksheetParser.parse(sheet_xml, shared_strings: shared_strings, styles: styles)
      raw_columns = Ooxml::WorksheetParser.parse_columns(sheet_xml)

      rows = raw_rows.map { |rr| build_row_from_raw(rr, sheet_hyperlinks, sheet_comments, styles, date1904: date1904) }
      columns = raw_columns.map do |rc|
        # Columns from OOXML are 1-based min/max ranges; convert to 0-based
        col_unmapped = {}
        col_unmapped[:style_index] = rc[:style_index] if rc[:style_index]
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

      sheet_auto_filter = auto_filter || if sheet_xml.include?("<autoFilter")
                                           listener = Ooxml::Reader::AutoFilterListener.new
                                           Ooxml::XmlParser.parse(sheet_xml, listener)
                                           listener.ref
                                         end

      sheet_tables = tables || resolve_tables(zip_reader: zip_reader, entry_name: entry_name)
      sheet_shapes = shapes || resolve_shapes(zip_reader: zip_reader, entry_name: entry_name)

      sheet_view_data = if freeze_pane.nil? || sheet_view.nil?
                          if sheet_xml.include?("<sheetView")
                            v_listener = Ooxml::Reader::SheetViewListener.new
                            Ooxml::XmlParser.parse(sheet_xml, v_listener)
                            { view: v_listener.view, pane: v_listener.pane }
                          else
                            { view: {}, pane: nil }
                          end
                        else
                          { view: sheet_view, pane: freeze_pane }
                        end
      sheet_freeze_pane = freeze_pane || sheet_view_data[:pane]
      sheet_view_val = sheet_view || sheet_view_data[:view] || {}

      sheet_prot = sheet_protection || if sheet_xml.include?("<sheetProtection")
                                         prot_listener = Ooxml::Reader::SheetProtectionListener.new
                                         Ooxml::XmlParser.parse(sheet_xml, prot_listener)
                                         prot_listener.protection
                                       end

      sheet_props = sheet_properties || if sheet_xml.include?("<sheetPr")
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

      print_data = if page_setup.nil? || page_margins.nil? || header_footer.nil? || print_options.nil? || row_breaks.nil? || col_breaks.nil?
                     if sheet_xml.match?(/<(?:pageSetup|pageMargins|headerFooter|printOptions|rowBreaks|colBreaks)\b/)
                       pg_listener = Ooxml::Reader::PrintPageListener.new
                       Ooxml::XmlParser.parse(sheet_xml, pg_listener)
                       {
                         page_margins: pg_listener.page_margins,
                         page_setup: pg_listener.page_setup,
                         header_footer: pg_listener.header_footer,
                         print_options: pg_listener.print_options,
                         row_breaks: pg_listener.row_breaks,
                         col_breaks: pg_listener.col_breaks
                       }
                     else
                       { page_margins: nil, page_setup: {}, header_footer: {}, print_options: {}, row_breaks: [], col_breaks: [] }
                     end
                   else
                     { page_margins: page_margins, page_setup: page_setup, header_footer: header_footer, print_options: print_options, row_breaks: row_breaks, col_breaks: col_breaks }
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

      Elements::Worksheet.new(
        name: name,
        rows: rows,
        columns: columns,
        conditional_formatting: cfs,
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
        shapes: sheet_shapes,
        freeze_pane: sheet_freeze_pane,
        page_setup: page_setup || print_data[:page_setup] || {},
        page_margins: page_margins || print_data[:page_margins],
        header_footer: header_footer || print_data[:header_footer] || {},
        print_options: print_options || print_data[:print_options] || {},
        sheet_protection: sheet_prot,
        sheet_view: sheet_view_val,
        sheet_properties: sheet_props,
        sheet_format: sheet_fmt,
        row_breaks: row_breaks || print_data[:row_breaks] || [],
        col_breaks: col_breaks || print_data[:col_breaks] || [],
        protected_ranges: sheet_prot_ranges,
        cell_watches: sheet_watches
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
        result[link[:ref]] = entry unless entry.empty?
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

        from_col = raw_img[:from_col] || 0
        from_row = raw_img[:from_row] || 0
        to_col = raw_img[:to_col]
        to_row = raw_img[:to_row]
        cx = raw_img[:cx]
        cy = raw_img[:cy]
        name = raw_img[:name]

        Elements::Image.new(
          filename: filename || "image.png",
          name: name,
          data: img_data,
          from_col: from_col,
          from_row: from_row,
          to_col: to_col,
          to_row: to_row,
          cx: cx,
          cy: cy,
          unmapped_data: raw_img
        )
      end
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

    #: (untyped raw_row, ?Hash[String, Hash[Symbol, untyped]] hyperlinks, ?(Array[Hash[Symbol, untyped]] | Hash[String, Hash[Symbol, untyped]]) comments, ?Hash[untyped, untyped]? styles, ?date1904: bool) -> (Elements::Row | untyped)
    def build_row_from_raw(raw_row, hyperlinks = {}, comments = [], styles = nil, date1904: false)
      return raw_row if raw_row.is_a?(Elements::Row)

      comments_map = if comments.is_a?(Hash)
                       comments
                     else
                       comments.each_with_object({}) { |c, acc| (r = c[:ref] || c[:cell]) && (acc[r.to_s.upcase] = c) }
                     end

      raw_cells = raw_row[:cells]
      cells = raw_cells.map do |rc|
        if rc.is_a?(Elements::Cell)
          hl = hyperlinks && !hyperlinks.empty? ? hyperlinks[rc.ref] : nil
          cm = comments_map && !comments_map.empty? ? comments_map[rc.ref] : nil
          fmt = rc.format_code || (NumberFormatter.format_code_for(rc.style_index, styles) if styles && rc.style_index)
          if hl || cm || (fmt && rc.format_code != fmt) || (date1904 && !rc.date1904?)
            rc.with(hyperlink: hl || rc.hyperlink, comment: cm || rc.comment, format_code: fmt || rc.format_code, date1904: date1904 || rc.date1904?)
          else
            rc
          end
        else
          parsed = Elements::Cell.parse_ref(rc[:ref]) if rc[:ref]
          row_idx = parsed ? parsed[0] : raw_row[:index]
          col_idx = parsed ? parsed[1] : 0
          ref = rc[:ref] || "#{Elements::Cell.column_letter(col_idx)}#{row_idx + 1}"
          hl = hyperlinks ? hyperlinks[ref] : nil
          cm = comments_map ? comments_map[ref] : nil
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
            date1904: date1904
          )
        end
      end
      attrs = raw_row[:attrs] || {}
      row_errors = Elements::Row.validate(raw_row[:index], cells)
      if !row_errors.empty? && raw_row[:source]
        # simplecov:disable
        # Edge case / untested delegation block
        row_errors = row_errors.map do |err|
          "#{err} (at #{raw_row[:source][:part]} row #{raw_row[:source][:row] + 1})"
          # simplecov:enable
        end
      end
      row_unmapped = {}
      row_unmapped[:style_index] = attrs[:style_index] if attrs[:style_index]
      Elements::Row.new(
        index: raw_row[:index],
        cells: cells,
        height: attrs[:height],
        hidden: attrs[:hidden] || false,
        custom_height: attrs[:custom_height] || false,
        collapsed: attrs[:collapsed] || false,
        outline_level: attrs[:outline_level],
        unmapped_data: row_unmapped,
        errors: row_errors
      )
    end

    #: (Elements::Cell cell, Array[String] sst, Hash[String, Integer] sst_index, ?date1904: bool) -> Hash[Symbol, untyped]
    def build_raw_cell(cell, sst, sst_index, date1904: false)
      # simplecov:disable
      # Edge case / untested delegation block
      ref = cell.ref
      value = cell.value
      result = { ref: ref, style_index: cell.style_index }

      case value
      when String, Xlsxrb::Elements::RichText
        idx = sst_index[value] ||= begin
          sst << value
          sst.size - 1
        end
        result[:value] = idx
        result[:type] = "s"
      when true, false
        result[:value] = value
        result[:type] = "b"
      when Integer, Float
        result[:value] = value
      when Xlsxrb::Elements::CellError
        result[:value] = value.code
        result[:type] = "e"
      when Date
        result[:value] = Xlsxrb::Ooxml::Utils.date_to_serial(value, date1904: date1904)
      when Time
        result[:value] = Xlsxrb::Ooxml::Utils.datetime_to_serial(value, date1904: date1904)
      # simplecov:enable
      when NilClass
        # empty cell
      end

      # simplecov:disable
      # Edge case / untested delegation block
      if cell.formula
        f = cell.formula
        if f.is_a?(Elements::Formula)
          result[:formula] = f.expression
          result[:formula_type] = f.type if f.type
          result[:formula_ref] = f.ref if f.ref
          result[:formula_si] = f.shared_index if f.shared_index
          result[:formula_ca] = true if f.calculate_always
          unless f.cached_value.nil?
            # Cached value is written as-is (not through SST)
            result[:value] = f.cached_value
            case f.cached_value
            when true, false
              result[:type] = "b"
            when Elements::CellError
              result[:type] = "e"
              result[:value] = f.cached_value.code
            else
              result.delete(:type) # Ensure no type is set; cached values are plain text in <v>
            end
            # simplecov:enable
          end
        else
          # simplecov:disable
          # Edge case / untested delegation block
          result[:formula] = f
          # simplecov:enable
        end
      end
      # simplecov:disable
      # Edge case / untested delegation block
      result
      # simplecov:enable
    end

    #: (Elements::Row row) -> Hash[Symbol, untyped]
    def build_row_attrs(row)
      # simplecov:disable
      # Edge case / untested delegation block
      attrs = {}
      attrs[:height] = row.height if row.height
      attrs[:hidden] = true if row.hidden
      attrs[:custom_height] = true if row.custom_height
      attrs[:collapsed] = true if row.collapsed
      attrs[:outline_level] = row.outline_level if row.outline_level
      attrs
      # simplecov:enable
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

    #: (untyped dest, Ooxml::ZipReader zip_reader, Hash[String, String] replaced_entries, ?password: String?) -> void
    def perform_modify_write(dest, zip_reader, replaced_entries, password: nil)
      norm_replaced = {}
      replaced_entries.each do |k, v|
        norm_replaced[k.tr("\\", "/").downcase] = [k.tr("\\", "/"), v]
      end

      write_zip = lambda do |zip_out|
        Ooxml::ZipWriter.open(zip_out) do |zip_writer|
          written_keys = Set.new

          zip_reader.entry_names.each do |name|
            norm_name = name.tr("\\", "/").downcase
            next if written_keys.include?(norm_name)

            written_keys << norm_name
            canonical_path = name.tr("\\", "/")

            if norm_replaced.key?(norm_name)
              _orig_key, content = norm_replaced[norm_name]
              zip_writer.add_entry(canonical_path, content)
            else
              zip_reader.copy_to_writer(name, zip_writer, target_path: canonical_path)
            end
          end

          norm_replaced.each do |norm_name, (canonical_name, content)|
            next if written_keys.include?(norm_name)

            written_keys << norm_name
            zip_writer.add_entry(canonical_name, content)
          end
        end
      end

      if password && !password.empty?
        buf = StringIO.new
        buf.binmode
        write_zip.call(buf)
        encrypted_bytes = Ooxml::Crypto.encrypt(buf.string.b, password)
        if dest.is_a?(String)
          File.binwrite(dest, encrypted_bytes)
        elsif dest.respond_to?(:write)
          dest.write(encrypted_bytes)
        end
      else
        write_zip.call(dest)
      end
    end

    #: (String sst_xml, Array[untyped] new_strings) -> String
    def append_to_shared_strings_xml(sst_xml, new_strings)
      return sst_xml if new_strings.nil? || new_strings.empty?

      sst_match = sst_xml.match(/<([a-zA-Z0-9_]+:)?sst\b([^>]*)>/)
      prefix_with_colon = sst_match ? (sst_match[1] || "") : ""
      si_tag = "#{prefix_with_colon}si"
      t_tag = "#{prefix_with_colon}t"

      new_si_tags = +""
      new_strings.each do |str|
        s = str.to_s
        new_si_tags << "<#{si_tag}><#{t_tag}>#{Ooxml::XmlBuilder.escape(s)}</#{t_tag}></#{si_tag}>"
      end

      close_tag = "</#{prefix_with_colon}sst>"
      sst_close = sst_xml.rindex(close_tag)
      updated_xml = if sst_close
                      prefix = sst_xml[0...sst_close]
                      suffix = sst_xml[sst_close..]
                      "#{prefix}#{new_si_tags}#{suffix}"
                    else
                      self_closing = sst_xml.match(%r{<#{Regexp.escape(prefix_with_colon)}sst\b[^>]*/>})
                      return sst_xml unless self_closing

                      sc_idx = self_closing.begin(0)
                      sc_end = self_closing.end(0)
                      opening = self_closing[0].sub(%r{/>\z}, ">")
                      prefix = sst_xml[0...sc_idx]
                      suffix = sst_xml[sc_end..]
                      "#{prefix}#{opening}#{new_si_tags}#{close_tag}#{suffix}"
                    end

      updated_xml.sub(/<#{Regexp.escape(prefix_with_colon)}sst\b([^>]*)>/) do
        attrs = Regexp.last_match(1)
        if attrs =~ /\bcount="(\d+)"/
          old_count = Regexp.last_match(1).to_i
          new_count = old_count + new_strings.size
          attrs = attrs.sub(/\bcount="\d+"/, "count=\"#{new_count}\"")
        else
          attrs = "#{attrs} count=\"#{new_strings.size}\""
        end
        if attrs =~ /\buniqueCount="(\d+)"/
          old_ucount = Regexp.last_match(1).to_i
          new_ucount = old_ucount + new_strings.size
          attrs = attrs.sub(/\buniqueCount="\d+"/, "uniqueCount=\"#{new_ucount}\"")
        else
          attrs = "#{attrs} uniqueCount=\"#{new_strings.size}\""
        end
        "<#{prefix_with_colon}sst#{attrs}>"
      end
    end

    #: (String sheet_xml, String new_sheet_data_xml) -> String
    def replace_sheet_data(sheet_xml, new_sheet_data_xml)
      orig_encoding = sheet_xml.encoding
      sheet_b = sheet_xml.b
      sd_term = sheet_b.index("sheetData")
      return sheet_xml unless sd_term

      sheet_data_start = sheet_b.rindex("<", sd_term)
      return sheet_xml unless sheet_data_start

      prefix_len = sd_term - (sheet_data_start + 1)
      tag_prefix = sheet_b.byteslice(sheet_data_start + 1, prefix_len).force_encoding(orig_encoding)

      open_tag_end = sheet_b.index(">", sd_term)
      return sheet_xml unless open_tag_end

      prefix = sheet_b.byteslice(0...sheet_data_start)
      suffix = if sheet_b.getbyte(open_tag_end - 1) == 47
                 sheet_b.byteslice((open_tag_end + 1)..) || "".b
               else
                 close_tag = "</#{tag_prefix}sheetData>".b
                 close_tag_start = sheet_b.index(close_tag, open_tag_end)
                 return sheet_xml unless close_tag_start

                 sheet_b.byteslice((close_tag_start + close_tag.bytesize)..) || "".b
               end

      new_sheet_data_xml = new_sheet_data_xml.gsub(%r{</?(sheetData|row|c|v|f|is|t)\b}) { |m| m.sub(%r{(</?)}, "\\1#{tag_prefix}") } if tag_prefix && !tag_prefix.empty?

      result = prefix << new_sheet_data_xml.b << suffix
      result.force_encoding(orig_encoding)
    end

    #: (String sheet_xml, Array[Elements::Row] rows) -> String
    def update_dimension(sheet_xml, rows)
      return sheet_xml if rows.empty?

      min_row = rows.first.index
      max_row = rows.last.index
      min_col = rows.map { |r| r.cells.first&.column_index }.compact.min || 0
      max_col = rows.map { |r| r.cells.last&.column_index }.compact.max || 0
      min_col_letter = Elements::Cell.column_letter(min_col)
      max_col_letter = Elements::Cell.column_letter(max_col)
      dim_ref = if min_row == max_row && min_col == max_col
                  "#{min_col_letter}#{min_row + 1}"
                else
                  "#{min_col_letter}#{min_row + 1}:#{max_col_letter}#{max_row + 1}"
                end

      sheet_xml.sub(/<([a-zA-Z0-9_]+:)?dimension\b([^>]*)\bref="[^"]*"/, "<\\1dimension\\2ref=\"#{dim_ref}\"")
    end

    #: (String sheet_xml, String? orig_rels_xml, Hash[String, untyped] new_hyperlinks) -> [String, String?]
    def update_sheet_hyperlinks(sheet_xml, orig_rels_xml, new_hyperlinks)
      existing_rels = []
      max_rid_num = 0
      if orig_rels_xml && !orig_rels_xml.empty?
        orig_rels_xml.scan(%r{<Relationship\b([^>]*)/?>}) do |m|
          attrs = m[0]
          id = attrs[/Id="([^"]*)"/, 1]
          type = attrs[/Type="([^"]*)"/, 1]
          target = attrs[/Target="([^"]*)"/, 1]
          target_mode = attrs[/TargetMode="([^"]*)"/, 1]
          if id =~ /\ArId(\d+)\z/
            num = Regexp.last_match(1).to_i
            max_rid_num = num if num > max_rid_num
          end
          existing_rels << { id: id, type: type, target: target, target_mode: target_mode } unless type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
        end
      end

      sheet_rels = existing_rels.dup
      cur_rid = max_rid_num
      hl_elements = []

      ws_prefix = sheet_xml[/<([a-zA-Z0-9_]+:)?worksheet\b/, 1] || ""

      new_hyperlinks.each do |cell_ref, link|
        entry = link.is_a?(Hash) ? link.dup : { url: link.to_s }
        url = entry[:url]
        loc = entry[:location]
        if url&.start_with?("#")
          loc ||= url.delete_prefix("#")
          url = nil
        end
        display = entry[:display]
        tooltip = entry[:tooltip]

        attrs = if url && !url.empty?
                  cur_rid += 1
                  rid = "rId#{cur_rid}"
                  sheet_rels << {
                    id: rid,
                    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                    target: url,
                    target_mode: "External"
                  }
                  %(ref="#{cell_ref}" r:id="#{rid}")
                elsif loc && !loc.empty?
                  %(ref="#{cell_ref}" location="#{Ooxml::XmlBuilder.escape(loc)}")
                else
                  next
                end

        attrs << %( display="#{Ooxml::XmlBuilder.escape(display)}") if display
        attrs << %( tooltip="#{Ooxml::XmlBuilder.escape(tooltip)}") if tooltip
        hl_elements << "<#{ws_prefix}hyperlink #{attrs}/>"
      end

      updated_sheet_xml = sheet_xml.dup
      root_tag = updated_sheet_xml[/\A\s*(?:<\?xml[^>]*\?>\s*)?<([a-zA-Z0-9_]+:)?worksheet\b[^>]*/, 0]
      updated_sheet_xml.sub!(/<([a-zA-Z0-9_]+:)?worksheet\b/, '\0 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"') if hl_elements.any? && (root_tag.nil? || !root_tag.match?(/\bxmlns:r\s*=/))

      new_hyperlinks_xml = hl_elements.empty? ? "" : "<#{ws_prefix}hyperlinks>#{hl_elements.join}</#{ws_prefix}hyperlinks>"

      if updated_sheet_xml =~ /<([a-zA-Z0-9_]+:)?hyperlinks\b/
        updated_sheet_xml.sub!(%r{<(([a-zA-Z0-9_]+:)?hyperlinks)\b(?:[^>]*/>|[^>]*>.*?</\1>)}m) { new_hyperlinks_xml }
      elsif hl_elements.any?
        search_start = 0
        %w[sheetData customSheetViews mergeCells conditionalFormatting dataValidations].each do |tag|
          pattern = if ws_prefix.empty?
                      %r{</#{tag}>|<#{tag}\b[^>]*/\s*>}
                    else
                      %r{</#{Regexp.escape(ws_prefix)}#{tag}>|<#{Regexp.escape(ws_prefix)}#{tag}\b[^>]*/\s*>}
                    end
          if (m = updated_sheet_xml.match(pattern))
            end_pos = m.end(0)
            search_start = end_pos if end_pos > search_start
          end
        end

        tags_sub = "printOptions|pageMargins|pageSetup|headerFooter|rowBreaks|colBreaks|customProperties|cellWatches|ignoredErrors|smartTags|drawing|legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst"
        insert_pattern = if ws_prefix.empty?
                           %r{<(?:#{tags_sub})\b|</worksheet>}
                         else
                           %r{<#{Regexp.escape(ws_prefix)}(?:#{tags_sub})\b|</#{Regexp.escape(ws_prefix)}worksheet>}
                         end
        match = updated_sheet_xml.match(insert_pattern, search_start)
        if match
          updated_sheet_xml.insert(match.begin(0), new_hyperlinks_xml)
        else
          updated_sheet_xml << new_hyperlinks_xml
        end
      end

      updated_rels_xml = if sheet_rels.empty? && (orig_rels_xml.nil? || orig_rels_xml.empty?)
                           nil
                         else
                           io = StringIO.new
                           b = Ooxml::XmlBuilder.new(io)
                           b.declaration
                           b.open_tag("Relationships", { xmlns: "http://schemas.openxmlformats.org/package/2006/relationships" })
                           sheet_rels.each do |rel|
                             rel_attrs = { Id: rel[:id], Type: rel[:type], Target: rel[:target] }
                             rel_attrs[:TargetMode] = rel[:target_mode] if rel[:target_mode]
                             b.empty_tag("Relationship", rel_attrs)
                           end
                           b.close_tag("Relationships")
                           io.string
                         end

      [updated_sheet_xml, updated_rels_xml]
    end

    #: (String sheet_xml, String _rels_path, String? orig_rels_xml, Array[Hash[Symbol, untyped]] new_comments, Ooxml::ZipReader zip_reader, Integer sheet_idx, String entry_name, Hash[String, String] replaced_entries) -> [String, String?]
    def update_sheet_comments(sheet_xml, _rels_path, orig_rels_xml, new_comments, zip_reader, sheet_idx, entry_name, replaced_entries)
      existing_rels = []
      max_rid_num = 0
      if orig_rels_xml && !orig_rels_xml.empty?
        orig_rels_xml.scan(%r{<Relationship\b([^>]*)/?>}) do |m|
          attrs = m[0]
          id = attrs[/Id="([^"]*)"/, 1]
          type = attrs[/Type="([^"]*)"/, 1]
          target = attrs[/Target="([^"]*)"/, 1]
          target_mode = attrs[/TargetMode="([^"]*)"/, 1]
          if id =~ /\ArId(\d+)\z/
            num = Regexp.last_match(1).to_i
            max_rid_num = num if num > max_rid_num
          end
          existing_rels << { id: id, type: type, target: target, target_mode: target_mode }
        end
      end

      comments_rel = existing_rels.find { |r| r[:type]&.end_with?("/comments") }
      vml_rel = existing_rels.find { |r| r[:type]&.end_with?("/vmlDrawing") }

      writer = Ooxml::Writer.new
      new_comments.each do |c|
        writer.add_comment(c[:cell] || c[:ref], c[:text], author: c[:author] || "Author")
      end
      comments_xml = writer.send(:generate_comments_xml, writer.comments)
      vml_xml = writer.send(:generate_vml_drawing_xml, writer.comments)

      updated_sheet_xml = sheet_xml.dup

      if comments_rel
        base_dir = File.dirname(entry_name)
        comments_path = resolve_relative_path(base_dir, comments_rel[:target])
        replaced_entries[comments_path] = comments_xml
        if vml_rel
          vml_path = resolve_relative_path(base_dir, vml_rel[:target])
          replaced_entries[vml_path] = vml_xml
        end
      elsif new_comments.any?
        existing_comment_count = zip_reader.entry_names.count { |n| n =~ %r{\Axl/comments\d+\.xml\z} }
        comment_file_num = existing_comment_count + sheet_idx + 1
        comments_path = "xl/comments#{comment_file_num}.xml"
        vml_path = "xl/drawings/vmlDrawing#{comment_file_num}.vml"
        replaced_entries[comments_path] = comments_xml
        replaced_entries[vml_path] = vml_xml

        max_rid_num += 1
        c_rid = "rId#{max_rid_num}"
        existing_rels << {
          id: c_rid,
          type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
          target: "../comments#{comment_file_num}.xml"
        }

        max_rid_num += 1
        vml_rid = "rId#{max_rid_num}"
        existing_rels << {
          id: vml_rid,
          type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
          target: "../drawings/vmlDrawing#{comment_file_num}.vml"
        }

        ws_prefix = updated_sheet_xml[/<([a-zA-Z0-9_]+:)?worksheet\b/, 1] || ""
        legacy_pattern = ws_prefix.empty? ? /<legacyDrawing\b/ : /<#{Regexp.escape(ws_prefix)}legacyDrawing\b/
        unless updated_sheet_xml =~ legacy_pattern
          legacy_tag = %(<#{ws_prefix}legacyDrawing r:id="#{vml_rid}"/>)
          search_start = 0
          %w[sheetData customSheetViews mergeCells conditionalFormatting dataValidations hyperlinks pageMargins pageSetup].each do |tag|
            pattern = if ws_prefix.empty?
                        %r{</#{tag}>|<#{tag}\b[^>]*/\s*>}
                      else
                        %r{</#{Regexp.escape(ws_prefix)}#{tag}>|<#{Regexp.escape(ws_prefix)}#{tag}\b[^>]*/\s*>}
                      end
            if (m = updated_sheet_xml.match(pattern))
              end_pos = m.end(0)
              search_start = end_pos if end_pos > search_start
            end
          end

          tags_sub = "legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst"
          insert_pattern = if ws_prefix.empty?
                             %r{<(?:#{tags_sub})\b|</worksheet>}
                           else
                             %r{<#{Regexp.escape(ws_prefix)}(?:#{tags_sub})\b|</#{Regexp.escape(ws_prefix)}worksheet>}
                           end
          match = updated_sheet_xml.match(insert_pattern, search_start)
          if match
            updated_sheet_xml.insert(match.begin(0), legacy_tag)
          else
            updated_sheet_xml << legacy_tag
          end
        end

        root_tag = updated_sheet_xml[/\A\s*(?:<\?xml[^>]*\?>\s*)?<([a-zA-Z0-9_]+:)?worksheet\b[^>]*/, 0]
        updated_sheet_xml.sub!(/<([a-zA-Z0-9_]+:)?worksheet\b/, '\0 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"') if root_tag.nil? || !root_tag.match?(/\bxmlns:r\s*=/)

        ct_xml = replaced_entries["[Content_Types].xml"] || zip_reader.read_entry("[Content_Types].xml")
        if ct_xml && !ct_xml.include?("comments#{comment_file_num}.xml")
          override_tag = %(<Override PartName="/#{comments_path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>)
          ct_xml = ct_xml.sub(%r{</Types>}, "#{override_tag}</Types>")
          unless ct_xml.include?('Extension="vml"')
            default_vml = %(<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>)
            ct_xml = ct_xml.sub(%r{</Types>}, "#{default_vml}</Types>")
          end
          replaced_entries["[Content_Types].xml"] = ct_xml
        end
      end

      io = StringIO.new
      b = Ooxml::XmlBuilder.new(io)
      b.declaration
      b.open_tag("Relationships", { xmlns: "http://schemas.openxmlformats.org/package/2006/relationships" })
      existing_rels.each do |rel|
        rel_attrs = { Id: rel[:id], Type: rel[:type], Target: rel[:target] }
        rel_attrs[:TargetMode] = rel[:target_mode] if rel[:target_mode]
        b.empty_tag("Relationship", rel_attrs)
      end
      b.close_tag("Relationships")

      [updated_sheet_xml, io.string]
    end

    #: (Array[Elements::Row] rows, Array[untyped] shared_strings, Hash[untyped, Integer] sst_index, ?use_sst: bool, ?date1904: bool) -> String
    def serialize_sheet_data_xml(rows, shared_strings, sst_index, use_sst: true, date1904: false)
      io = StringIO.new
      io.write("<sheetData>")
      ws = Ooxml::WorksheetWriter.new(io, date1904: date1904)
      ws.instance_variable_set(:@started, true)

      rows.each do |row|
        attrs = {}
        attrs[:height] = row.height if row.height
        attrs[:hidden] = true if row.hidden
        attrs[:custom_height] = true if row.custom_height
        attrs[:outline_level] = row.outline_level if row.outline_level

        if use_sst
          row.cells.each do |cell|
            val = cell.value
            next unless val.is_a?(String) || val.is_a?(Elements::RichText)
            next if sst_index.key?(val)

            shared_strings << val
            sst_index[val] = shared_strings.size - 1
          end
        end

        ws.write_row(
          row.index,
          row.cells,
          attrs: attrs,
          unmapped: row.unmapped_data || [],
          sst_index: use_sst ? sst_index : nil
        )
      end

      buf = ws.instance_variable_get(:@row_buffer)
      if buf && !buf.empty?
        io.write(buf)
        buf.clear
      end
      io.write("</sheetData>")
      io.string
    end
  end

  # Builds a raw cell hash from a value for streaming writes.
  #
  # @param row_index [Integer] 0-based row index.
  # @param col_index [Integer] 0-based column index.
  # @param value [Object] Cell value.
  # @param sst [Array<String>] Shared strings array.
  # @param sst_index [Hash{String => Integer}] Shared strings index mapping.
  # @return [Hash{Symbol => Object}] Raw cell hash.
  # @api public
  #: (Integer row_index, Integer col_index, untyped value, Array[String] sst, Hash[String, Integer] sst_index, ?date1904: bool) -> Hash[Symbol, untyped]
  def self.build_raw_cell_from_value(row_index, col_index, value, sst, sst_index, date1904: false)
    # simplecov:disable
    # Edge case / untested delegation block
    ref = "#{Elements::Cell.column_letter(col_index)}#{row_index + 1}"
    result = { ref: ref }

    case value
    when Elements::Formula
      result[:formula] = value.expression
      result[:formula_type] = value.type if value.type
      result[:formula_ref] = value.ref if value.ref
      result[:formula_si] = value.shared_index if value.shared_index
      result[:formula_ca] = true if value.calculate_always
      unless value.cached_value.nil?
        cv = value.cached_value
        result[:value] = cv
        case cv
        when true, false
          result[:type] = "b"
        when Xlsxrb::Elements::CellError
          result[:type] = "e"
          result[:value] = cv.code
        end
      end
    when Hash
      if value.key?(:formula)
        result[:formula] = value[:formula]
        result[:formula_type] = value[:type] if value[:type]
        result[:formula_ref] = value[:ref] if value[:ref]
        si = value[:shared_index] || value[:si]
        result[:formula_si] = si if si
        result[:formula_ca] = true if value[:calculate_always]
        cv = value.key?(:value) ? value[:value] : value[:cached_value]
        unless cv.nil?
          result[:value] = cv
          case cv
          when true, false
            result[:type] = "b"
          when Xlsxrb::Elements::CellError
            result[:type] = "e"
            result[:value] = cv.code
          end
        end
      end
    when String, Xlsxrb::Elements::RichText
      idx = sst_index[value] ||= begin
        sst << value
        sst.size - 1
      end
      result[:value] = idx
      result[:type] = "s"
    when true, false
      result[:value] = value
      result[:type] = "b"
    when Integer, Float
      result[:value] = value
    when Xlsxrb::Elements::CellError
      result[:value] = value.code
      result[:type] = "e"
    when Date
      result[:value] = Xlsxrb::Ooxml::Utils.date_to_serial(value, date1904: date1904)
    when Time
      result[:value] = Xlsxrb::Ooxml::Utils.datetime_to_serial(value, date1904: date1904)
    # simplecov:enable
    when NilClass
      # empty cell
    end

    # simplecov:disable
    # Edge case / untested delegation block
    result
    # simplecov:enable
  end
end
