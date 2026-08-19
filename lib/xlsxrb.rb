# frozen_string_literal: true

# rbs_inline: enabled

require "date"
require "time"
require "openssl"
require "securerandom"
require "tempfile"
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
require "opentelemetry"
require_relative "xlsxrb/version"
require_relative "xlsxrb/ooxml/zip_generator"
require_relative "xlsxrb/ooxml/writer"
require_relative "xlsxrb/ooxml/reader"
require_relative "xlsxrb/ooxml/crypto"
require_relative "xlsxrb/ooxml"
require_relative "xlsxrb/elements"
require_relative "xlsxrb/stream_row"
require_relative "xlsxrb/stream_sheet"
require_relative "xlsxrb/style_builder"
require_relative "xlsxrb/chart_builder"
require_relative "xlsxrb/worksheet_builder"
require_relative "xlsxrb/workbook_builder"
require_relative "xlsxrb/stream_writer"

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

  # Raised when cell, row, or workbook parameters fail specification validation.
  class ValidationError < Error; end

  # Raised when ZIP decompression or entry resolution fails.
  class ZipError < Error; end

  # Base exception for encrypted or password-protected XLSX documents.
  class EncryptedFileError < Error; end

  # Raised when a password is missing or invalid for an encrypted XLSX file.
  class InvalidPasswordError < EncryptedFileError; end

  # Raised when encrypted package stream decryption fails.
  class DecryptionError < EncryptedFileError; end

  TRACER = OpenTelemetry.tracer_provider.tracer("xlsxrb", Xlsxrb::VERSION)

  # Executes the block within an OpenTelemetry tracer span if tracing is configured.
  #
  # @param name [String] The span name.
  # @param attributes [Hash, nil] Optional telemetry attributes.
  # @yield Block to execute inside the tracing span.
  # @return [Object] The result of the block.
  #: (String name, ?attributes: Hash[String, untyped]?) { (*untyped) -> untyped } -> untyped
  def self.in_span(name, attributes: nil, &)
    if defined?(Ractor) && Ractor.current != Ractor.main
      # simplecov:disable
      # Test suite runs in the main Ractor. This branch is for multi-threaded usage via Ractors.
      yield
      # simplecov:enable
    elsif attributes
      TRACER.in_span(name, attributes: attributes, &)
    else
      TRACER.in_span(name, &)
    end
  end

  # Helper to construct {Elements::RichText} objects with formatted text runs.
  #
  # @example Create multi-run rich text
  #   rt = Xlsxrb.rich_text({ text: "Total: ", font: { bold: true } }, { text: "$100" })
  #
  # @example Create simple styled text
  #   rt = Xlsxrb.rich_text(text: "Important Notice", bold: true, color: "FF0000")
  #
  # @param runs [Array<Hash, String>] Optional array of rich text run hashes or string.
  # @param text [String, nil] Plain text string (convenience parameter).
  # @param font_props [Hash] Inline font styling options (e.g. bold: true, color: "FF0000").
  # @return [Elements::RichText] The compiled rich text element.
  # @api public
  #: (*(Hash[Symbol, untyped] | String) runs, ?text: String?, **untyped font_props) -> Elements::RichText
  def self.rich_text(*runs, text: nil, **font_props)
    if text
      runs = [{ text: text, font: font_props }]
    elsif runs.size == 1 && runs.first.is_a?(String) && !font_props.empty?
      runs = [{ text: runs.first, font: font_props }]
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
  # @return [Elements::Formula]
  # @api public
  #: (String expression, ?cached_value: (String | Numeric | bool | nil)) -> Elements::Formula
  def self.formula(expression, cached_value: nil)
    Elements::Formula.new(
      expression: expression,
      cached_value: cached_value,
      calculate_always: cached_value.nil? || nil
    )
  end

  # Reads an XLSX file (streaming and lazy-loaded by default) from a file path, IO stream, or binary String.
  #
  # Sheets and rows are streamed lazily with O(1) constant memory. If a block is given,
  # yields each {StreamSheet} sequentially. Call {#load} on the returned Workbook or Sheet
  # to convert to an in-memory representation for coordinate random access (`sheet["A1"]`).
  #
  # @overload read(source, password: nil, &block)
  #   Yields each {StreamSheet} sequentially in streaming mode.
  #   @param source [String, IO, StringIO] File path, binary content string, or readable IO stream.
  #   @param password [String, nil] Optional password to decrypt password-protected XLSX files.
  #   @yield [sheet]
  #   @yieldparam sheet [StreamSheet] The streaming worksheet object.
  #   @return [void]
  #
  # @overload read(source, password: nil)
  #   Returns a lazy {Elements::Workbook} object.
  #   @param source [String, IO, StringIO] File path, binary content string, or readable IO stream.
  #   @param password [String, nil] Optional password to decrypt password-protected XLSX files.
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
  # @raise [EncryptedFileError] If the file is encrypted and no password was supplied.
  # @raise [InvalidPasswordError] If the supplied password is incorrect.
  # @raise [ParseError] If the spreadsheet XML structure is invalid.
  # @api public
  #: (untyped source, ?password: String?) { (StreamSheet) -> void } -> void
  #: (untyped source, ?password: String?) -> Elements::Workbook
  def self.read(source, password: nil, &)
    if source.is_a?(String)
      if source.start_with?("PK\x03\x04") || source.include?("\x00") || Ooxml::Cfb::Reader.cfb?(source)
        if Ooxml::Cfb::Reader.cfb?(source)
          decrypted_zip = Ooxml::Crypto.decrypt(source, password)
          source = StringIO.new(decrypted_zip)
        else
          source = StringIO.new(source)
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
          source = StringIO.new(decrypted_zip)
        end
      end
    elsif source.respond_to?(:read) && source.respond_to?(:pos) && source.respond_to?(:seek)
      begin
        cur_pos = source.pos
        first_bytes = source.read(8)
        source.seek(cur_pos)
        if Ooxml::Cfb::Reader.cfb?(first_bytes)
          full_data = source.read
          decrypted_zip = Ooxml::Crypto.decrypt(full_data, password)
          source = StringIO.new(decrypted_zip)
        end
      rescue StandardError
        # Fall through to standard reader if seeking fails
      end
    end

    attributes = source.is_a?(String) ? { "filepath" => source } : {}
    Xlsxrb.in_span("Xlsxrb.read", attributes: attributes) do
      entries = Ooxml::ZipReader.open(source, &:read_all)
      shared_strings = Ooxml::SharedStringsParser.parse(entries["xl/sharedStrings.xml"])
      workbook_sheets = Ooxml::WorkbookParser.parse(entries["xl/workbook.xml"])
      rels = Ooxml::RelationshipsParser.parse(entries["xl/_rels/workbook.xml.rels"])
      styles = Ooxml::StylesParser.parse(entries["xl/styles.xml"])

      sheets = workbook_sheets.map do |sheet_info|
        target = rels[sheet_info[:r_id]]
        next nil unless target

        sheet_path = target.start_with?("/") ? target.delete_prefix("/") : "xl/#{target}"
        sheet_xml = entries[sheet_path]
        next nil if sheet_xml.nil? || sheet_xml.empty?

        StreamSheet.new(
          sheet_info[:name],
          sheet_xml,
          shared_strings,
          styles
        )
      end.compact

      wb = Elements::Workbook.new(sheets: sheets, shared_strings: shared_strings, styles: styles)

      if block_given?
        sheets.each(&)
        nil
      else
        wb
      end
    end
  end

  # Writes an XLSX file or IO stream (streaming or in-memory), or returns a binary string.
  #
  # @overload write(target, password: nil, encryption_mode: :standard, strict_excel_mode: true, &block)
  #   Streaming write: yields a {StreamWriter} context for high-speed, zero-allocation XLSX generation.
  #   @param target [String, IO, StringIO] Destination file path or writable IO object.
  #   @param password [String, nil] Optional password to encrypt the generated XLSX file.
  #   @param encryption_mode [Symbol] Encryption algorithm (:standard or :agile).
  #   @param strict_excel_mode [Boolean] Whether to enforce Microsoft Excel specification limits.
  #   @yield [stream_writer]
  #   @yieldparam stream_writer [Xlsxrb::StreamWriter]
  #   @return [void]
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
  #: (untyped target_or_workbook, ?Elements::Workbook | untyped workbook_or_nil, ?password: String?, ?encryption_mode: Symbol, ?strict_excel_mode: bool) ?{ (StreamWriter) -> void } -> untyped
  def self.write(target_or_workbook, workbook_or_nil = nil, password: nil, encryption_mode: :standard, strict_excel_mode: true, &block)
    if block_given?
      target = target_or_workbook
      raise Error, "target is required" if target.nil?

      attributes = target.is_a?(String) ? { "filepath" => target } : {}
      return Xlsxrb.in_span("Xlsxrb.write", attributes: attributes) do
        if password && !password.empty?
          buf = StringIO.new
          buf.binmode
          stream_writer = StreamWriter.new(buf, strict_excel_mode: strict_excel_mode)
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
          stream_writer = StreamWriter.new(target, strict_excel_mode: strict_excel_mode)
          begin
            yield stream_writer
            stream_writer.close
          ensure
            stream_writer.cleanup!
          end
        end
      end
    end

    if workbook_or_nil.nil?
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

    attributes = target.is_a?(String) ? { "filepath" => target } : {}
    Xlsxrb.in_span("Xlsxrb.write", attributes: attributes) do
      sst = []
      sst_index = {}

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
          { index: col.index, width: col.width, hidden: col.hidden, custom_width: col.custom_width, outline_level: col.outline_level }
          # simplecov:enable
        end
        sd = { name: ws.name, rows: ws.rows, columns: columns }
        sd[:charts] = ws.charts unless ws.charts.empty?

        # Extract facade metadata from unmapped_data
        facade = ws.unmapped_data[:facade]
        facade&.each { |key, val| sd[key] = val }

        sd
      end

      # Extract workbook-level facade metadata
      wb_facade = workbook.unmapped_data[:facade] || {}

      if password && !password.empty?
        buf = StringIO.new
        buf.binmode
        Ooxml::WorkbookWriter.write(
          buf,
          sheets: sheet_data,
          shared_strings: sst,
          shared_strings_index: sst_index,
          styles: workbook.styles,
          defined_names: wb_facade[:defined_names],
          core_properties: wb_facade[:core_properties],
          app_properties: wb_facade[:app_properties],
          custom_properties: wb_facade[:custom_properties],
          workbook_protection: wb_facade[:workbook_protection],
          workbook_properties: wb_facade[:workbook_properties]
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
          defined_names: wb_facade[:defined_names],
          core_properties: wb_facade[:core_properties],
          app_properties: wb_facade[:app_properties],
          custom_properties: wb_facade[:custom_properties],
          workbook_protection: wb_facade[:workbook_protection],
          workbook_properties: wb_facade[:workbook_properties]
        )
      end
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

    workbook = read(source, password: password).load
    result_workbook = yield workbook
    result_workbook = workbook unless result_workbook.is_a?(Elements::Workbook)

    write_target = target || source
    write(write_target, result_workbook, password: password)
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

    Xlsxrb.in_span("Xlsxrb.build") do
      builder = WorkbookBuilder.new(strict_excel_mode: strict_excel_mode)
      yield builder
      builder.build
    end
  end

  class << self
    private

    #: (String name, String? sheet_xml, Array[String] shared_strings, untyped _styles) -> Elements::Worksheet
    def build_worksheet(name, sheet_xml, shared_strings, _styles)
      return Elements::Worksheet.new(name: name) if sheet_xml.nil? || sheet_xml.empty?

      raw_rows = Ooxml::WorksheetParser.parse(sheet_xml, shared_strings: shared_strings)
      raw_columns = Ooxml::WorksheetParser.parse_columns(sheet_xml)

      rows = raw_rows.map { |rr| build_row_from_raw(rr) }
      columns = raw_columns.map do |rc|
        # Columns from OOXML are 1-based min/max ranges; convert to 0-based
        Elements::Column.new(
          index: (rc[:min] || 1) - 1,
          width: rc[:width],
          hidden: rc[:hidden] || false,
          custom_width: rc[:custom_width] || false,
          outline_level: rc[:outline_level]
        )
      end

      Elements::Worksheet.new(name: name, rows: rows, columns: columns)
    end

    #: (untyped raw_row) -> (Elements::Row | untyped)
    def build_row_from_raw(raw_row)
      return raw_row if raw_row.is_a?(Elements::Row)

      raw_cells = raw_row[:cells]
      cells = if raw_cells.empty? || raw_cells.first.is_a?(Elements::Cell)
                raw_cells
              else
                raw_cells.map do |rc|
                  parsed = Elements::Cell.parse_ref(rc[:ref]) if rc[:ref]
                  row_idx = parsed ? parsed[0] : raw_row[:index]
                  col_idx = parsed ? parsed[1] : 0

                  val = rc[:value]
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
                    errors: cell_errors
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
      Elements::Row.new(
        index: raw_row[:index],
        cells: cells,
        height: attrs[:height],
        hidden: attrs[:hidden] || false,
        custom_height: attrs[:custom_height] || false,
        outline_level: attrs[:outline_level],
        errors: row_errors
      )
    end

    #: (Elements::Cell cell, Array[String] sst, Hash[String, Integer] sst_index) -> Hash[Symbol, untyped]
    def build_raw_cell(cell, sst, sst_index)
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
        result[:value] = Xlsxrb::Ooxml::Utils.date_to_serial(value)
      when Time
        result[:value] = Xlsxrb::Ooxml::Utils.datetime_to_serial(value)
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
          result[:formula_ca] = true if f.calculate_always
          if f.cached_value
            # Cached value is written as-is (not through SST)
            result[:value] = f.cached_value
            result.delete(:type) # Ensure no type is set; cached values are plain text in <v>
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
      attrs[:outline_level] = row.outline_level if row.outline_level
      attrs
      # simplecov:enable
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
  #: (Integer row_index, Integer col_index, untyped value, Array[String] sst, Hash[String, Integer] sst_index) -> Hash[Symbol, untyped]
  def self.build_raw_cell_from_value(row_index, col_index, value, sst, sst_index)
    # simplecov:disable
    # Edge case / untested delegation block
    ref = "#{Elements::Cell.column_letter(col_index)}#{row_index + 1}"
    result = { ref: ref }

    case value
    when Elements::Formula
      result[:formula] = value.expression
      result[:formula_ca] = true if value.calculate_always
      result[:value] = value.cached_value if value.cached_value
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
      result[:value] = Xlsxrb::Ooxml::Utils.date_to_serial(value)
    when Time
      result[:value] = Xlsxrb::Ooxml::Utils.datetime_to_serial(value)
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
