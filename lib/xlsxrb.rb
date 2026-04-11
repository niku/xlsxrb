# frozen_string_literal: true

require "date"
require "openssl"
require "securerandom"
require_relative "xlsxrb/version"
require_relative "xlsxrb/ooxml/zip_generator"
require_relative "xlsxrb/ooxml/writer"
require_relative "xlsxrb/ooxml/reader"
require_relative "xlsxrb/ooxml"
require_relative "xlsxrb/elements"

# Ruby XLSX read/write library.
module Xlsxrb
  class Error < StandardError; end

  # --- Facade API ---

  # Reads an XLSX file into an Elements::Workbook.
  # source: file path (String) or IO object.
  def self.read(source)
    entries = Ooxml::ZipReader.open(source, &:read_all)
    shared_strings = Ooxml::SharedStringsParser.parse(entries["xl/sharedStrings.xml"])
    styles = Ooxml::StylesParser.parse(entries["xl/styles.xml"])
    workbook_sheets = Ooxml::WorkbookParser.parse(entries["xl/workbook.xml"])
    rels = Ooxml::RelationshipsParser.parse(entries["xl/_rels/workbook.xml.rels"])

    sheets = workbook_sheets.map do |sheet_info|
      target = rels[sheet_info[:r_id]]
      next nil unless target

      sheet_path = target.start_with?("/") ? target.delete_prefix("/") : "xl/#{target}"
      sheet_xml = entries[sheet_path]
      build_worksheet(sheet_info[:name], sheet_xml, shared_strings, styles)
    end.compact

    Elements::Workbook.new(sheets: sheets, shared_strings: shared_strings, styles: styles)
  end

  # Writes an Elements::Workbook to an XLSX file.
  # target: file path (String) or IO object.
  def self.write(target, workbook)
    raise Error, "target is required" if target.nil?
    raise Error, "workbook must be an Elements::Workbook" unless workbook.is_a?(Elements::Workbook)

    sst = []
    sst_index = {}

    # Collect shared strings and build index
    sheet_data = workbook.sheets.map do |ws|
      rows = ws.rows.map do |row|
        cells = row.cells.map do |cell|
          raw = build_raw_cell(cell, sst, sst_index)
          raw
        end
        { index: row.index, cells: cells, attrs: build_row_attrs(row), unmapped: [] }
      end
      columns = ws.columns.map do |col|
        { index: col.index, width: col.width, hidden: col.hidden, custom_width: col.custom_width }
      end
      sd = { name: ws.name, rows: rows, columns: columns }
      sd[:charts] = ws.charts unless ws.charts.empty?
      sd
    end

    Ooxml::WorkbookWriter.write(target, sheets: sheet_data, shared_strings: sst, styles: workbook.styles)
  end

  # Streaming read: yields Elements::Row one at a time.
  # source: file path (String) or IO object.
  # Options:
  #   sheet: sheet index (0-based Integer) or name (String). Defaults to 0.
  def self.foreach(source, sheet: 0, &block)
    return enum_for(:foreach, source, sheet: sheet) unless block

    entries = Ooxml::ZipReader.open(source, &:read_all)
    shared_strings = Ooxml::SharedStringsParser.parse(entries["xl/sharedStrings.xml"])
    workbook_sheets = Ooxml::WorkbookParser.parse(entries["xl/workbook.xml"])
    rels = Ooxml::RelationshipsParser.parse(entries["xl/_rels/workbook.xml.rels"])

    target_sheet = case sheet
                   when Integer
                     workbook_sheets[sheet]
                   when String
                     workbook_sheets.find { |s| s[:name] == sheet }
                   end
    return unless target_sheet

    target = rels[target_sheet[:r_id]]
    return unless target

    sheet_path = target.start_with?("/") ? target.delete_prefix("/") : "xl/#{target}"
    sheet_xml = entries[sheet_path]
    return if sheet_xml.nil? || sheet_xml.empty?

    Ooxml::WorksheetParser.each_row(sheet_xml, shared_strings: shared_strings) do |raw_row|
      row = build_row_from_raw(raw_row)
      block.call(row)
    end
  end

  # Streaming write: yields a StreamWriter context for building XLSX on-the-fly.
  # target: file path (String) or IO object.
  def self.generate(target, &block)
    raise Error, "target is required" if target.nil?
    raise Error, "block is required" unless block

    stream_writer = StreamWriter.new(target)
    block.call(stream_writer)
    stream_writer.close
  end

  # Builds an Elements::Workbook in memory using a DSL.
  def self.build(&block)
    raise Error, "block is required" unless block

    builder = WorkbookBuilder.new
    block.call(builder)
    builder.build
  end

  # DSL context for Xlsxrb.build.
  class WorkbookBuilder
    def initialize
      @sheets = []
    end

    # Add a new sheet.
    def add_sheet(name = nil, &block)
      name ||= "Sheet#{@sheets.size + 1}"
      sheet_builder = WorksheetBuilder.new(name)
      block.call(sheet_builder) if block_given?
      @sheets << sheet_builder.build
    end

    def build
      Elements::Workbook.new(sheets: @sheets)
    end
  end

  # DSL context for a single worksheet in Xlsxrb.build.
  class WorksheetBuilder
    def initialize(name)
      @name = name
      @rows = []
      @columns = []
      @charts = []
    end

    # Add a row of values to the sheet.
    def add_row(values, height: nil, hidden: false, custom_height: false, outline_level: nil)
      row_index = @rows.size
      cells = values.each_with_index.map do |val, col_index|
        Elements::Cell.new(row_index: row_index, column_index: col_index, value: val)
      end

      @rows << Elements::Row.new(
        index: row_index,
        cells: cells,
        height: height,
        hidden: hidden,
        custom_height: custom_height || !height.nil?,
        outline_level: outline_level
      )
    end

    # Set column width for a 0-based column index.
    def set_column(index, width: nil, hidden: false, custom_width: false, outline_level: nil)
      @columns << Elements::Column.new(
        index: index,
        width: width,
        hidden: hidden,
        custom_width: custom_width || !width.nil?,
        outline_level: outline_level
      )
    end

    def add_chart(**options)
      @charts << options
    end

    def build
      Elements::Worksheet.new(name: @name, rows: @rows, columns: @columns, charts: @charts)
    end
  end

  # DSL context for Xlsxrb.generate streaming writes.
  class StreamWriter
    def initialize(target)
      @target = target
      @sst = []
      @sst_index = {}
      @sheets = []
      @current_sheet = nil
      @current_rows = []
      @current_columns = []
      @current_charts = []
    end

    # Start or switch to a named sheet.
    def add_sheet(name = nil)
      flush_current_sheet
      name ||= "Sheet#{@sheets.size + 1}"
      @current_sheet = name
      @current_rows = []
      @current_columns = []
      @current_charts = []

      return unless block_given?

      yield self
      flush_current_sheet
    end

    # Add a row of values. values is an Array.
    def add_row(values, height: nil, hidden: false)
      add_sheet if @current_sheet.nil?

      row_index = @current_rows.size
      cells = values.each_with_index.map do |val, col_idx|
        Xlsxrb.build_raw_cell_from_value(row_index, col_idx, val, @sst, @sst_index)
      end
      attrs = {}
      attrs[:height] = height if height
      attrs[:hidden] = true if hidden
      @current_rows << { index: row_index, cells: cells, attrs: attrs, unmapped: [] }
    end

    # Set column width for a 0-based column index.
    def set_column(index, width: nil, hidden: false)
      add_sheet if @current_sheet.nil?

      @current_columns << { index: index, width: width, hidden: hidden, custom_width: !width.nil? }
    end

    # Add a chart to the current sheet.
    def add_chart(**options)
      add_sheet if @current_sheet.nil?

      @current_charts << options
    end

    def close
      flush_current_sheet
      Ooxml::WorkbookWriter.write(@target, sheets: @sheets, shared_strings: @sst)
    end

    private

    def flush_current_sheet
      return unless @current_sheet

      sheet_data = { name: @current_sheet, rows: @current_rows, columns: @current_columns }
      sheet_data[:charts] = @current_charts unless @current_charts.empty?
      @sheets << sheet_data
      @current_sheet = nil
    end
  end

  class << self
    private

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

    def build_row_from_raw(raw_row)
      cells = raw_row[:cells].map do |rc|
        parsed = Elements::Cell.parse_ref(rc[:ref]) if rc[:ref]
        row_idx = parsed ? parsed[0] : raw_row[:index]
        col_idx = parsed ? parsed[1] : 0

        Elements::Cell.new(
          row_index: row_idx,
          column_index: col_idx,
          value: rc[:value],
          formula: rc[:formula],
          style_index: rc[:style_index]
        )
      end
      attrs = raw_row[:attrs] || {}
      Elements::Row.new(
        index: raw_row[:index],
        cells: cells,
        height: attrs[:height],
        hidden: attrs[:hidden] || false,
        custom_height: attrs[:custom_height] || false,
        outline_level: attrs[:outline_level]
      )
    end

    def build_raw_cell(cell, sst, sst_index)
      ref = cell.ref
      value = cell.value
      result = { ref: ref, style_index: cell.style_index }

      case value
      when String
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
      when Date
        result[:value] = Xlsxrb::Ooxml::Utils.date_to_serial(value)
      when Time
        result[:value] = Xlsxrb::Ooxml::Utils.datetime_to_serial(value)
      when NilClass
        # empty cell
      end

      result[:formula] = cell.formula if cell.formula
      result
    end

    def build_row_attrs(row)
      attrs = {}
      attrs[:height] = row.height if row.height
      attrs[:hidden] = true if row.hidden
      attrs[:custom_height] = true if row.custom_height
      attrs
    end
  end

  # Builds a raw cell hash from a value for streaming writes.
  def self.build_raw_cell_from_value(row_index, col_index, value, sst, sst_index)
    ref = "#{Elements::Cell.column_letter(col_index)}#{row_index + 1}"
    result = { ref: ref }

    case value
    when String
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
    when Date
      result[:value] = Xlsxrb::Ooxml::Utils.date_to_serial(value)
    when Time
      result[:value] = Xlsxrb::Ooxml::Utils.datetime_to_serial(value)
    when NilClass
      # empty cell
    end

    result
  end
end
