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
require_relative "xlsxrb/style_builder"

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
      @sheet_builders = [] # Keep track of sheet builders for style processing
    end

    # Add a new sheet.
    def add_sheet(name = nil, &block)
      name ||= "Sheet#{@sheets.size + 1}"
      sheet_builder = WorksheetBuilder.new(name)
      block.call(sheet_builder) if block_given?
      @sheet_builders << sheet_builder
      @sheets << sheet_builder.build
    end

    def build
      # Process styles from all sheets and collect style definitions
      processed_sheets, styles_definition = process_styles(@sheets)
      Elements::Workbook.new(sheets: processed_sheets, styles: styles_definition)
    end

    private

    def process_styles(sheets)
      # Collect all unique StyleBuilders from all sheets
      all_style_builders = {}
      @sheet_builders.each do |sheet_builder|
        sheet_builder.styles.each do |style_name, style_builder|
          all_style_builders[style_name] = style_builder
        end
      end

      return [sheets, {}] if all_style_builders.empty?

      # Create a temporary writer to register styles and get numeric IDs
      temp_writer = Ooxml::Writer.new
      style_name_to_id = {}
      all_style_builders.each do |name, builder|
        style_id = builder.register_with(temp_writer)
        style_name_to_id[name] = style_id
      end

      # Capture the style definitions from the temporary writer
      styles_definition = extract_styles_from_writer(temp_writer)

      # Update all cells with their resolved style IDs
      updated_sheets = sheets.map do |sheet|
        new_rows = sheet.rows.map do |row|
          new_cells = row.cells.map do |cell|
            # If style_index is a string (style name), resolve it to a numeric ID
            if cell.style_index.is_a?(String) && style_name_to_id.key?(cell.style_index)
              Elements::Cell.new(
                row_index: cell.row_index,
                column_index: cell.column_index,
                value: cell.value,
                formula: cell.formula,
                style_index: style_name_to_id[cell.style_index],
                unmapped_data: cell.unmapped_data,
                errors: cell.errors
              )
            else
              cell
            end
          end
          Elements::Row.new(
            index: row.index,
            cells: new_cells,
            height: row.height,
            hidden: row.hidden,
            custom_height: row.custom_height,
            outline_level: row.outline_level,
            unmapped_data: row.unmapped_data,
            errors: row.errors
          )
        end
        Elements::Worksheet.new(
          name: sheet.name,
          rows: new_rows,
          columns: sheet.columns,
          charts: sheet.charts,
          unmapped_data: sheet.unmapped_data,
          errors: sheet.errors
        )
      end

      [updated_sheets, styles_definition]
    end

    def extract_styles_from_writer(writer)
      # Extract style definitions from the writer that can be reused
      # This captures the fonts, fills, borders, and xf entries that were created
      {
        fonts: writer.instance_variable_get(:@fonts).dup,
        fills: writer.instance_variable_get(:@fills).dup,
        borders: writer.instance_variable_get(:@borders).dup,
        xf_entries: writer.instance_variable_get(:@xf_entries).dup,
        num_fmts: writer.instance_variable_get(:@num_fmts).dup
      }
    end
  end

  # DSL context for a single worksheet in Xlsxrb.build.
  class WorksheetBuilder
    def initialize(name)
      @name = name
      @rows = []
      @columns = []
      @charts = []
      @styles = {} # { style_name => StyleBuilder }
      @style_index_map = {} # { style_name => xf_index } (populated at build time)
    end

    # Define a named style that can be applied to cells.
    def add_style(name, &block)
      style_builder = StyleBuilder.new(name)
      block.call(style_builder) if block_given?
      @styles[name] = style_builder
      style_builder
    end

    # Add a row of values to the sheet.
    # values:: Array of cell values
    # styles:: Hash mapping column indices to style names, or Array of style names for each column
    def add_row(values, styles: nil, height: nil, hidden: false, custom_height: false, outline_level: nil)
      row_index = @rows.size
      cells = values.each_with_index.map do |val, col_index|
        style_name = case styles
                     when Hash
                       styles[col_index]
                     when Array
                       styles[col_index]
                     end
        # Style index will be resolved at build time
        Elements::Cell.new(
          row_index: row_index,
          column_index: col_index,
          value: val,
          style_index: style_name # Will store the style name for now
        )
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

    # Internal: returns styles for later processing by WorkbookBuilder
    attr_reader :styles
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
      @styles = {} # { style_name => StyleBuilder }
      @cell_styles = {} # { "SheetName!A1" => style_name }
    end

    # Define a named style that can be applied to cells.
    def add_style(name, &block)
      style_builder = StyleBuilder.new(name)
      block.call(style_builder) if block_given?
      @styles[name] = style_builder
      style_builder
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
    # styles:: Hash mapping column indices to style names, or Array of style names for each column
    def add_row(values, styles: nil, height: nil, hidden: false)
      add_sheet if @current_sheet.nil?

      row_index = @current_rows.size
      cells = values.each_with_index.map do |val, col_idx|
        cell = Xlsxrb.build_raw_cell_from_value(row_index, col_idx, val, @sst, @sst_index)

        # Track style assignment if provided
        style_name = case styles
                     when Hash
                       styles[col_idx]
                     when Array
                       styles[col_idx]
                     end

        if style_name && @styles.key?(style_name)
          cell_ref = cell[:ref]
          @cell_styles["#{@current_sheet}!#{cell_ref}"] = style_name
        end

        cell
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

      # Process styles before writing
      result = apply_styles(@sheets)
      sheet_data_with_styles = result[:sheets]
      styles_definition = result[:styles]

      Ooxml::WorkbookWriter.write(@target, sheets: sheet_data_with_styles || @sheets, shared_strings: @sst, styles: styles_definition)
    end

    private

    def flush_current_sheet
      return unless @current_sheet

      sheet_data = { name: @current_sheet, rows: @current_rows, columns: @current_columns }
      sheet_data[:charts] = @current_charts unless @current_charts.empty?
      @sheets << sheet_data
      @current_sheet = nil
    end

    def apply_styles(sheets)
      return { sheets: sheets, styles: {} } if @styles.empty?

      # Create a temporary writer to register all styles and get numeric IDs
      temp_writer = Ooxml::Writer.new
      style_name_to_id = {}

      @styles.each do |name, builder|
        style_id = builder.register_with(temp_writer)
        style_name_to_id[name] = style_id
      end

      # Extract style definitions from the temporary writer
      styles_definition = {
        fonts: temp_writer.instance_variable_get(:@fonts).dup,
        fills: temp_writer.instance_variable_get(:@fills).dup,
        borders: temp_writer.instance_variable_get(:@borders).dup,
        xf_entries: temp_writer.instance_variable_get(:@xf_entries).dup,
        num_fmts: temp_writer.instance_variable_get(:@num_fmts).dup
      }

      # Apply style IDs to cells based on tracked assignments
      updated_sheets = sheets.map do |sheet|
        new_rows = sheet[:rows].map do |row|
          new_cells = row[:cells].map do |cell|
            cell_key = "#{sheet[:name]}!#{cell[:ref]}"
            if @cell_styles.key?(cell_key)
              style_name = @cell_styles[cell_key]
              cell.merge(style_index: style_name_to_id[style_name])
            else
              cell
            end
          end
          row.merge(cells: new_cells)
        end
        sheet.merge(rows: new_rows)
      end

      { sheets: updated_sheets, styles: styles_definition }
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
