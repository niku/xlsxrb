# frozen_string_literal: true

require_relative "zip_generator"

module Xlsxrb
  # Writes spreadsheet data into an XLSX file.
  class Writer
    XML_HEADER = %(<?xml version="1.0" encoding="UTF-8"?>)
    CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
    REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
    SSML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    CELL_ADDRESS_PATTERN = /\A([A-Z]{1,3})(\d+)\z/
    MAX_ROW = 1_048_576
    MAX_COLUMN_INDEX = 16_384 # XFD

    def initialize
      @sheets = { "Sheet1" => {} }
      @column_widths = { "Sheet1" => {} }
      @row_attrs = { "Sheet1" => {} }
      @merge_cells = { "Sheet1" => [] }
      @hyperlinks = { "Sheet1" => {} }
      @sheet_order = ["Sheet1"]
    end

    # Adds a new sheet. Raises if name is already taken.
    def add_sheet(name)
      raise ArgumentError, "sheet already exists: #{name}" if @sheets.key?(name)

      @sheets[name] = {}
      @column_widths[name] = {}
      @row_attrs[name] = {}
      @merge_cells[name] = []
      @hyperlinks[name] = {}
      @sheet_order << name
    end

    # Registers a cell value at the given address (e.g. "A1").
    def set_cell(cell_address, value, sheet: nil)
      validate_cell_address!(cell_address)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @sheets.key?(sheet_name)

      @sheets[sheet_name][cell_address] = value
    end

    # Returns the registered cells for the first (or given) sheet.
    def cells(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @sheets[sheet_name] || {}
    end

    # Sets the width for a column (e.g. "A", "BC").
    def set_column_width(col_letter, width, sheet: nil)
      raise ArgumentError, "column must be a String of uppercase letters" unless col_letter.is_a?(String) && col_letter.match?(/\A[A-Z]+\z/)

      col_index = column_letter_to_index(col_letter)
      raise ArgumentError, "column out of range: #{col_letter}" unless col_index.between?(1, MAX_COLUMN_INDEX)

      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @column_widths.key?(sheet_name)

      @column_widths[sheet_name][col_letter] = width
    end

    # Returns column widths for the first (or given) sheet.
    def column_widths(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @column_widths[sheet_name] || {}
    end

    # Sets a row height.
    def set_row_height(row_num, height, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @row_attrs.key?(sheet_name)
      raise ArgumentError, "row must be a positive Integer" unless row_num.is_a?(Integer) && row_num >= 1

      @row_attrs[sheet_name][row_num] ||= {}
      @row_attrs[sheet_name][row_num][:height] = height
    end

    # Hides a row.
    def set_row_hidden(row_num, hidden: true, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @row_attrs.key?(sheet_name)
      raise ArgumentError, "row must be a positive Integer" unless row_num.is_a?(Integer) && row_num >= 1

      @row_attrs[sheet_name][row_num] ||= {}
      @row_attrs[sheet_name][row_num][:hidden] = hidden
    end

    # Returns row attributes for the first (or given) sheet.
    def row_attributes(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @row_attrs[sheet_name] || {}
    end

    # Merges a range of cells (e.g. "A1:B2").
    def merge_cells(range, sheet: nil)
      raise ArgumentError, "range must be a String like 'A1:B2'" unless range.is_a?(String) && range.match?(/\A[A-Z]+\d+:[A-Z]+\d+\z/)

      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @merge_cells.key?(sheet_name)

      @merge_cells[sheet_name] << range
    end

    # Returns merged cell ranges for the first (or given) sheet.
    def merged_cells(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @merge_cells[sheet_name] || []
    end

    # Adds a hyperlink on a cell.
    def add_hyperlink(cell_address, url, sheet: nil)
      validate_cell_address!(cell_address)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @hyperlinks.key?(sheet_name)

      @hyperlinks[sheet_name][cell_address] = url
    end

    # Returns hyperlinks for the first (or given) sheet.
    def hyperlinks(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @hyperlinks[sheet_name] || {}
    end

    # Returns ordered sheet names.
    attr_reader :sheet_order

    # Writes the workbook as an XLSX file to the given path.
    def write(filepath)
      entries = {
        "[Content_Types].xml" => generate_content_types_xml,
        "_rels/.rels" => generate_rels_root,
        "xl/workbook.xml" => generate_workbook_xml,
        "xl/_rels/workbook.xml.rels" => generate_workbook_rels
      }

      @sheet_order.each_with_index do |sheet_name, i|
        entries["xl/worksheets/sheet#{i + 1}.xml"] = generate_worksheet_xml(
          @sheets[sheet_name], @column_widths[sheet_name], @row_attrs[sheet_name],
          @merge_cells[sheet_name], @hyperlinks[sheet_name]
        )
        next if @hyperlinks[sheet_name].empty?

        entries["xl/worksheets/_rels/sheet#{i + 1}.xml.rels"] = generate_worksheet_rels(@hyperlinks[sheet_name])
      end

      generator = ZipGenerator.new(filepath)
      entries.each { |path, content| generator.add_entry(path, content) }
      generator.generate
    end

    private

    def generate_content_types_xml
      parts = [
        XML_HEADER,
        %(<Types xmlns="#{CT_NS}">),
        %(<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>),
        %(<Default Extension="xml" ContentType="application/xml"/>),
        %(<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>)
      ]
      @sheet_order.each_with_index do |_, i|
        parts << %(<Override PartName="/xl/worksheets/sheet#{i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>)
      end
      parts << "</Types>"
      parts.join
    end

    def generate_rels_root
      parts = [
        XML_HEADER,
        %(<Relationships xmlns="#{REL_NS}">),
        %(<Relationship Id="rId1" Type="#{DOC_REL_NS}/officeDocument" Target="xl/workbook.xml"/>),
        "</Relationships>"
      ]
      parts.join
    end

    def generate_workbook_xml
      parts = [
        XML_HEADER,
        %(<workbook xmlns="#{SSML_NS}" xmlns:r="#{DOC_REL_NS}">),
        "<sheets>"
      ]
      @sheet_order.each_with_index do |name, i|
        parts << %(<sheet name="#{xml_escape(name)}" sheetId="#{i + 1}" r:id="rId#{i + 1}"/>)
      end
      parts << "</sheets>"
      parts << "</workbook>"
      parts.join
    end

    def generate_workbook_rels
      parts = [
        XML_HEADER,
        %(<Relationships xmlns="#{REL_NS}">)
      ]
      @sheet_order.each_with_index do |_, i|
        parts << %(<Relationship Id="rId#{i + 1}" Type="#{DOC_REL_NS}/worksheet" Target="worksheets/sheet#{i + 1}.xml"/>)
      end
      parts << "</Relationships>"
      parts.join
    end

    def generate_worksheet_xml(sheet_cells, sheet_col_widths, sheet_row_attrs, sheet_merge_cells, sheet_hyperlinks)
      worksheet_attrs = %(xmlns="#{SSML_NS}")
      worksheet_attrs << %( xmlns:r="#{DOC_REL_NS}") unless sheet_hyperlinks.empty?
      parts = [
        XML_HEADER,
        "<worksheet #{worksheet_attrs}>"
      ]

      # Emit <cols> if column widths are defined.
      unless sheet_col_widths.empty?
        parts << "<cols>"
        sheet_col_widths.sort_by { |col, _| column_letter_to_index(col) }.each do |col_letter, width|
          idx = column_letter_to_index(col_letter)
          parts << %(<col min="#{idx}" max="#{idx}" width="#{width}" customWidth="1"/>)
        end
        parts << "</cols>"
      end

      parts << "<sheetData>"

      # Group cells by row number.
      cells_by_row = {}
      sheet_cells.each do |address, value|
        row_num = extract_row_number(address)
        col_letter = extract_column_letter(address)
        cells_by_row[row_num] ||= {}
        cells_by_row[row_num][col_letter] = value
      end

      # Include rows that have attributes but no cells.
      sheet_row_attrs.each_key { |rn| cells_by_row[rn] ||= {} }

      # Emit rows in ascending order.
      cells_by_row.sort.each do |row_num, row_cells|
        attrs = %(r="#{row_num}")
        ra = sheet_row_attrs[row_num]
        if ra
          attrs << %( ht="#{ra[:height]}" customHeight="1") if ra[:height]
          attrs << %( hidden="1") if ra[:hidden]
        end
        parts << "<row #{attrs}>"
        row_cells.sort_by { |col, _| column_letter_to_index(col) }.each do |col_letter, value|
          cell_ref = "#{col_letter}#{row_num}"
          parts << cell_xml(cell_ref, value)
        end
        parts << "</row>"
      end

      parts << "</sheetData>"

      # Emit <mergeCells> if merge ranges are defined.
      unless sheet_merge_cells.empty?
        parts << %(<mergeCells count="#{sheet_merge_cells.size}">)
        sheet_merge_cells.each { |ref| parts << %(<mergeCell ref="#{ref}"/>) }
        parts << "</mergeCells>"
      end

      # Emit <hyperlinks> if hyperlinks are defined.
      unless sheet_hyperlinks.empty?
        parts << "<hyperlinks>"
        sheet_hyperlinks.each_with_index do |(cell_ref, _url), idx|
          parts << %(<hyperlink ref="#{cell_ref}" r:id="rId#{idx + 1}"/>)
        end
        parts << "</hyperlinks>"
      end

      parts << "</worksheet>"
      parts.join
    end

    def generate_worksheet_rels(sheet_hyperlinks)
      parts = [
        XML_HEADER,
        %(<Relationships xmlns="#{REL_NS}">)
      ]
      sheet_hyperlinks.each_with_index do |(_cell_ref, url), idx|
        parts << %(<Relationship Id="rId#{idx + 1}" Type="#{DOC_REL_NS}/hyperlink" Target="#{xml_escape(url)}" TargetMode="External"/>)
      end
      parts << "</Relationships>"
      parts.join
    end

    def xml_escape(value)
      value.to_s
           .gsub("&", "&amp;")
           .gsub("<", "&lt;")
           .gsub(">", "&gt;")
           .gsub('"', "&quot;")
           .gsub("'", "&apos;")
    end

    def cell_xml(cell_ref, value)
      case value
      when Formula
        parts = %(<c r="#{cell_ref}"><f>#{xml_escape(value.expression)}</f>)
        parts << "<v>#{xml_escape(value.cached_value.to_s)}</v>" unless value.cached_value.nil?
        parts << "</c>"
        parts
      when true, false
        %(<c r="#{cell_ref}" t="b"><v>#{value ? 1 : 0}</v></c>)
      when Numeric
        %(<c r="#{cell_ref}"><v>#{value}</v></c>)
      else
        %(<c r="#{cell_ref}" t="inlineStr"><is><t>#{xml_escape(value)}</t></is></c>)
      end
    end

    def validate_cell_address!(cell_address)
      raise ArgumentError, "cell address must be a String" unless cell_address.is_a?(String)

      match = cell_address.match(CELL_ADDRESS_PATTERN)
      raise ArgumentError, "invalid cell address: #{cell_address.inspect}" unless match

      row_num = match[2].to_i
      raise ArgumentError, "row out of range: #{row_num}" unless row_num.between?(1, MAX_ROW)

      col_index = column_letter_to_index(match[1])
      raise ArgumentError, "column out of range: #{match[1]}" unless col_index.between?(1, MAX_COLUMN_INDEX)
    end

    def column_letter_to_index(letters)
      letters.chars.reduce(0) { |sum, char| (sum * 26) + (char.ord - "A".ord + 1) }
    end

    # Extracts the column letter(s) from a cell address, e.g. "A" from "A1".
    def extract_column_letter(cell_address)
      cell_address.match(/^([A-Z]+)/)[1]
    end

    # Extracts the row number from a cell address, e.g. 1 from "A1".
    def extract_row_number(cell_address)
      cell_address.match(/(\d+)$/)[1].to_i
    end
  end
end
