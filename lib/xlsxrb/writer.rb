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
      @cells = {}
    end

    # Registers a cell value at the given address (e.g. "A1").
    def set_cell(cell_address, value)
      validate_cell_address!(cell_address)
      @cells[cell_address] = value
    end

    # Returns the registered cells.
    attr_reader :cells

    # Writes the workbook as an XLSX file to the given path.
    def write(filepath)
      entries = {
        "[Content_Types].xml" => generate_content_types_xml,
        "_rels/.rels" => generate_rels_root,
        "xl/workbook.xml" => generate_workbook_xml,
        "xl/worksheets/sheet1.xml" => generate_worksheet1_xml,
        "xl/_rels/workbook.xml.rels" => generate_workbook_rels
      }

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
        %(<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>),
        %(<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>),
        "</Types>"
      ]
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
        "<sheets>",
        %(<sheet name="Sheet1" sheetId="1" r:id="rId1"/>),
        "</sheets>",
        "</workbook>"
      ]
      parts.join
    end

    def generate_workbook_rels
      parts = [
        XML_HEADER,
        %(<Relationships xmlns="#{REL_NS}">),
        %(<Relationship Id="rId1" Type="#{DOC_REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>),
        "</Relationships>"
      ]
      parts.join
    end

    def generate_worksheet1_xml
      parts = [
        XML_HEADER,
        %(<worksheet xmlns="#{SSML_NS}">),
        "<sheetData>"
      ]

      # Group cells by row number.
      cells_by_row = {}
      @cells.each do |address, value|
        row_num = extract_row_number(address)
        col_letter = extract_column_letter(address)
        cells_by_row[row_num] ||= {}
        cells_by_row[row_num][col_letter] = value.to_s
      end

      # Emit rows in ascending order.
      cells_by_row.sort.each do |row_num, row_cells|
        parts << %(<row r="#{row_num}">)
        row_cells.sort.each do |col_letter, value|
          cell_ref = "#{col_letter}#{row_num}"
          parts << %(<c r="#{cell_ref}" t="inlineStr"><is><t>#{xml_escape(value)}</t></is></c>)
        end
        parts << "</row>"
      end

      parts << "</sheetData>"
      parts << "</worksheet>"
      parts.join
    end

    def xml_escape(value)
      value
        .gsub("&", "&amp;")
        .gsub("<", "&lt;")
        .gsub(">", "&gt;")
        .gsub('"', "&quot;")
        .gsub("'", "&apos;")
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
