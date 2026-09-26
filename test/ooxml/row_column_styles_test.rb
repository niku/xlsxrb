# frozen_string_literal: true

require "test_helper"

class RowColumnStylesTest < Test::Unit::TestCase
  def test_row_and_column_styles_roundtrip
    Dir.mktmpdir do |dir|
      filepath = File.join(dir, "test.xlsx")
      # Write a workbook with row styles, column styles, and border color
      styles = {
        fonts: [{ name: "Calibri", sz: 11 }, { name: "Arial", sz: 14, bold: true }],
        fills: [{ pattern: "none" }, { pattern: "gray125" }, { pattern: "solid", fg_color: "FFFFFF00" }],
        borders: [{ left: { style: "thin", color: "FFFF0000" } }],
        cell_xfs: [
          { font_id: 0, fill_id: 0, border_id: 0 },
          { font_id: 1, fill_id: 2, border_id: 0 }
        ]
      }

      wb = Xlsxrb::Elements::Workbook.new(
        sheets: [
          Xlsxrb::Elements::Worksheet.new(
            name: "StyledSheet",
            rows: [
              Xlsxrb::Elements::Row.new(
                index: 0,
                cells: [Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "Hello")],
                unmapped_data: { style_index: 1 }
              )
            ],
            columns: [
              Xlsxrb::Elements::Column.new(
                index: 0,
                width: 20.0,
                unmapped_data: { style_index: 1 }
              )
            ]
          )
        ],
        styles: styles
      )

      Xlsxrb.write(filepath, wb)

      # Read back and verify styles, row style_index, and column style_index
      parsed_wb = Xlsxrb.read(filepath).load
      parsed_ws = parsed_wb.sheets.first
      refute_nil parsed_ws

      # Verify row style_index
      assert_equal(1, parsed_ws.rows.first.style_index)

      # Verify column style_index
      assert_equal(1, parsed_ws.columns.first.style_index)

      # Verify border color parsed
      parsed_styles = parsed_wb.styles
      refute_nil parsed_styles
      border = parsed_styles[:borders]&.first
      assert_equal("FFFF0000", border&.dig(:left, :color, :rgb))
    end
  end

  def test_workbook_writer_num_fmts_hash_and_array_support
    Dir.mktmpdir do |dir|
      # Test with Hash format (as returned by Reader)
      filepath_hash = File.join(dir, "hash_styles.xlsx")
      wb_hash = Xlsxrb::Elements::Workbook.new(
        sheets: [Xlsxrb::Elements::Worksheet.new(name: "S", rows: [Xlsxrb::Elements::Row.new(index: 0)])],
        styles: { num_fmts: { 164 => "#,##0.00" } }
      )
      Xlsxrb.write(filepath_hash, wb_hash)

      entries = Xlsxrb::Ooxml::ZipReader.open(filepath_hash, &:read_all)
      styles_xml = entries["xl/styles.xml"]
      assert_match(%r{<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0\.00"/></numFmts>}, styles_xml)

      parsed = Xlsxrb.read(filepath_hash).load
      assert_equal("#,##0.00", parsed.styles.dig(:num_fmts, 164))

      # Test with Array format
      filepath_arr = File.join(dir, "arr_styles.xlsx")
      wb_arr = Xlsxrb::Elements::Workbook.new(
        sheets: [Xlsxrb::Elements::Worksheet.new(name: "S", rows: [Xlsxrb::Elements::Row.new(index: 0)])],
        styles: { num_fmts: [{ num_fmt_id: 165, format_code: "$#,##0" }] }
      )
      Xlsxrb.write(filepath_arr, wb_arr)

      entries_arr = Xlsxrb::Ooxml::ZipReader.open(filepath_arr, &:read_all)
      styles_xml_arr = entries_arr["xl/styles.xml"]
      assert_match(%r{<numFmts count="1"><numFmt numFmtId="165" formatCode="\$#,##0"/></numFmts>}, styles_xml_arr)
    end
  end
end
