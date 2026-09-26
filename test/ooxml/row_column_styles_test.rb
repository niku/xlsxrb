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
end
