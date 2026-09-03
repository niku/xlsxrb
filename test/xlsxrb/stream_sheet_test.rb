# frozen_string_literal: true

require "test_helper"

class StreamSheetTest < Test::Unit::TestCase
  cover Xlsxrb::StreamSheet

  test "stream_sheet each_row, each_cell, and load into Elements::Worksheet" do
    xml = <<~XML
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
          <row r="1">
            <c r="A1"><v>100</v></c>
            <c r="B1" t="s"><v>0</v></c>
          </row>
          <row r="2">
            <c r="A2"><v>200</v></c>
          </row>
        </sheetData>
      </worksheet>
    XML

    sheet = Xlsxrb::StreamSheet.new("DataSheet", xml.b, ["Hello"])

    assert_equal("DataSheet", sheet.name)

    # Enumerators
    assert_instance_of(Enumerator, sheet.each_row)
    assert_instance_of(Enumerator, sheet.each_cell)

    # each_row enumeration
    rows = sheet.each_row.to_a
    assert_equal(2, rows.size)
    assert_equal(0, rows[0].index)
    assert_equal(1, rows[1].index)

    # each_cell enumeration across rows
    cells = sheet.each_cell.to_a
    assert_equal(3, cells.size)
    assert_equal(100, cells[0].value)
    assert_equal("Hello", cells[1].value)
    assert_equal(200, cells[2].value)

    # Enumerable each delegation
    assert_equal(2, sheet.count)
    assert_equal([0, 1], sheet.map(&:index))

    # load into in-memory Worksheet
    ws = sheet.load
    assert_instance_of(Xlsxrb::Elements::Worksheet, ws)
    assert_equal("DataSheet", ws.name)
    assert_equal(100, ws["A1"].value)
    assert_equal("Hello", ws["B1"].value)
    assert_equal(200, ws["A2"].value)

    # to_worksheet alias
    assert_equal(ws["A1"].value, sheet.to_worksheet["A1"].value)
  end
end
