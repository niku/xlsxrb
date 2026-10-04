# frozen_string_literal: true

require "test_helper"

class DateTimeStyleTest < Test::Unit::TestCase
  test "auto injects default date and time styles" do
    temp = Tempfile.new(["test_dates", ".xlsx"])
    Xlsxrb.write(temp.path) do |w|
      w.sheet("S") do |s|
        s.row([Date.new(2026, 1, 1), Time.new(2026, 1, 1, 12, 0, 0)])
      end
    end
    wb = Xlsxrb.read(temp.path).load
    sheet = wb.sheets.first
    date_cell = sheet.rows.first.cells[0]
    time_cell = sheet.rows.first.cells[1]

    # style_index should be > 0 (0 is normal)
    assert date_cell.style_index.positive?
    assert time_cell.style_index.positive?

    # Ensure numFmt is associated with these styles
    date_xf = wb.styles[:cell_xfs][date_cell.style_index]
    time_xf = wb.styles[:cell_xfs][time_cell.style_index]

    assert date_xf[:num_fmt_id].positive?
    assert time_xf[:num_fmt_id].positive?
  end

  test "stream writer with date1904 property serializes dates with 1904 epoch and reads back correctly" do
    target_date = Date.new(2026, 9, 27)
    target_time = Time.utc(2026, 9, 27, 12, 0, 0)
    temp = Tempfile.new(["test_date1904", ".xlsx"])

    Xlsxrb.write(temp.path) do |w|
      w.workbook_property(:date1904, true)
      w.sheet("Dates1904") do |s|
        s.row([target_date, target_time])
      end
    end

    # Verify workbook.xml contains date1904="1"
    reader = Xlsxrb::Ooxml::Reader.new(temp.path)
    assert_equal(true, reader.workbook_properties[:date1904])
    assert_true(reader.date1904?)

    # Verify resolved date and datetime cells match expected values
    resolved_cells = reader.cells(sheet: "Dates1904")
    assert_equal(target_date, resolved_cells["A1"])
    assert_equal(target_time, resolved_cells["B1"])

    # Verify Xlsxrb.read parses date1904? flag and cells
    wb = Xlsxrb.read(temp.path).load
    assert_true(wb.date1904?)
    sheet = wb["Dates1904"]
    assert_equal(target_date, sheet["A1"].to_date(date1904: true))
    assert_equal(target_time, sheet["B1"].to_time(date1904: true))

    # Verify raw serial values in sheet XML are 1462 less than 1900 system
    sheet_xml = reader.send(:load_worksheet_xml, "Dates1904")
    assert_includes(sheet_xml, "<v>44830</v>")
  end

  test "in-memory Elements::Workbook with date1904 round-trips through Xlsxrb.write and Xlsxrb.read" do
    target_date = Date.new(2026, 9, 27)
    styles = {
      cell_xfs: [
        { num_fmt_id: 0 },
        { num_fmt_id: 14 }
      ],
      num_fmts: {}
    }
    cell = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: target_date, style_index: 1)
    row = Xlsxrb::Elements::Row.new(index: 0, cells: [cell])
    ws = Xlsxrb::Elements::Worksheet.new(name: "Sheet1", rows: [row])
    wb = Xlsxrb::Elements::Workbook.new(
      sheets: [ws],
      styles: styles,
      date1904: true
    )
    assert_true(wb.date1904?)

    temp = Tempfile.new(["test_in_memory_1904", ".xlsx"])
    Xlsxrb.write(temp.path, wb)

    # Read back and verify round-trip
    wb_read = Xlsxrb.read(temp.path).load
    assert_true(wb_read.date1904?)
    assert_equal(target_date, wb_read["Sheet1"]["A1"].to_date(date1904: wb_read.date1904?))

    # Verify reading back in-memory binary string
    wb_buf = Xlsxrb.read(File.binread(temp.path)).load
    assert_true(wb_buf.date1904?)
    assert_equal(target_date, wb_buf["Sheet1"]["A1"].to_date(date1904: wb_buf.date1904?))

    reader = Xlsxrb::Ooxml::Reader.new(temp.path)
    assert_equal(true, reader.workbook_properties[:date1904])
    assert_true(reader.date1904?)
    assert_equal(target_date, reader.cells(sheet: "Sheet1")["A1"])
  end

  test "DOM Ooxml::Writer with date1904 serializes dates with 1904 offset" do
    writer = Xlsxrb::Ooxml::Writer.new
    writer.workbook_property(:date1904, true)
    target_date = Date.new(2026, 9, 27)
    target_time = Time.utc(2026, 9, 27, 12, 0, 0)
    writer.set_cell("A1", target_date)
    writer.set_cell("B1", target_time)

    temp = Tempfile.new(["test_writer_1904", ".xlsx"])
    writer.write(temp.path)

    reader = Xlsxrb::Ooxml::Reader.new(temp.path)
    assert_equal(true, reader.workbook_properties[:date1904])
    cells = reader.cells(sheet: "Sheet1")
    assert_equal(target_date, cells["A1"])
    assert_equal(target_time, cells["B1"])
  end
end
