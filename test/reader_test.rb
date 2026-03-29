# frozen_string_literal: true

require "test_helper"
require "tempfile"

class ReaderTest < Test::Unit::TestCase
  test "reads inline string written by writer" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    assert_equal({ "A1" => "hello" }, reader.cells)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "reads multiple inline string cells" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("B2", "world")
    writer.set_cell("A1", "hello")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    assert_equal({ "A1" => "hello", "B2" => "world" }, reader.cells)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "reads multiple inline string cells in the same row" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("B1", "world")
    writer.set_cell("A1", "hello")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    assert_equal({ "A1" => "hello", "B1" => "world" }, reader.cells)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips cells with multi-letter column references" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("AA1", "third")
    writer.set_cell("B1", "second")
    writer.set_cell("A1", "first")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    assert_equal({ "A1" => "first", "B1" => "second", "AA1" => "third" }, reader.cells)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips numeric cells" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 42)
    writer.set_cell("B1", 3.14)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    assert_equal({ "A1" => 42, "B1" => 3.14 }, reader.cells)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips boolean cells" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", true)
    writer.set_cell("B1", false)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    assert_equal({ "A1" => true, "B1" => false }, reader.cells)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips empty string cells" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "")
    writer.set_cell("B1", "hello")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    assert_equal({ "A1" => "", "B1" => "hello" }, reader.cells)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips formula cells" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.set_cell("A2", 20)
    writer.set_cell("A3", Xlsxrb::Formula.new(expression: "SUM(A1:A2)", cached_value: "30"))
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cells = reader.cells
    assert_equal(10, cells["A1"])
    assert_equal(20, cells["A2"])
    assert_instance_of(Xlsxrb::Formula, cells["A3"])
    assert_equal("SUM(A1:A2)", cells["A3"].expression)
    assert_equal("30", cells["A3"].cached_value)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips multiple sheets" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.add_sheet("Data")
    writer.set_cell("A1", "main", sheet: "Sheet1")
    writer.set_cell("A1", "data", sheet: "Data")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    assert_equal(%w[Sheet1 Data], reader.sheet_names)
    assert_equal({ "A1" => "main" }, reader.cells)
    assert_equal({ "A1" => "main" }, reader.cells(sheet: "Sheet1"))
    assert_equal({ "A1" => "data" }, reader.cells(sheet: "Data"))
    assert_equal({ "A1" => "main" }, reader.cells(sheet: 0))
    assert_equal({ "A1" => "data" }, reader.cells(sheet: 1))
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips column widths" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_column_width("A", 20.0)
    writer.set_column_width("C", 15.5)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    assert_equal({ "A" => 20.0, "C" => 15.5 }, reader.columns)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips row attributes" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_row_height(1, 25.0)
    writer.set_row_hidden(3)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    row_attrs = reader.row_attributes
    assert_in_delta(25.0, row_attrs[1][:height], 0.01)
    assert_equal(true, row_attrs[3][:hidden])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips merge cells" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.merge_cells("A1:B2")
    writer.merge_cells("C3:D4")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    assert_equal(%w[A1:B2 C3:D4], reader.merged_cells)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips hyperlinks" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Example")
    writer.add_hyperlink("A1", "https://example.com")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    assert_equal({ "A1" => "https://example.com" }, reader.hyperlinks)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips cell number formats" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    fmt_id = writer.add_number_format("0.00")
    writer.set_cell("A1", 3.14)
    writer.set_cell_format("A1", fmt_id)
    writer.set_cell("B1", 42)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    assert_equal({ "A1" => "0.00" }, reader.cell_formats)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips Date cells" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", Date.new(2024, 1, 15))
    writer.set_cell("B1", 42)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cells = reader.cells
    assert_equal(Date.new(2024, 1, 15), cells["A1"])
    assert_equal(42, cells["B1"])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "serial_to_date and date_to_serial round-trip" do
    date = Date.new(2024, 6, 15)
    serial = Xlsxrb.date_to_serial(date)
    assert_equal(date, Xlsxrb.serial_to_date(serial))

    # Excel epoch: Jan 1, 1900 = serial 1
    assert_equal(1, Xlsxrb.date_to_serial(Date.new(1900, 1, 1)))
    assert_equal(Date.new(1900, 1, 1), Xlsxrb.serial_to_date(1))
  end

  test "round-trips auto filter" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Name")
    writer.set_cell("B1", "Age")
    writer.set_auto_filter("A1:B10")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    assert_equal("A1:B10", reader.auto_filter)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips core properties" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_core_property(:title, "My Workbook")
    writer.set_core_property(:creator, "Test User")
    writer.set_core_property(:created, "2024-01-15T00:00:00Z")
    writer.set_core_property(:modified, "2024-01-16T12:00:00Z")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    props = reader.core_properties
    assert_equal("My Workbook", props[:title])
    assert_equal("Test User", props[:creator])
    assert_equal("2024-01-15T00:00:00Z", props[:created])
    assert_equal("2024-01-16T12:00:00Z", props[:modified])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end
end
