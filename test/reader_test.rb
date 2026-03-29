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
end
