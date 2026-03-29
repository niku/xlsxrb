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
end
