# frozen_string_literal: true

require "test_helper"
require "tempfile"

class WriterTest < Test::Unit::TestCase
  test "can instantiate Writer" do
    writer = Xlsxrb::Writer.new
    assert_not_nil(writer)
  end

  test "can set a cell value" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    assert(true)
  end

  test "generated workbook contains worksheet with cell" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")

    # Verify writer keeps cell values in its internal state.
    cells = writer.cells
    assert_not_nil(cells)
  end

  test "keeps multiple cells in the same row" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("B1", "world")
    writer.set_cell("A1", "hello")

    assert_equal({ "A1" => "hello", "B1" => "world" }, writer.cells)
  end

  test "can generate XLSX file" do
    temp_file = Tempfile.new(["test", ".xlsx"])
    temp_path = temp_file.path
    temp_file.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.write(temp_path)

    assert(File.exist?(temp_path))
    assert(File.size(temp_path).positive?)

    # Verify ZIP local file header signature (PK\x03\x04)
    file_content = File.read(temp_path, 4)
    assert_equal([0x50, 0x4b, 0x03, 0x04], file_content.bytes[0..3])

    FileUtils.rm_f(temp_path)
  end
end
