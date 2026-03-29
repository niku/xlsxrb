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

  test "rejects invalid cell addresses" do
    writer = Xlsxrb::Writer.new
    assert_raise(ArgumentError) { writer.set_cell("", "v") }
    assert_raise(ArgumentError) { writer.set_cell("1A", "v") }
    assert_raise(ArgumentError) { writer.set_cell("A0", "v") }
    assert_raise(ArgumentError) { writer.set_cell("a1", "v") }
    assert_raise(ArgumentError) { writer.set_cell("XFE1", "v") }
    assert_raise(ArgumentError) { writer.set_cell("A1048577", "v") }
  end

  test "accepts valid boundary cell addresses" do
    writer = Xlsxrb::Writer.new
    assert_nothing_raised { writer.set_cell("A1", "v") }
    assert_nothing_raised { writer.set_cell("XFD1048576", "v") }
    assert_nothing_raised { writer.set_cell("Z1", "v") }
    assert_nothing_raised { writer.set_cell("AA1", "v") }
  end

  test "stores numeric values" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 42)
    writer.set_cell("B1", 3.14)

    assert_equal({ "A1" => 42, "B1" => 3.14 }, writer.cells)
  end

  test "stores boolean values" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", true)
    writer.set_cell("B1", false)

    assert_equal({ "A1" => true, "B1" => false }, writer.cells)
  end

  test "stores empty string as a cell value" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "")

    assert_equal({ "A1" => "" }, writer.cells)
  end

  test "orders cells by column index within a row" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("AA1", "third")
    writer.set_cell("B1", "second")
    writer.set_cell("A1", "first")

    assert_equal({ "A1" => "first", "B1" => "second", "AA1" => "third" }, writer.cells)
  end

  test "adds multiple sheets" do
    writer = Xlsxrb::Writer.new
    writer.add_sheet("Data")
    writer.set_cell("A1", "main", sheet: "Sheet1")
    writer.set_cell("A1", "data", sheet: "Data")

    assert_equal(%w[Sheet1 Data], writer.sheet_order)
    assert_equal({ "A1" => "main" }, writer.cells(sheet: "Sheet1"))
    assert_equal({ "A1" => "data" }, writer.cells(sheet: "Data"))
  end

  test "rejects duplicate sheet names" do
    writer = Xlsxrb::Writer.new
    assert_raise(ArgumentError) { writer.add_sheet("Sheet1") }
  end

  test "rejects unknown sheet in set_cell" do
    writer = Xlsxrb::Writer.new
    assert_raise(ArgumentError) { writer.set_cell("A1", "v", sheet: "NoSuchSheet") }
  end

  test "stores column widths" do
    writer = Xlsxrb::Writer.new
    writer.set_column_width("A", 20)
    writer.set_column_width("C", 15.5)

    assert_equal({ "A" => 20, "C" => 15.5 }, writer.column_widths)
  end

  test "rejects invalid column letter in set_column_width" do
    writer = Xlsxrb::Writer.new
    assert_raise(ArgumentError) { writer.set_column_width("a", 10) }
    assert_raise(ArgumentError) { writer.set_column_width("1", 10) }
    assert_raise(ArgumentError) { writer.set_column_width("", 10) }
  end

  test "stores row height and hidden attributes" do
    writer = Xlsxrb::Writer.new
    writer.set_row_height(1, 25.0)
    writer.set_row_hidden(3)

    expected = { 1 => { height: 25.0 }, 3 => { hidden: true } }
    assert_equal(expected, writer.row_attributes)
  end

  test "stores merge cell ranges" do
    writer = Xlsxrb::Writer.new
    writer.merge_cells("A1:B2")
    writer.merge_cells("C3:D4")

    assert_equal(%w[A1:B2 C3:D4], writer.merged_cells)
  end

  test "rejects invalid merge cell range" do
    writer = Xlsxrb::Writer.new
    assert_raise(ArgumentError) { writer.merge_cells("A1") }
    assert_raise(ArgumentError) { writer.merge_cells("") }
  end

  test "stores hyperlinks" do
    writer = Xlsxrb::Writer.new
    writer.add_hyperlink("A1", "https://example.com")
    writer.add_hyperlink("B1", "https://github.com")

    assert_equal({ "A1" => "https://example.com", "B1" => "https://github.com" }, writer.hyperlinks)
  end

  test "adds number formats and assigns to cells" do
    writer = Xlsxrb::Writer.new
    fmt_id = writer.add_number_format("0.00")
    writer.set_cell("A1", 3.14)
    writer.set_cell_format("A1", fmt_id)

    assert_equal(164, fmt_id)

    # Same format code returns same id.
    assert_equal(164, writer.add_number_format("0.00"))

    # Different format gets next id.
    assert_equal(165, writer.add_number_format("#,##0"))
  end

  test "stores Date values" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", Date.new(2024, 1, 15))

    assert_equal({ "A1" => Date.new(2024, 1, 15) }, writer.cells)
  end

  test "stores auto filter range" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Name")
    writer.set_cell("B1", "Age")
    writer.set_auto_filter("A1:B10")

    assert_equal("A1:B10", writer.auto_filter)
  end

  test "rejects invalid auto filter range" do
    writer = Xlsxrb::Writer.new
    assert_raise(ArgumentError) { writer.set_auto_filter("A1") }
    assert_raise(ArgumentError) { writer.set_auto_filter("") }
  end

  test "stores core properties" do
    writer = Xlsxrb::Writer.new
    writer.set_core_property(:title, "My Workbook")
    writer.set_core_property(:creator, "Test User")
    writer.set_core_property(:created, "2024-01-15T00:00:00Z")
    writer.set_core_property(:modified, "2024-01-16T12:00:00Z")

    props = writer.core_properties
    assert_equal("My Workbook", props[:title])
    assert_equal("Test User", props[:creator])
    assert_equal("2024-01-15T00:00:00Z", props[:created])
    assert_equal("2024-01-16T12:00:00Z", props[:modified])
  end

  test "stores app properties" do
    writer = Xlsxrb::Writer.new
    writer.set_app_property(:application, "Xlsxrb")
    writer.set_app_property(:app_version, "1.0.0")
    writer.set_app_property(:heading_pairs, [["Worksheets", 2]])
    writer.set_app_property(:titles_of_parts, %w[Sheet1 Data])

    props = writer.app_properties
    assert_equal("Xlsxrb", props[:application])
    assert_equal("1.0.0", props[:app_version])
    assert_equal([["Worksheets", 2]], props[:heading_pairs])
    assert_equal(%w[Sheet1 Data], props[:titles_of_parts])
  end

  test "stores workbook properties" do
    writer = Xlsxrb::Writer.new
    writer.set_workbook_property(:date1904, false)
    writer.set_workbook_property(:default_theme_version, 166_925)

    props = writer.workbook_properties
    assert_equal(false, props[:date1904])
    assert_equal(166_925, props[:default_theme_version])
  end

  test "stores workbook view properties" do
    writer = Xlsxrb::Writer.new
    writer.set_workbook_view(:active_tab, 1)
    writer.set_workbook_view(:first_sheet, 0)

    views = writer.workbook_views
    assert_equal(1, views[:active_tab])
    assert_equal(0, views[:first_sheet])
  end

  test "stores calc properties" do
    writer = Xlsxrb::Writer.new
    writer.set_calc_property(:calc_id, 191_029)
    writer.set_calc_property(:full_calc_on_load, true)

    props = writer.calc_properties
    assert_equal(191_029, props[:calc_id])
    assert_equal(true, props[:full_calc_on_load])
  end

  test "sets sheet state" do
    writer = Xlsxrb::Writer.new
    writer.add_sheet("Hidden")
    writer.add_sheet("VeryHidden")
    writer.set_sheet_state("Hidden", :hidden)
    writer.set_sheet_state("VeryHidden", :very_hidden)

    assert_equal(:visible, writer.sheet_state("Sheet1"))
    assert_equal(:hidden, writer.sheet_state("Hidden"))
    assert_equal(:very_hidden, writer.sheet_state("VeryHidden"))
  end

  test "rejects invalid sheet state" do
    writer = Xlsxrb::Writer.new
    assert_raise(ArgumentError) { writer.set_sheet_state("Sheet1", :invalid) }
    assert_raise(ArgumentError) { writer.set_sheet_state("NoSuch", :hidden) }
  end
end
