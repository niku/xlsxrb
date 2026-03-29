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

  test "round-trips app properties" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_app_property(:application, "Xlsxrb")
    writer.set_app_property(:app_version, "1.0.0")
    writer.set_app_property(:heading_pairs, [["Worksheets", 1]])
    writer.set_app_property(:titles_of_parts, ["Sheet1"])
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    props = reader.app_properties
    assert_equal("Xlsxrb", props[:application])
    assert_equal("1.0.0", props[:app_version])
    assert_equal([["Worksheets", 1]], props[:heading_pairs])
    assert_equal(["Sheet1"], props[:titles_of_parts])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips workbook properties and views" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.add_sheet("Data")
    writer.set_cell("A1", "hello", sheet: "Sheet1")
    writer.set_workbook_property(:date1904, false)
    writer.set_workbook_property(:default_theme_version, 166_925)
    writer.set_workbook_view(:active_tab, 1)
    writer.set_workbook_view(:first_sheet, 0)
    writer.set_calc_property(:calc_id, 191_029)
    writer.set_calc_property(:full_calc_on_load, true)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    wp = reader.workbook_properties
    assert_equal(false, wp[:date1904])
    assert_equal(166_925, wp[:default_theme_version])

    wv = reader.workbook_views
    assert_equal(1, wv[:active_tab])
    assert_equal(0, wv[:first_sheet])

    cp = reader.calc_properties
    assert_equal(191_029, cp[:calc_id])
    assert_equal(true, cp[:full_calc_on_load])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips sheet states" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.add_sheet("Hidden")
    writer.add_sheet("VeryHidden")
    writer.set_cell("A1", "main", sheet: "Sheet1")
    writer.set_sheet_state("Hidden", :hidden)
    writer.set_sheet_state("VeryHidden", :very_hidden)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    states = reader.sheet_states
    assert_equal(:visible, states["Sheet1"])
    assert_equal(:hidden, states["Hidden"])
    assert_equal(:very_hidden, states["VeryHidden"])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips defined names" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.add_sheet("Data")
    writer.set_cell("A1", "hello", sheet: "Sheet1")
    writer.add_defined_name("MyRange", "Sheet1!$A$1:$B$10")
    writer.add_defined_name("LocalName", "Data!$C$1", sheet: "Data")
    writer.add_defined_name("HiddenName", "42", hidden: true)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    dns = reader.defined_names
    assert_equal(3, dns.size)

    assert_equal("MyRange", dns[0][:name])
    assert_equal("Sheet1!$A$1:$B$10", dns[0][:value])
    assert_nil(dns[0][:local_sheet_id])
    assert_equal(false, dns[0][:hidden])

    assert_equal("LocalName", dns[1][:name])
    assert_equal(1, dns[1][:local_sheet_id])

    assert_equal("HiddenName", dns[2][:name])
    assert_equal(true, dns[2][:hidden])
    assert_equal("42", dns[2][:value])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips sheet properties" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_sheet_property(:tab_color, "FF0000FF")
    writer.set_sheet_property(:summary_below, false)
    writer.set_sheet_property(:summary_right, true)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    props = reader.sheet_properties
    assert_equal("FF0000FF", props[:tab_color])
    assert_equal(false, props[:summary_below])
    assert_equal(true, props[:summary_right])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips dimension" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("B2", "hello")
    writer.set_cell("D5", "world")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    assert_equal("B2:D5", reader.dimension)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips sheet format properties" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_sheet_format(:default_row_height, 18.0)
    writer.set_sheet_format(:default_col_width, 12.5)
    writer.set_sheet_format(:base_col_width, 10)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    fmt = reader.sheet_format
    assert_equal(18.0, fmt[:default_row_height])
    assert_equal(12.5, fmt[:default_col_width])
    assert_equal(10, fmt[:base_col_width])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips row outline level and collapsed" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_row_outline_level(2, 1)
    writer.set_row_collapsed(3)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    attrs = reader.row_attributes
    assert_equal(1, attrs[2][:outline_level])
    assert_equal(true, attrs[3][:collapsed])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips column attributes" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_column_attribute("B", :hidden, true)
    writer.set_column_attribute("C", :outline_level, 2)
    writer.set_column_attribute("C", :collapsed, true)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    ca = reader.column_attributes
    assert_equal(true, ca["B"][:hidden])
    assert_equal(2, ca["C"][:outline_level])
    assert_equal(true, ca["C"][:collapsed])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips sheet view properties" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_sheet_view(:show_grid_lines, false)
    writer.set_sheet_view(:zoom_scale, 150)
    writer.set_sheet_view(:right_to_left, true)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    sv = reader.sheet_view
    assert_equal(false, sv[:show_grid_lines])
    assert_equal(150, sv[:zoom_scale])
    assert_equal(true, sv[:right_to_left])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips freeze pane" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_freeze_pane(row: 1, col: 1)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    fp = reader.freeze_pane
    assert_equal(1, fp[:row])
    assert_equal(1, fp[:col])
    assert_equal(:frozen, fp[:state])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips selection" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_selection("B2", sqref: "B2:C3")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    sel = reader.selection
    assert_equal("B2", sel[:active_cell])
    assert_equal("B2:C3", sel[:sqref])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips print options and page margins" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_print_option(:grid_lines, true)
    writer.set_print_option(:horizontal_centered, true)
    writer.set_page_margins(left: 0.7, right: 0.7, top: 0.75, bottom: 0.75)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    po = reader.print_options
    assert_equal(true, po[:grid_lines])
    assert_equal(true, po[:horizontal_centered])

    pm = reader.page_margins
    assert_in_delta(0.7, pm[:left], 0.001)
    assert_in_delta(0.75, pm[:top], 0.001)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips page setup and header footer" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_page_setup(:orientation, "landscape")
    writer.set_page_setup(:paper_size, 9)
    writer.set_header_footer(:odd_header, "&CPage &P")
    writer.set_header_footer(:odd_footer, "&CFooter")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    ps = reader.page_setup
    assert_equal("landscape", ps[:orientation])
    assert_equal(9, ps[:paper_size])

    hf = reader.header_footer
    assert_equal("&CPage &P", hf[:odd_header])
    assert_equal("&CFooter", hf[:odd_footer])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips row and col breaks" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.add_row_break(10)
    writer.add_row_break(20)
    writer.add_col_break(5)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    assert_equal([10, 20], reader.row_breaks)
    assert_equal([5], reader.col_breaks)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end
end
