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
    writer.set_row_style(5, 0)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    row_attrs = reader.row_attributes
    assert_in_delta(25.0, row_attrs[1][:height], 0.01)
    assert_equal(true, row_attrs[3][:hidden])
    assert_equal(0, row_attrs[5][:style])
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
    assert_equal({ "A1" => { url: "https://example.com" } }, reader.hyperlinks)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips hyperlinks with display tooltip and location" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Example")
    writer.add_hyperlink("A1", "https://example.com", display: "Example Site", tooltip: "Click to visit")
    writer.set_cell("B1", "Page")
    writer.add_hyperlink("B1", "https://example.com/page", location: "Sheet2!A1")
    writer.set_cell("C1", "Internal")
    writer.add_hyperlink("C1", location: "Sheet1!D1")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    expected = {
      "A1" => { url: "https://example.com", display: "Example Site", tooltip: "Click to visit" },
      "B1" => { url: "https://example.com/page", location: "Sheet2!A1" },
      "C1" => { location: "Sheet1!D1" }
    }
    assert_equal(expected, reader.hyperlinks)
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

  test "resolves built-in number formats from numFmtId" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    builtin_date_style = writer.add_cell_style(num_fmt_id: 14)
    builtin_text_style = writer.add_cell_style(num_fmt_id: 49)
    writer.set_cell("A1", 45_292)
    writer.set_cell_style("A1", builtin_date_style)
    writer.set_cell("B1", "text")
    writer.set_cell_style("B1", builtin_text_style)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    assert_equal("mm-dd-yy", reader.cell_formats["A1"])
    assert_equal("@", reader.cell_formats["B1"])
    cs = reader.cell_styles
    assert_equal("mm-dd-yy", cs["A1"][:num_fmt])
    assert_equal("@", cs["B1"][:num_fmt])
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
    assert_equal([10, 20], reader.row_breaks.map { |b| b[:id] })
    assert_equal([5], reader.col_breaks.map { |b| b[:id] })
    assert_equal(true, reader.row_breaks.first[:man])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips filter columns" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_auto_filter("A1:C10")
    writer.add_filter_column(0, { type: :filters, values: %w[A B] })
    writer.add_filter_column(1, { type: :custom, operator: "greaterThan", val: "100" })
    writer.add_filter_column(2, { type: :top10, top: true, percent: false, val: 5 })
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    fc = reader.filter_columns
    assert_equal(:filters, fc[0][:type])
    assert_equal(%w[A B], fc[0][:values])
    assert_equal(:custom, fc[1][:type])
    assert_equal("greaterThan", fc[1][:operator])
    assert_equal("100", fc[1][:val])
    assert_equal(:top10, fc[2][:type])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips sort state" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_sort_state("A1:B10", [{ ref: "A1:A10" }, { ref: "B1:B10", descending: true }])
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    ss = reader.sort_state
    assert_equal("A1:B10", ss[:ref])
    assert_equal(2, ss[:sort_conditions].size)
    assert_equal("A1:A10", ss[:sort_conditions][0][:ref])
    assert_equal(true, ss[:sort_conditions][1][:descending])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips data validations" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.add_data_validation("A1:A100", type: "whole", operator: "between",
                                          formula1: "1", formula2: "100",
                                          show_error_message: true, error: "Must be 1-100")
    writer.add_data_validation("B1:B100", type: "list", formula1: '"Yes,No"',
                                          show_input_message: true, prompt: "Choose one")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    dvs = reader.data_validations
    assert_equal(2, dvs.size)
    assert_equal("A1:A100", dvs[0][:sqref])
    assert_equal("whole", dvs[0][:type])
    assert_equal("between", dvs[0][:operator])
    assert_equal("1", dvs[0][:formula1])
    assert_equal("100", dvs[0][:formula2])
    assert_equal(true, dvs[0][:show_error_message])
    assert_equal("Must be 1-100", dvs[0][:error])
    assert_equal("B1:B100", dvs[1][:sqref])
    assert_equal("list", dvs[1][:type])
    assert_equal(true, dvs[1][:show_input_message])
    assert_equal("Choose one", dvs[1][:prompt])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips data validation deep attributes" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 50)
    writer.add_data_validation("A1:A100", type: "whole", operator: "between",
                                          formula1: "1", formula2: "100",
                                          allow_blank: true,
                                          error_style: "warning",
                                          error_title: "Bad Value",
                                          error: "Please enter 1-100",
                                          show_error_message: true,
                                          prompt_title: "Input Needed",
                                          prompt: "Enter a number",
                                          show_input_message: true)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    dvs = reader.data_validations
    assert_equal(1, dvs.size)
    dv = dvs[0]
    assert_equal(true, dv[:allow_blank])
    assert_equal("warning", dv[:error_style])
    assert_equal("Bad Value", dv[:error_title])
    assert_equal("Please enter 1-100", dv[:error])
    assert_equal("Input Needed", dv[:prompt_title])
    assert_equal("Enter a number", dv[:prompt])
    assert_equal(true, dv[:show_error_message])
    assert_equal(true, dv[:show_input_message])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips cellStyleXfs and named cell styles" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    fid = writer.add_font(bold: true, sz: 14, name: "Arial")
    writer.add_named_cell_style(name: "Heading1", font_id: fid, builtin_id: 1)
    writer.set_cell("A1", "Hello")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    csxfs = reader.cell_style_xfs
    assert_equal(2, csxfs.size)
    assert_equal(fid, csxfs[1][:font_id])

    ncs = reader.named_cell_styles
    assert_equal(2, ncs.size)
    assert_equal("Normal", ncs[0][:name])
    assert_equal(0, ncs[0][:builtin_id])
    assert_equal("Heading1", ncs[1][:name])
    assert_equal(1, ncs[1][:xf_id])
    assert_equal(1, ncs[1][:builtin_id])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "resolves cellXf xfId linkage through cellStyleXfs" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    style_xf_id = writer.add_named_cell_style(name: "DateBase", num_fmt_id: 14)
    cell_xf_id = writer.add_cell_style(xf_id: style_xf_id)
    writer.set_cell("A1", 45_292)
    writer.set_cell_style("A1", cell_xf_id)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    assert_equal("mm-dd-yy", reader.cell_formats["A1"])
    assert_equal("mm-dd-yy", reader.cell_styles["A1"][:num_fmt])
    assert_equal(Date.new(2024, 1, 1), reader.cells["A1"])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips conditional formatting rules" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.add_conditional_format("A1:A10", type: :cell_is, operator: "greaterThan",
                                            formula: "100", priority: 1, format_id: 0)
    writer.add_conditional_format("B1:B10", type: :color_scale, priority: 2,
                                            color_scale: {
                                              cfvo: [{ type: "min" }, { type: "max" }],
                                              colors: %w[FF0000FF FFFF0000]
                                            })
    writer.add_conditional_format("C1:C10", type: :data_bar, priority: 3,
                                            data_bar: {
                                              cfvo: [{ type: "min" }, { type: "max" }],
                                              color: "FF638EC6"
                                            })
    writer.add_conditional_format("D1:D10", type: :icon_set, priority: 4,
                                            icon_set: {
                                              icon_set: "3TrafficLights1",
                                              cfvo: [{ type: "percent", val: "0" },
                                                     { type: "percent", val: "33" },
                                                     { type: "percent", val: "67" }]
                                            })
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cfs = reader.conditional_formats
    assert_equal(4, cfs.size)

    assert_equal("A1:A10", cfs[0][:sqref])
    assert_equal("cellIs", cfs[0][:type])
    assert_equal("greaterThan", cfs[0][:operator])
    assert_equal(1, cfs[0][:priority])
    assert_equal(0, cfs[0][:format_id])
    assert_equal(["100"], cfs[0][:formulas])

    assert_equal("colorScale", cfs[1][:type])
    assert_equal(2, cfs[1][:color_scale][:cfvo].size)
    assert_equal("min", cfs[1][:color_scale][:cfvo][0][:type])
    assert_equal([{ rgb: "FF0000FF" }, { rgb: "FFFF0000" }], cfs[1][:color_scale][:colors])

    assert_equal("dataBar", cfs[2][:type])
    assert_equal(2, cfs[2][:data_bar][:cfvo].size)
    assert_equal({ rgb: "FF638EC6" }, cfs[2][:data_bar][:color])

    assert_equal("iconSet", cfs[3][:type])
    assert_equal("3TrafficLights1", cfs[3][:icon_set][:icon_set])
    assert_equal(3, cfs[3][:icon_set][:cfvo].size)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips dataBar deep attributes" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 50)
    writer.add_conditional_format("A1:A10", type: :data_bar, priority: 1,
                                            data_bar: {
                                              cfvo: [{ type: "min" }, { type: "max" }],
                                              color: "FF638EC6",
                                              min_length: 5, max_length: 90, show_value: false
                                            })
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cfs = reader.conditional_formats
    assert_equal(1, cfs.size)
    db = cfs[0][:data_bar]
    assert_equal(5, db[:min_length])
    assert_equal(90, db[:max_length])
    assert_equal(false, db[:show_value])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips iconSet deep attributes" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 50)
    writer.add_conditional_format("A1:A10", type: :icon_set, priority: 1,
                                            icon_set: {
                                              icon_set: "3Arrows",
                                              cfvo: [{ type: "percent", val: "0" },
                                                     { type: "percent", val: "33" },
                                                     { type: "percent", val: "67" }],
                                              reverse: true, show_value: false
                                            })
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cfs = reader.conditional_formats
    assert_equal(1, cfs.size)
    is = cfs[0][:icon_set]
    assert_equal("3Arrows", is[:icon_set])
    assert_equal(true, is[:reverse])
    assert_equal(false, is[:show_value])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips cell styles with fonts fills and borders" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    fid = writer.add_font(bold: true, sz: 14, name: "Arial", color: "FFFF0000")
    fill_id = writer.add_fill(pattern: "solid", fg_color: "FF00FF00")
    brd_id = writer.add_border(left: { style: "thin", color: "FF000000" },
                               right: { style: "thin" },
                               top: { style: "thin" },
                               bottom: { style: "thin" })
    style_id = writer.add_cell_style(font_id: fid, fill_id: fill_id, border_id: brd_id)
    writer.set_cell("A1", "styled")
    writer.set_cell_style("A1", style_id)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cs = reader.cell_styles
    assert(cs.key?("A1"), "A1 should have a style")
    assert_equal(true, cs["A1"][:font][:bold])
    assert_equal(14.0, cs["A1"][:font][:sz])
    assert_equal("FFFF0000", cs["A1"][:font][:color])
    assert_equal("solid", cs["A1"][:fill][:pattern])
    assert_equal("FF00FF00", cs["A1"][:fill][:fg_color])
    assert_equal("thin", cs["A1"][:border][:left][:style])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips dxf entries" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.add_dxf(font: { bold: true, color: "FFFF0000" },
                   fill: { pattern: "solid", fg_color: "FFFFFF00" })
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    dxfs = reader.dxfs
    assert_equal(1, dxfs.size)
    assert_equal(true, dxfs[0][:font][:bold])
    assert_equal("FFFF0000", dxfs[0][:font][:color])
    assert_equal("solid", dxfs[0][:fill][:pattern])
    assert_equal("FFFFFF00", dxfs[0][:fill][:fg_color])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips tables" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Name")
    writer.set_cell("B1", "Age")
    writer.set_cell("A2", "Alice")
    writer.set_cell("B2", 30)
    writer.add_table("A1:B2", columns: %w[Name Age])
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    tbls = reader.tables
    assert_equal(1, tbls.size)
    assert_equal("A1:B2", tbls[0][:ref])
    assert_equal([{ name: "Name" }, { name: "Age" }], tbls[0][:columns])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips table with totals row and enhanced columns" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Item")
    writer.set_cell("B1", "Price")
    writer.set_cell("C1", "Tax")
    writer.add_table("A1:C5", columns: [
                       "Item",
                       { name: "Price", totals_row_function: "sum" },
                       { name: "Tax", calculated_column_formula: "[Price]*0.1" }
                     ], totals_row_count: 1, style: { name: "TableStyleLight1", show_row_stripes: false })
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    tbls = reader.tables
    assert_equal(1, tbls.size)
    assert_equal(1, tbls[0][:totals_row_count])
    assert_equal("sum", tbls[0][:columns][1][:totals_row_function])
    assert_equal("[Price]*0.1", tbls[0][:columns][2][:calculated_column_formula])
    assert_equal("TableStyleLight1", tbls[0][:style][:name])
    assert_equal(false, tbls[0][:style][:show_row_stripes])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips tableColumn extended attributes (totalsRowLabel, dxfIds, dataCellStyle)" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Item")
    writer.set_cell("B1", "Price")
    writer.add_table("A1:B5", columns: [
                       { name: "Item", totals_row_label: "Total", header_row_dxf_id: 1 },
                       { name: "Price", totals_row_function: "sum", data_dxf_id: 2,
                         totals_row_dxf_id: 3, data_cell_style: "Currency" }
                     ], totals_row_count: 1)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    tbls = reader.tables
    cols = tbls[0][:columns]
    assert_equal("Total", cols[0][:totals_row_label])
    assert_equal(1, cols[0][:header_row_dxf_id])
    assert_equal(2, cols[1][:data_dxf_id])
    assert_equal(3, cols[1][:totals_row_dxf_id])
    assert_equal("Currency", cols[1][:data_cell_style])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips totalsRowFormula in table column" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Item")
    writer.set_cell("B1", "Price")
    writer.add_table("A1:B5", columns: [
                       "Item",
                       { name: "Price", totals_row_function: "custom",
                         totals_row_formula: "SUBTOTAL(109,[Price])" }
                     ], totals_row_count: 1)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    tbls = reader.tables
    cols = tbls[0][:columns]
    assert_equal("custom", cols[1][:totals_row_function])
    assert_equal("SUBTOTAL(109,[Price])", cols[1][:totals_row_formula])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips chart with multiple series and axis titles" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Cat")
    writer.set_cell("B1", 10)
    writer.set_cell("C1", 20)
    writer.add_chart(type: :bar, title: "Multi",
                     series: [
                       { cat_ref: "Sheet1!$A$1:$A$3", val_ref: "Sheet1!$B$1:$B$3" },
                       { cat_ref: "Sheet1!$A$1:$A$3", val_ref: "Sheet1!$C$1:$C$3" }
                     ],
                     legend: { position: "b" },
                     data_labels: { show_val: true },
                     cat_axis_title: "Category",
                     val_axis_title: "Value")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    charts = reader.charts
    assert_equal(1, charts.size)
    assert_equal("barChart", charts[0][:chart_type])
    assert_equal("Multi", charts[0][:title])
    assert_equal(2, charts[0][:series].size)
    assert_equal("b", charts[0][:legend][:position])
    assert_equal(true, charts[0][:data_labels][:show_val])
    assert_equal("Category", charts[0][:cat_axis_title])
    assert_equal("Value", charts[0][:val_axis_title])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips shapes with preset geometry and text" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.add_shape(preset: "ellipse", text: "Hello", name: "Oval 1",
                     from_col: 1, from_row: 2, to_col: 4, to_row: 6)
    writer.add_shape(preset: "roundRect", name: "RR 1",
                     from_col: 5, from_row: 0, to_col: 8, to_row: 3)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    shapes = reader.shapes
    assert_equal(2, shapes.size)
    assert_equal("ellipse", shapes[0][:preset])
    assert_equal("Hello", shapes[0][:text])
    assert_equal("Oval 1", shapes[0][:name])
    assert_equal(1, shapes[0][:from_col])
    assert_equal(2, shapes[0][:from_row])
    assert_equal("roundRect", shapes[1][:preset])
    assert_equal("RR 1", shapes[1][:name])
    assert_nil(shapes[1][:text])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips pivot table with col_fields and items" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Category")
    writer.set_cell("B1", "Region")
    writer.set_cell("C1", "Amount")
    writer.add_pivot_table("Sheet1!A1:C4",
                           row_fields: [0],
                           col_fields: [1],
                           data_fields: [{ fld: 2, name: "Sum of Amount", subtotal: "sum" }],
                           field_names: %w[Category Region Amount],
                           items: { 0 => %w[A B C], 1 => %w[East West] })
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    pts = reader.pivot_tables
    assert_equal(1, pts.size)
    assert_equal([0], pts[0][:row_fields])
    assert_equal([1], pts[0][:col_fields])
    assert_equal("axisRow", pts[0][:fields][0][:axis])
    assert_equal("axisCol", pts[0][:fields][1][:axis])
    assert_equal(true, pts[0][:fields][2][:data_field])
    # Items parsed from pivotField
    assert_equal(4, pts[0][:fields][0][:items].size) # 3 data + 1 default
    assert_equal(0, pts[0][:fields][0][:items][0][:x])
    assert_equal("default", pts[0][:fields][0][:items].last[:t])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips pivot table extended attributes (dataCaption, dataOnRows, grandTotals, compact, outline, showHeaders)" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Category")
    writer.set_cell("B1", "Region")
    writer.set_cell("C1", "Amount")
    writer.add_pivot_table("Sheet1!A1:C4",
                           row_fields: [0],
                           data_fields: [{ fld: 2, name: "Sum of Amount", subtotal: "sum" }],
                           field_names: %w[Category Region Amount],
                           data_caption: "Custom Caption",
                           data_on_rows: true,
                           row_grand_totals: false,
                           col_grand_totals: false,
                           compact: false,
                           outline: false,
                           show_headers: false)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    pts = reader.pivot_tables
    assert_equal(1, pts.size)
    assert_equal("Custom Caption", pts[0][:data_caption])
    assert_equal(true, pts[0][:data_on_rows])
    assert_equal(false, pts[0][:row_grand_totals])
    assert_equal(false, pts[0][:col_grand_totals])
    assert_equal(false, pts[0][:compact])
    assert_equal(false, pts[0][:outline])
    assert_equal(false, pts[0][:show_headers])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "reader parses pivotField showAll, compact, outline, subtotalTop, numFmtId, sortType" do
    require "xlsxrb/reader"
    pivot_xml = <<~XML
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                            name="PivotTable1" cacheId="0" dataCaption="Values">
        <pivotFields count="2">
          <pivotField axis="axisRow" showAll="0" compact="0" outline="0" subtotalTop="0" numFmtId="164" sortType="ascending"/>
          <pivotField dataField="1" showAll="1"/>
        </pivotFields>
      </pivotTableDefinition>
    XML
    listener = Xlsxrb::Reader::PivotTableListener.new
    parser = REXML::Parsers::SAX2Parser.new(pivot_xml)
    parser.listen(listener)
    parser.parse
    fields = listener.pivot_table[:fields]
    assert_equal(2, fields.size)
    f0 = fields[0]
    assert_equal(false, f0[:show_all])
    assert_equal(false, f0[:compact])
    assert_equal(false, f0[:outline])
    assert_equal(false, f0[:subtotal_top])
    assert_equal(164, f0[:num_fmt_id])
    assert_equal("ascending", f0[:sort_type])
    f1 = fields[1]
    assert_equal(true, f1[:show_all])
  end

  test "round-trips pivotTableDefinition grandTotalCaption, errorCaption, missingCaption, tag, version attrs" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Category")
    writer.set_cell("B1", "Region")
    writer.set_cell("C1", "Amount")
    writer.add_pivot_table("Sheet1!A1:C4",
                           row_fields: [0],
                           data_fields: [{ fld: 2, name: "Sum", subtotal: "sum" }],
                           field_names: %w[Category Region Amount],
                           grand_total_caption: "Grand Total",
                           error_caption: "#N/A",
                           show_error: true,
                           missing_caption: "(blank)",
                           show_missing: false,
                           tag: "custom-tag",
                           indent: 2,
                           published: true,
                           created_version: 6,
                           updated_version: 8,
                           min_refreshable_version: 3)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    pts = reader.pivot_tables
    assert_equal(1, pts.size)
    assert_equal("Grand Total", pts[0][:grand_total_caption])
    assert_equal("#N/A", pts[0][:error_caption])
    assert_equal(true, pts[0][:show_error])
    assert_equal("(blank)", pts[0][:missing_caption])
    assert_equal(false, pts[0][:show_missing])
    assert_equal("custom-tag", pts[0][:tag])
    assert_equal(2, pts[0][:indent])
    assert_equal(true, pts[0][:published])
    assert_equal(6, pts[0][:created_version])
    assert_equal(8, pts[0][:updated_version])
    assert_equal(3, pts[0][:min_refreshable_version])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips pivotTableDefinition applyXxxFormats attributes" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Cat")
    writer.set_cell("B1", 10)
    writer.add_pivot_table("Sheet1!A1:B2",
                           row_fields: [0],
                           data_fields: [{ fld: 1, name: "Sum", subtotal: "sum" }],
                           apply_number_formats: true,
                           apply_border_formats: true,
                           apply_font_formats: false,
                           apply_pattern_formats: false,
                           apply_alignment_formats: false,
                           apply_width_height_formats: false)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    pts = reader.pivot_tables
    assert_equal(1, pts.size)
    assert_equal(true, pts[0][:apply_number_formats])
    assert_equal(true, pts[0][:apply_border_formats])
    assert_equal(false, pts[0][:apply_font_formats])
    assert_equal(false, pts[0][:apply_pattern_formats])
    assert_equal(false, pts[0][:apply_alignment_formats])
    assert_equal(false, pts[0][:apply_width_height_formats])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips pivotTableStyleInfo" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Category")
    writer.set_cell("B1", "Region")
    writer.set_cell("C1", "Amount")
    writer.add_pivot_table("Sheet1!A1:C4",
                           row_fields: [0],
                           data_fields: [{ fld: 2, name: "Sum", subtotal: "sum" }],
                           field_names: %w[Category Region Amount],
                           pivot_table_style: { name: "PivotStyleLight16",
                                                show_row_headers: true,
                                                show_col_headers: false,
                                                show_row_stripes: true,
                                                show_col_stripes: false,
                                                show_last_column: true })
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    pts = reader.pivot_tables
    psi = pts[0][:pivot_table_style]
    assert_not_nil(psi)
    assert_equal("PivotStyleLight16", psi[:name])
    assert_equal(true, psi[:show_row_headers])
    assert_equal(false, psi[:show_col_headers])
    assert_equal(true, psi[:show_row_stripes])
    assert_equal(false, psi[:show_col_stripes])
    assert_equal(true, psi[:show_last_column])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips dataField showDataAs, baseField, baseItem, numFmtId" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Category")
    writer.set_cell("B1", "Region")
    writer.set_cell("C1", "Amount")
    writer.add_pivot_table("Sheet1!A1:C4",
                           row_fields: [0],
                           data_fields: [{ fld: 2, name: "% of Total", subtotal: "sum",
                                           show_data_as: "percentOfTotal", base_field: 0, base_item: 0, num_fmt_id: 10 }],
                           field_names: %w[Category Region Amount])
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    pts = reader.pivot_tables
    assert_equal(1, pts.size)
    df = pts[0][:data_fields][0]
    assert_equal("percentOfTotal", df[:show_data_as])
    assert_equal(0, df[:base_field])
    assert_equal(0, df[:base_item])
    assert_equal(10, df[:num_fmt_id])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips external links" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.add_external_link(target: "Book2.xlsx", sheet_names: %w[Data Summary])
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    els = reader.external_links
    assert_equal(1, els.size)
    assert_equal("Book2.xlsx", els[0][:target])
    assert_equal(%w[Data Summary], els[0][:sheet_names])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "format_variant returns transitional for standard writer output" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    assert_equal(:transitional, reader.format_variant)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "format_variant returns strict for strict namespace workbook" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    # Create a minimal XLSX with strict namespace using raw entries only.
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.write(xlsx_path)

    # Read all entries, patch workbook.xml, then write using add_raw_entry only.
    reader = Xlsxrb::Reader.new(xlsx_path)
    entries = {}
    reader.entry_names.each do |name|
      content = reader.raw_entry(name)
      if name == "xl/workbook.xml"
        content = content.gsub(
          "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          "http://purl.oclc.org/ooxml/spreadsheetml/main/2006/main"
        )
      end
      entries[name] = content
    end

    # Write patched file using ZipGenerator directly.
    gen = Xlsxrb::ZipGenerator.new(xlsx_path)
    entries.each { |name, data| gen.add_entry(name, data) }
    gen.generate

    reader2 = Xlsxrb::Reader.new(xlsx_path)
    assert_equal(:strict, reader2.format_variant)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips shared string table mode" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.use_shared_strings!
    writer.set_cell("A1", "hello")
    writer.set_cell("B1", "hello")
    writer.set_cell("C1", "world")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cells = reader.cells
    assert_equal("hello", cells["A1"])
    assert_equal("hello", cells["B1"])
    assert_equal("world", cells["C1"])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  # --- Phase 2: Reader unit tests ---

  test "round-trips images through writer and reader" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    png_bytes = [
      0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
      0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
      0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
      0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
      0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
      0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
      0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC,
      0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
      0x44, 0xAE, 0x42, 0x60, 0x82
    ].pack("C*")

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "img test")
    writer.insert_image(png_bytes, ext: "png", from_col: 1, from_row: 2, to_col: 6, to_row: 12,
                                   name: "RoundTrip Pic", description: "Round-trip image description")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    imgs = reader.images
    assert_equal(1, imgs.size)
    assert_equal("RoundTrip Pic", imgs[0][:name])
    assert_equal("Round-trip image description", imgs[0][:description])
    assert_equal(1, imgs[0][:from_col])
    assert_equal(2, imgs[0][:from_row])
    assert_equal(6, imgs[0][:to_col])
    assert_equal(12, imgs[0][:to_row])
    assert_not_nil(imgs[0][:target])

    # Verify image data is accessible.
    img_data = reader.raw_entry("xl/media/image1.png")
    assert_equal(png_bytes, img_data)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips charts through writer and reader" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Category")
    writer.set_cell("B1", "Value")
    writer.add_chart(type: :bar, title: "My Chart")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    charts = reader.charts
    assert_equal(1, charts.size)
    assert_equal("barChart", charts[0][:chart_type])
    assert_equal("My Chart", charts[0][:title])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips image editAs attribute" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    png_bytes = [
      0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
      0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
      0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
      0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
      0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
      0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
      0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC,
      0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
      0x44, 0xAE, 0x42, 0x60, 0x82
    ].pack("C*")

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.insert_image(png_bytes, ext: "png", name: "Pic1", edit_as: "absolute")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    imgs = reader.images
    assert_equal(1, imgs.size)
    assert_equal("absolute", imgs[0][:edit_as])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips shape editAs attribute" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", text: "Hello", edit_as: "absolute")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    shapes = reader.shapes
    assert_equal(1, shapes.size)
    assert_equal("absolute", shapes[0][:edit_as])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips shape description attribute" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", text: "Hello", description: "A rectangle shape")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    shapes = reader.shapes
    assert_equal(1, shapes.size)
    assert_equal("A rectangle shape", shapes[0][:description])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips chart editAs attribute" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_chart(type: :bar, title: "Chart1", edit_as: "oneCell")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    charts = reader.charts
    assert_equal(1, charts.size)
    assert_equal("oneCell", charts[0][:edit_as])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips chart anchor positions" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_chart(type: :bar, title: "Positioned",
                     cat_ref: "Sheet1!$A$1:$A$1", val_ref: "Sheet1!$A$1:$A$1",
                     from_col: 2, from_row: 3, to_col: 8, to_row: 18,
                     from_col_off: 100, from_row_off: 200, to_col_off: 300, to_row_off: 400)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    charts = reader.charts
    assert_equal(1, charts.size)
    assert_equal(2, charts[0][:from_col])
    assert_equal(3, charts[0][:from_row])
    assert_equal(8, charts[0][:to_col])
    assert_equal(18, charts[0][:to_row])
    assert_equal(100, charts[0][:from_col_off])
    assert_equal(200, charts[0][:from_row_off])
    assert_equal(300, charts[0][:to_col_off])
    assert_equal(400, charts[0][:to_row_off])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips chart name and description" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_chart(type: :bar, title: "Sales",
                     cat_ref: "Sheet1!$A$1:$A$1", val_ref: "Sheet1!$A$1:$A$1",
                     name: "Chart 1", description: "A sales chart")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    charts = reader.charts
    assert_equal(1, charts.size)
    assert_equal("Chart 1", charts[0][:name])
    assert_equal("A sales chart", charts[0][:description])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips image clientData fLocksWithSheet and fPrintsWithSheet" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    png_bytes = [
      0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
      0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
      0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
      0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
      0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
      0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
      0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC,
      0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
      0x44, 0xAE, 0x42, 0x60, 0x82
    ].pack("C*")

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.insert_image(png_bytes, ext: "png", name: "Pic1", locks_with_sheet: false, prints_with_sheet: false)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    imgs = reader.images
    assert_equal(1, imgs.size)
    assert_equal(false, imgs[0][:locks_with_sheet])
    assert_equal(false, imgs[0][:prints_with_sheet])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips image anchor offsets" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    png_bytes = [
      0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
      0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
      0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
      0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
      0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
      0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
      0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC,
      0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
      0x44, 0xAE, 0x42, 0x60, 0x82
    ].pack("C*")

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.insert_image(png_bytes, ext: "png", name: "Pic1",
                                   from_col: 1, from_row: 2, to_col: 5, to_row: 8,
                                   from_col_off: 100_000, from_row_off: 200_000,
                                   to_col_off: 300_000, to_row_off: 400_000)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    imgs = reader.images
    assert_equal(1, imgs.size)
    assert_equal(100_000, imgs[0][:from_colOff])
    assert_equal(200_000, imgs[0][:from_rowOff])
    assert_equal(300_000, imgs[0][:to_colOff])
    assert_equal(400_000, imgs[0][:to_rowOff])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips comments through writer and reader" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.add_comment("A1", "Test note", author: "Me")
    writer.add_comment("B2", "Second note", author: "You")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    comments = reader.comments
    assert_equal(2, comments.size)
    assert_equal("A1", comments[0][:ref])
    assert_equal("Me", comments[0][:author])
    assert_equal("Test note", comments[0][:text])
    assert_equal("B2", comments[1][:ref])
    assert_equal("You", comments[1][:author])
    assert_equal("Second note", comments[1][:text])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips rich text comments" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "data")
    rt = Xlsxrb::RichText.new(runs: [
                                { text: "Bold", font: { bold: true, sz: 9, name: "Calibri" } },
                                { text: " normal" }
                              ])
    writer.add_comment("A1", rt, author: "Tester")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    comments = reader.comments
    assert_equal(1, comments.size)
    c = comments[0]
    assert_equal("A1", c[:ref])
    assert_equal("Tester", c[:author])
    assert_instance_of(Xlsxrb::RichText, c[:text])
    assert_equal("Bold normal", c[:text].to_s)
    runs = c[:text].runs
    assert_equal(2, runs.size)
    assert_equal("Bold", runs[0][:text])
    assert_equal(true, runs[0][:font][:bold])
    assert_equal(9.0, runs[0][:font][:sz])
    assert_equal("Calibri", runs[0][:font][:name])
    assert_equal(" normal", runs[1][:text])
    assert_nil(runs[1][:font])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "entry_names lists all ZIP entries" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    names = reader.entry_names
    assert(names.include?("[Content_Types].xml"))
    assert(names.include?("xl/workbook.xml"))
    assert(names.include?("xl/worksheets/sheet1.xml"))
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "macros? returns false for normal xlsx" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    assert_false(reader.macros?)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "copy_entries_from preserves all content" do
    source_tempfile = Tempfile.new(["xlsxrb-src", ".xlsx"])
    source_path = source_tempfile.path
    source_tempfile.close

    output_tempfile = Tempfile.new(["xlsxrb-out", ".xlsx"])
    output_path = output_tempfile.path
    output_tempfile.close

    # Write source with images and comments.
    png_bytes = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
                 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
                 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
                 0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
                 0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
                 0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC,
                 0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
                 0x44, 0xAE, 0x42, 0x60, 0x82].pack("C*")
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "original")
    writer.insert_image(png_bytes, ext: "png")
    writer.add_comment("A1", "Original comment", author: "Author")
    writer.write(source_path)

    # Copy entries through another Writer.
    writer2 = Xlsxrb::Writer.new
    writer2.copy_entries_from(source_path)
    writer2.write(output_path)

    # Verify output has images and comments.
    reader = Xlsxrb::Reader.new(output_path)
    imgs = reader.images
    assert_equal(1, imgs.size)
    comments = reader.comments
    assert_equal(1, comments.size)
    assert_equal("Original comment", comments[0][:text])
  ensure
    File.delete(source_path) if source_path && File.exist?(source_path)
    File.delete(output_path) if output_path && File.exist?(output_path)
  end

  test "round-trips sheet protection through writer and reader" do
    writer = Xlsxrb::Writer.new
    writer.set_sheet_protection(password: "CF1A", objects: true, scenarios: true)
    writer.set_cell("A1", "Protected")

    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    prot = reader.sheet_protection
    assert_not_nil(prot)
    assert_equal(true, prot[:sheet])
    assert_equal("CF1A", prot[:password])
    assert_equal(true, prot[:objects])
    assert_equal(true, prot[:scenarios])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips workbook protection through writer and reader" do
    writer = Xlsxrb::Writer.new
    writer.set_workbook_protection(lock_structure: true)
    writer.set_cell("A1", "test")

    xlsx_path = Tempfile.new(["xlsxrb-test", ".xlsx"]).path
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    prot = reader.workbook_protection
    assert_not_nil(prot)
    assert_equal(true, prot[:lock_structure])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips workbook protection lockRevision and revision algorithm attrs" do
    writer = Xlsxrb::Writer.new
    writer.set_workbook_protection(
      lock_structure: true,
      lock_revision: true,
      revisions_algorithm_name: "SHA-512",
      revisions_hash_value: "abc123",
      revisions_salt_value: "salt456",
      revisions_spin_count: 100_000
    )
    writer.set_cell("A1", "test")

    xlsx_path = Tempfile.new(["xlsxrb-test", ".xlsx"]).path
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    prot = reader.workbook_protection
    assert_not_nil(prot)
    assert_equal(true, prot[:lock_structure])
    assert_equal(true, prot[:lock_revision])
    assert_equal("SHA-512", prot[:revisions_algorithm_name])
    assert_equal("abc123", prot[:revisions_hash_value])
    assert_equal("salt456", prot[:revisions_salt_value])
    assert_equal(100_000, prot[:revisions_spin_count])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips rich text inline through writer and reader" do
    writer = Xlsxrb::Writer.new
    rt = Xlsxrb::RichText.new(runs: [
                                { text: "Bold", font: { bold: true, sz: 14.0, color: "FFFF0000" } },
                                { text: " Normal" }
                              ])
    writer.set_cell("A1", rt)

    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cells = reader.cells
    result = cells["A1"]
    assert_instance_of(Xlsxrb::RichText, result)
    assert_equal(2, result.runs.size)
    assert_equal("Bold", result.runs[0][:text])
    assert_equal(true, result.runs[0][:font][:bold])
    assert_equal(14.0, result.runs[0][:font][:sz])
    assert_equal(" Normal", result.runs[1][:text])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips rich text through shared strings" do
    writer = Xlsxrb::Writer.new
    writer.use_shared_strings!
    rt = Xlsxrb::RichText.new(runs: [
                                { text: "Italic", font: { italic: true } },
                                { text: " plain" }
                              ])
    writer.set_cell("A1", rt)

    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cells = reader.cells
    result = cells["A1"]
    assert_instance_of(Xlsxrb::RichText, result)
    assert_equal(2, result.runs.size)
    assert_equal("Italic", result.runs[0][:text])
    assert_equal(true, result.runs[0][:font][:italic])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips shared formulas through writer and reader" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.set_cell("A2", 20)
    writer.set_cell("B1", Xlsxrb::Formula.new(expression: "A1*2", type: :shared, ref: "B1:B2", shared_index: 0, cached_value: "20"))
    writer.set_cell("B2", Xlsxrb::Formula.new(expression: "", type: :shared, shared_index: 0, cached_value: "40"))

    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cells = reader.cells
    b1 = cells["B1"]
    assert_instance_of(Xlsxrb::Formula, b1)
    assert_equal(:shared, b1.type)
    assert_equal("B1:B2", b1.ref)
    assert_equal(0, b1.shared_index)
    assert_equal("A1*2", b1.expression)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips array formulas through writer and reader" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.set_cell("A2", 2)
    writer.set_cell("B1", 3)
    writer.set_cell("B2", 4)
    writer.set_cell("C1", Xlsxrb::Formula.new(expression: "SUM(A1:A2*B1:B2)", type: :array, ref: "C1", cached_value: "11"))

    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cells = reader.cells
    c1 = cells["C1"]
    assert_instance_of(Xlsxrb::Formula, c1)
    assert_equal(:array, c1.type)
    assert_equal("C1", c1.ref)
    assert_equal("SUM(A1:A2*B1:B2)", c1.expression)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips formula calculate_always through writer and reader" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", Xlsxrb::Formula.new(expression: "NOW()", cached_value: "45000", calculate_always: true))
    writer.set_cell("B1", Xlsxrb::Formula.new(expression: "A1+1", cached_value: "2"))

    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cells = reader.cells
    a1 = cells["A1"]
    assert_instance_of(Xlsxrb::Formula, a1)
    assert_equal(true, a1.calculate_always)
    assert_equal("NOW()", a1.expression)

    b1 = cells["B1"]
    assert_instance_of(Xlsxrb::Formula, b1)
    assert_nil(b1.calculate_always)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips calcChain through writer and reader" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.set_cell("B1", Xlsxrb::Formula.new(expression: "A1*2", cached_value: "20"))

    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    chain = reader.calc_chain
    assert_equal(1, chain.size)
    assert_equal("B1", chain[0][:ref])
    assert_equal(1, chain[0][:sheet_id])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips cell alignment attributes" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    style_id = writer.add_cell_style(
      alignment: { horizontal: "center", vertical: "top", wrap_text: true,
                   text_rotation: 45, indent: 2, shrink_to_fit: true }
    )
    writer.set_cell("A1", "aligned")
    writer.set_cell_style("A1", style_id)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cs = reader.cell_styles
    assert(cs.key?("A1"), "A1 should have a style")
    alignment = cs["A1"][:alignment]
    assert_not_nil(alignment, "alignment should be present")
    assert_equal("center", alignment[:horizontal])
    assert_equal("top", alignment[:vertical])
    assert_equal(true, alignment[:wrap_text])
    assert_equal(45, alignment[:text_rotation])
    assert_equal(2, alignment[:indent])
    assert_equal(true, alignment[:shrink_to_fit])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips extended font attributes" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    fid = writer.add_font(
      bold: true, italic: true, strike: true, sz: 12, name: "Calibri",
      color: "FF0000FF", underline: "double", vert_align: "superscript",
      scheme: "minor", family: 2
    )
    style_id = writer.add_cell_style(font_id: fid)
    writer.set_cell("A1", "extended")
    writer.set_cell_style("A1", style_id)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cs = reader.cell_styles
    assert(cs.key?("A1"), "A1 should have a style")
    font = cs["A1"][:font]
    assert_not_nil(font, "font should be present")
    assert_equal(true, font[:bold])
    assert_equal(true, font[:italic])
    assert_equal(true, font[:strike])
    assert_equal("double", font[:underline])
    assert_equal("superscript", font[:vert_align])
    assert_equal("minor", font[:scheme])
    assert_equal(2, font[:family])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips gradient fill" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    fill_id = writer.add_fill(
      gradient: { type: "linear", degree: 90,
                  stops: [{ position: 0, color: "FFFF0000" }, { position: 1, color: "FF0000FF" }] }
    )
    style_id = writer.add_cell_style(fill_id: fill_id)
    writer.set_cell("A1", "gradient")
    writer.set_cell_style("A1", style_id)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cs = reader.cell_styles
    assert(cs.key?("A1"), "A1 should have a style")
    fill = cs["A1"][:fill]
    assert_not_nil(fill, "fill should be present")
    gradient = fill[:gradient]
    assert_not_nil(gradient, "gradient should be present")
    assert_equal("linear", gradient[:type])
    assert_equal(90.0, gradient[:degree])
    assert_equal(2, gradient[:stops].size)
    assert_equal(0.0, gradient[:stops][0][:position])
    assert_equal("FFFF0000", gradient[:stops][0][:color])
    assert_equal(1.0, gradient[:stops][1][:position])
    assert_equal("FF0000FF", gradient[:stops][1][:color])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips diagonal border" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    brd_id = writer.add_border(
      left: { style: "thin" }, diagonal: { style: "thin", color: "FFFF0000" },
      diagonal_up: true, diagonal_down: true
    )
    style_id = writer.add_cell_style(border_id: brd_id)
    writer.set_cell("A1", "diag")
    writer.set_cell_style("A1", style_id)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cs = reader.cell_styles
    assert(cs.key?("A1"), "A1 should have a style")
    border = cs["A1"][:border]
    assert_not_nil(border, "border should be present")
    assert_equal(true, border[:diagonal_up])
    assert_equal(true, border[:diagonal_down])
    assert_equal("thin", border[:diagonal][:style])
    assert_equal("FFFF0000", border[:diagonal][:color])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips cell protection attributes" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    style_id = writer.add_cell_style(protection: { locked: false, hidden: true })
    writer.set_cell("A1", "protected")
    writer.set_cell_style("A1", style_id)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cs = reader.cell_styles
    assert(cs.key?("A1"), "A1 should have a style")
    prot = cs["A1"][:protection]
    assert_not_nil(prot, "protection should be present")
    assert_equal(false, prot[:locked])
    assert_equal(true, prot[:hidden])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips font theme and indexed colors" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    fid1 = writer.add_font(sz: 11, name: "Calibri", theme: 1, tint: -0.25)
    fid2 = writer.add_font(sz: 11, name: "Calibri", indexed: 10)
    s1 = writer.add_cell_style(font_id: fid1)
    s2 = writer.add_cell_style(font_id: fid2)
    writer.set_cell("A1", "theme")
    writer.set_cell_style("A1", s1)
    writer.set_cell("A2", "indexed")
    writer.set_cell_style("A2", s2)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cs = reader.cell_styles
    assert(cs.key?("A1"), "A1 should have a style")
    assert_equal(1, cs["A1"][:font][:theme])
    assert_in_delta(-0.25, cs["A1"][:font][:tint], 0.001)
    assert(cs.key?("A2"), "A2 should have a style")
    assert_equal(10, cs["A2"][:font][:indexed])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips fill theme colors" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    fill_id = writer.add_fill(pattern: "solid", fg_color_theme: 4, fg_color_tint: 0.6)
    style_id = writer.add_cell_style(fill_id: fill_id)
    writer.set_cell("A1", "theme fill")
    writer.set_cell_style("A1", style_id)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cs = reader.cell_styles
    assert(cs.key?("A1"), "A1 should have a style")
    fill = cs["A1"][:fill]
    assert_not_nil(fill, "fill should be present")
    assert_equal(4, fill[:fg_color_theme])
    assert_in_delta(0.6, fill[:fg_color_tint], 0.001)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips fill auto colors" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    fill_id = writer.add_fill(pattern: "solid", fg_color_auto: true, bg_color_auto: true)
    style_id = writer.add_cell_style(fill_id: fill_id)
    writer.set_cell("A1", "auto fill")
    writer.set_cell_style("A1", style_id)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cs = reader.cell_styles
    assert(cs.key?("A1"), "A1 should have a style")
    fill = cs["A1"][:fill]
    assert_not_nil(fill, "fill should be present")
    assert_equal(true, fill[:fg_color_auto])
    assert_equal(true, fill[:bg_color_auto])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips border theme color and tabColor theme" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    brd_id = writer.add_border(left: { style: "thin", theme: 1, tint: -0.25 })
    style_id = writer.add_cell_style(border_id: brd_id)
    writer.set_cell("A1", "themed border")
    writer.set_cell_style("A1", style_id)
    writer.set_sheet_property(:tab_color_theme, 3)
    writer.set_sheet_property(:tab_color_tint, -0.5)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cs = reader.cell_styles
    border = cs["A1"][:border]
    assert_equal(1, border[:left][:theme])
    assert_in_delta(-0.25, border[:left][:tint], 0.001)

    props = reader.sheet_properties
    assert_equal(3, props[:tab_color_theme])
    assert_in_delta(-0.5, props[:tab_color_tint], 0.001)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips expanded conditional formatting rule types" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.add_dxf(font: { bold: true, color: "FFFF0000" })
    writer.add_conditional_format("A1:A10", type: :above_average, priority: 1,
                                            above_average: false, equal_average: true, format_id: 0)
    writer.add_conditional_format("B1:B10", type: :top10, priority: 2,
                                            rank: 5, percent: true, bottom: true, format_id: 0)
    writer.add_conditional_format("C1:C10", type: :duplicate_values, priority: 3, format_id: 0)
    writer.add_conditional_format("D1:D10", type: :contains_text, priority: 4, operator: "containsText",
                                            text: "hello", formula: 'NOT(ISERROR(SEARCH("hello",D1)))',
                                            format_id: 0)
    writer.add_conditional_format("E1:E10", type: :begins_with, priority: 5, operator: "beginsWith",
                                            text: "foo", formula: 'LEFT(E1,3)="foo"',
                                            format_id: 0)
    writer.add_conditional_format("F1:F10", type: :ends_with, priority: 6, operator: "endsWith",
                                            text: "bar", formula: 'RIGHT(F1,3)="bar"',
                                            format_id: 0)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cfs = reader.conditional_formats
    assert_equal(6, cfs.size)

    # aboveAverage (below average + equal)
    assert_equal("aboveAverage", cfs[0][:type])
    assert_equal(false, cfs[0][:above_average])
    assert_equal(true, cfs[0][:equal_average])

    # top10 (bottom 5 percent)
    assert_equal("top10", cfs[1][:type])
    assert_equal(5, cfs[1][:rank])
    assert_equal(true, cfs[1][:percent])
    assert_equal(true, cfs[1][:bottom])

    # duplicateValues
    assert_equal("duplicateValues", cfs[2][:type])

    # containsText
    assert_equal("containsText", cfs[3][:type])
    assert_equal("hello", cfs[3][:text])
    assert_equal(1, cfs[3][:formulas].size)

    # beginsWith
    assert_equal("beginsWith", cfs[4][:type])
    assert_equal("foo", cfs[4][:text])

    # endsWith
    assert_equal("endsWith", cfs[5][:type])
    assert_equal("bar", cfs[5][:text])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips rich text with extended font attributes" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    rt = Xlsxrb::RichText.new(runs: [
                                { text: "Strike", font: { strike: true, name: "Arial", sz: 11 } },
                                { text: "DblUnder", font: { underline: "double", name: "Arial", sz: 11 } },
                                { text: "Super", font: { vert_align: "superscript", name: "Arial", sz: 11 } },
                                { text: "Theme", font: { theme: 1, tint: 0.5, name: "Calibri", sz: 11, family: 2, scheme: "minor" } },
                                { text: "Indexed", font: { indexed: 10, name: "Calibri", sz: 11 } }
                              ])
    writer.set_cell("A1", rt)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    val = reader.cells["A1"]
    assert_instance_of(Xlsxrb::RichText, val)
    runs = val.runs

    assert_equal(5, runs.size)

    # Run 0: strike
    assert_equal("Strike", runs[0][:text])
    assert_equal(true, runs[0][:font][:strike])

    # Run 1: double underline
    assert_equal("DblUnder", runs[1][:text])
    assert_equal("double", runs[1][:font][:underline])

    # Run 2: superscript
    assert_equal("Super", runs[2][:text])
    assert_equal("superscript", runs[2][:font][:vert_align])

    # Run 3: theme color, family, scheme
    assert_equal("Theme", runs[3][:text])
    assert_equal(1, runs[3][:font][:theme])
    assert_in_delta(0.5, runs[3][:font][:tint], 0.001)
    assert_equal(2, runs[3][:font][:family])
    assert_equal("minor", runs[3][:font][:scheme])

    # Run 4: indexed color
    assert_equal("Indexed", runs[4][:text])
    assert_equal(10, runs[4][:font][:indexed])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips CF colorScale and dataBar with theme/indexed colors" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 50)
    writer.add_conditional_format("A1:A10", type: :color_scale, priority: 1,
                                            color_scale: {
                                              cfvo: [{ type: "min" }, { type: "max" }],
                                              colors: [{ theme: 4, tint: -0.25 }, { theme: 9 }]
                                            })
    writer.add_conditional_format("B1:B10", type: :data_bar, priority: 2,
                                            data_bar: {
                                              cfvo: [{ type: "min" }, { type: "max" }],
                                              color: { indexed: 10 }
                                            })
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cfs = reader.conditional_formats
    assert_equal(2, cfs.size)

    cs_colors = cfs[0][:color_scale][:colors]
    assert_equal(2, cs_colors.size)
    assert_equal(4, cs_colors[0][:theme])
    assert_in_delta(-0.25, cs_colors[0][:tint], 0.001)
    assert_equal(9, cs_colors[1][:theme])

    db_color = cfs[1][:data_bar][:color]
    assert_equal(10, db_color[:indexed])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips gradient fill with theme/indexed stop colors" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    fill_id = writer.add_fill(gradient: {
                                degree: 90,
                                stops: [{ position: 0, theme: 4, tint: -0.5 }, { position: 1, indexed: 12 }]
                              })
    style_id = writer.add_cell_style(fill_id: fill_id)
    writer.set_cell("A1", "themed gradient")
    writer.set_cell_style("A1", style_id)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cs = reader.cell_styles
    gradient = cs["A1"][:fill][:gradient]
    assert_not_nil(gradient, "gradient should be present")
    stops = gradient[:stops]
    assert_equal(2, stops.size)
    assert_equal(4, stops[0][:theme])
    assert_in_delta(-0.5, stops[0][:tint], 0.001)
    assert_equal(12, stops[1][:indexed])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips complete set of CF rule types" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.add_dxf(font: { bold: true, color: "FFFF0000" })
    writer.add_conditional_format("A1:A10", type: :expression, priority: 1,
                                            formula: "MOD(ROW(),2)=0", format_id: 0)
    writer.add_conditional_format("B1:B10", type: :unique_values, priority: 2, format_id: 0)
    writer.add_conditional_format("C1:C10", type: :not_contains_text, priority: 3, operator: "notContains",
                                            text: "bad", formula: 'ISERROR(SEARCH("bad",C1))',
                                            format_id: 0)
    writer.add_conditional_format("D1:D10", type: :contains_blanks, priority: 4,
                                            formula: "LEN(TRIM(D1))=0", format_id: 0)
    writer.add_conditional_format("E1:E10", type: :not_contains_blanks, priority: 5,
                                            formula: "LEN(TRIM(E1))>0", format_id: 0)
    writer.add_conditional_format("F1:F10", type: :time_period, priority: 6,
                                            time_period: "lastWeek",
                                            formula: "AND(TODAY()-7<=F1,F1<=TODAY())",
                                            format_id: 0)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cfs = reader.conditional_formats
    assert_equal(6, cfs.size)

    # expression
    assert_equal("expression", cfs[0][:type])
    assert_equal(["MOD(ROW(),2)=0"], cfs[0][:formulas])

    # uniqueValues
    assert_equal("uniqueValues", cfs[1][:type])

    # notContainsText
    assert_equal("notContainsText", cfs[2][:type])
    assert_equal("bad", cfs[2][:text])

    # containsBlanks
    assert_equal("containsBlanks", cfs[3][:type])

    # notContainsBlanks
    assert_equal("notContainsBlanks", cfs[4][:type])

    # timePeriod
    assert_equal("timePeriod", cfs[5][:type])
    assert_equal("lastWeek", cfs[5][:time_period])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips DXF with alignment, protection, and numFmt" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.add_dxf(
      font: { bold: true, color: "FFFF0000" },
      num_fmt: { num_fmt_id: 164, format_code: "#,##0.00" },
      alignment: { horizontal: "center", wrap_text: true },
      protection: { locked: false, hidden: true }
    )
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    dxfs = reader.dxfs
    assert_equal(1, dxfs.size)
    dxf = dxfs[0]

    assert_equal(true, dxf[:font][:bold])
    assert_equal(164, dxf[:num_fmt][:num_fmt_id])
    assert_equal("#,##0.00", dxf[:num_fmt][:format_code])
    assert_equal("center", dxf[:alignment][:horizontal])
    assert_equal(true, dxf[:alignment][:wrap_text])
    assert_equal(false, dxf[:protection][:locked])
    assert_equal(true, dxf[:protection][:hidden])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips error cell values through writer and reader" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", Xlsxrb::CellError.new(code: "#N/A"))
    writer.set_cell("B1", Xlsxrb::CellError.new(code: "#DIV/0!"))
    writer.set_cell("C1", Xlsxrb::CellError.new(code: "#VALUE!"))
    writer.set_cell("D1", Xlsxrb::CellError.new(code: "#REF!"))
    writer.set_cell("E1", Xlsxrb::CellError.new(code: "#NAME?"))
    writer.set_cell("F1", Xlsxrb::CellError.new(code: "#NUM!"))
    writer.set_cell("G1", Xlsxrb::CellError.new(code: "#NULL!"))
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cells = reader.cells

    %w[A1 B1 C1 D1 E1 F1 G1].each do |ref|
      assert_instance_of(Xlsxrb::CellError, cells[ref], "Expected CellError for #{ref}")
    end

    assert_equal("#N/A", cells["A1"].code)
    assert_equal("#DIV/0!", cells["B1"].code)
    assert_equal("#VALUE!", cells["C1"].code)
    assert_equal("#REF!", cells["D1"].code)
    assert_equal("#NAME?", cells["E1"].code)
    assert_equal("#NUM!", cells["F1"].code)
    assert_equal("#NULL!", cells["G1"].code)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips extended core properties through writer and reader" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_core_property(:title, "My Title")
    writer.set_core_property(:subject, "My Subject")
    writer.set_core_property(:creator, "Alice")
    writer.set_core_property(:keywords, "ruby, xlsx")
    writer.set_core_property(:description, "A test document")
    writer.set_core_property(:last_modified_by, "Bob")
    writer.set_core_property(:revision, "3")
    writer.set_core_property(:category, "Reports")
    writer.set_core_property(:content_status, "Draft")
    writer.set_core_property(:language, "en-US")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    props = reader.core_properties

    assert_equal("My Title", props[:title])
    assert_equal("My Subject", props[:subject])
    assert_equal("Alice", props[:creator])
    assert_equal("ruby, xlsx", props[:keywords])
    assert_equal("A test document", props[:description])
    assert_equal("Bob", props[:last_modified_by])
    assert_equal("3", props[:revision])
    assert_equal("Reports", props[:category])
    assert_equal("Draft", props[:content_status])
    assert_equal("en-US", props[:language])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips split pane through writer and reader" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_split_pane(x_split: 2400, y_split: 1800, top_left_cell: "C4")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    pane = reader.freeze_pane

    assert_equal(:split, pane[:state])
    assert_equal(2400, pane[:x_split])
    assert_equal(1800, pane[:y_split])
    assert_equal("C4", pane[:top_left_cell])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips colorFilter and iconFilter through writer and reader" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "H1")
    writer.set_cell("B1", "H2")
    writer.set_auto_filter("A1:B10")
    writer.add_filter_column(0, { type: :color_filter, dxf_id: 0, cell_color: false })
    writer.add_filter_column(1, { type: :icon_filter, icon_set: "3Arrows", icon_id: 2 })
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    filters = reader.filter_columns

    cf = filters[0]
    assert_equal(:color_filter, cf[:type])
    assert_equal(0, cf[:dxf_id])
    assert_equal(false, cf[:cell_color])

    icf = filters[1]
    assert_equal(:icon_filter, icf[:type])
    assert_equal("3Arrows", icf[:icon_set])
    assert_equal(2, icf[:icon_id])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips Time values as fractional serial through writer and reader" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    t = Time.utc(2024, 3, 15, 14, 30, 0)
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", t)
    writer.set_cell("B1", Date.new(2024, 1, 1))
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cells = reader.cells

    # Time cell should be returned as Time
    assert_instance_of(Time, cells["A1"])
    assert_equal(2024, cells["A1"].year)
    assert_equal(3, cells["A1"].month)
    assert_equal(15, cells["A1"].day)
    assert_equal(14, cells["A1"].hour)
    assert_equal(30, cells["A1"].min)
    assert_equal(0, cells["A1"].sec)

    # Date cell should still be returned as Date
    assert_instance_of(Date, cells["B1"])
    assert_equal(Date.new(2024, 1, 1), cells["B1"])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips print area and print titles through writer and reader" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_print_area("A1:D20")
    writer.set_print_titles(rows: "1:3", cols: "A:B")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)

    pa = reader.print_area
    assert_equal("$A$1:$D$20", pa)

    pt = reader.print_titles
    assert_not_nil(pt)
    assert(pt.include?("$A:$B"), "Expected print titles to contain '$A:$B' but got '#{pt}'")
    assert(pt.include?("$1:$3"), "Expected print titles to contain '$1:$3' but got '#{pt}'")
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips hashed password sheet protection through writer and reader" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "protected")
    hp = Xlsxrb.hash_password("secret", spin_count: 500)
    writer.set_sheet_protection(**hp)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    sp = reader.sheet_protection

    assert_equal("SHA-512", sp[:algorithm_name])
    assert_equal(hp[:hash_value], sp[:hash_value])
    assert_equal(hp[:salt_value], sp[:salt_value])
    assert_equal(500, sp[:spin_count])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips first page header/footer through writer and reader" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_header_footer(:first_header, "&CFirst Page")
    writer.set_header_footer(:first_footer, "&CPage &P")
    writer.set_header_footer(:different_first, true)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    hf = reader.header_footer

    assert_equal(true, hf[:different_first])
    assert_equal("&CFirst Page", hf[:first_header])
    assert_equal("&CPage &P", hf[:first_footer])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips extended page setup attributes through writer and reader" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_page_setup(:page_order, "overThenDown")
    writer.set_page_setup(:black_and_white, true)
    writer.set_page_setup(:draft, true)
    writer.set_page_setup(:first_page_number, 3)
    writer.set_page_setup(:use_first_page_number, true)
    writer.set_page_setup(:horizontal_dpi, 600)
    writer.set_page_setup(:vertical_dpi, 600)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    ps = reader.page_setup

    assert_equal("overThenDown", ps[:page_order])
    assert_equal(true, ps[:black_and_white])
    assert_equal(true, ps[:draft])
    assert_equal(3, ps[:first_page_number])
    assert_equal(true, ps[:use_first_page_number])
    assert_equal(600, ps[:horizontal_dpi])
    assert_equal(600, ps[:vertical_dpi])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips data validation showDropDown and imeMode through writer and reader" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_data_validation("A1:A10", type: "list",
                                         formula1: '"Yes,No"',
                                         show_drop_down: true,
                                         ime_mode: "hiragana")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    dvs = reader.data_validations

    assert_equal(1, dvs.size)
    dv = dvs.first
    assert_equal(true, dv[:show_drop_down])
    assert_equal("hiragana", dv[:ime_mode])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips alignment readingOrder and justifyLastLine through writer and reader" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    sid = writer.add_cell_style(alignment: { horizontal: "distributed", reading_order: 2, justify_last_line: true })
    writer.set_cell("A1", "RTL")
    writer.set_cell_style("A1", sid)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    styles = reader.cell_styles
    xf = styles.values.find { |s| s[:alignment]&.key?(:reading_order) }
    assert_not_nil(xf)
    assert_equal(2, xf[:alignment][:reading_order])
    assert_equal(true, xf[:alignment][:justify_last_line])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips font charset attribute through writer and reader" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    fid = writer.add_font(name: "MS Gothic", sz: 11, family: 3, charset: 128)
    sid = writer.add_cell_style(font_id: fid)
    writer.set_cell("A1", "テスト")
    writer.set_cell_style("A1", sid)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    styles = reader.cell_styles
    font = styles["A1"]&.dig(:font)
    assert_not_nil(font)
    assert_equal(128, font[:charset])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips sheetFormatPr extended attributes through writer and reader" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_sheet_format(:default_row_height, 15)
    writer.set_sheet_format(:outline_level_row, 3)
    writer.set_sheet_format(:outline_level_col, 2)
    writer.set_sheet_format(:zero_height, true)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    fmt = reader.sheet_format

    assert_equal(3, fmt[:outline_level_row])
    assert_equal(2, fmt[:outline_level_col])
    assert_equal(true, fmt[:zero_height])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips quotePrefix through writer and reader" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    sid = writer.add_cell_style(quote_prefix: true)
    writer.set_cell("A1", "001234")
    writer.set_cell_style("A1", sid)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    styles = reader.cell_styles
    assert_equal(true, styles["A1"][:quote_prefix])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips showFormulas on sheet view" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_sheet_view(:show_formulas, true)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    sv = reader.sheet_view
    assert_equal(true, sv[:show_formulas])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips codeName on sheet properties" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_sheet_property(:code_name, "MySheet")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    props = reader.sheet_properties
    assert_equal("MySheet", props[:code_name])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips workbook view visibility" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_workbook_view(:visibility, "hidden")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    wv = reader.workbook_views
    assert_equal("hidden", wv[:visibility])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips phonetic properties" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_phonetic_properties({ font_id: 1, type: "Hiragana", alignment: "center" })
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    pp = reader.phonetic_properties
    assert_not_nil(pp)
    assert_equal(1, pp[:font_id])
    assert_equal("Hiragana", pp[:type])
    assert_equal("center", pp[:alignment])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips custom document properties" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.add_custom_property("Project", "Alpha", type: :lpwstr)
    writer.add_custom_property("Version", 42, type: :i4)
    writer.add_custom_property("Active", true, type: :bool)
    writer.set_cell("A1", "data")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    props = reader.custom_properties
    assert_equal(3, props.size)

    project = props.find { |p| p[:name] == "Project" }
    assert_equal("Alpha", project[:value])
    assert_equal(:string, project[:type])

    version = props.find { |p| p[:name] == "Version" }
    assert_equal(42, version[:value])
    assert_equal(:number, version[:type])

    active = props.find { |p| p[:name] == "Active" }
    assert_equal(true, active[:value])
    assert_equal(:bool, active[:type])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips font shadow outline condense extend" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    fid = writer.add_font(name: "Arial", sz: 12, shadow: true, outline: true, condense: true, extend: true)
    sid = writer.add_cell_style(font_id: fid)
    writer.set_cell("A1", "effects")
    writer.set_cell_style("A1", sid)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    styles = reader.cell_styles
    styled_font = styles.values.map { |s| s[:font] }.compact.find { |f| f[:shadow] }
    assert_not_nil(styled_font)
    assert_equal(true, styled_font[:shadow])
    assert_equal(true, styled_font[:outline])
    assert_equal(true, styled_font[:condense])
    assert_equal(true, styled_font[:extend])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips sheetView showZeros, view, showOutlineSymbols, showRuler" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_sheet_view(:show_zeros, false)
    writer.set_sheet_view(:view, "pageBreakPreview")
    writer.set_sheet_view(:show_outline_symbols, false)
    writer.set_sheet_view(:show_ruler, false)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    sv = reader.sheet_view
    assert_equal(false, sv[:show_zeros])
    assert_equal("pageBreakPreview", sv[:view])
    assert_equal(false, sv[:show_outline_symbols])
    assert_equal(false, sv[:show_ruler])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips sheetView topLeftCell, colorId, zoomScaleNormal, etc." do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_sheet_view(:window_protection, true)
    writer.set_sheet_view(:default_grid_color, false)
    writer.set_sheet_view(:show_white_space, false)
    writer.set_sheet_view(:top_left_cell, "B5")
    writer.set_sheet_view(:color_id, 10)
    writer.set_sheet_view(:zoom_scale_normal, 80)
    writer.set_sheet_view(:zoom_scale_sheet_layout_view, 75)
    writer.set_sheet_view(:zoom_scale_page_layout_view, 90)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    sv = reader.sheet_view
    assert_equal(true, sv[:window_protection])
    assert_equal(false, sv[:default_grid_color])
    assert_equal(false, sv[:show_white_space])
    assert_equal("B5", sv[:top_left_cell])
    assert_equal(10, sv[:color_id])
    assert_equal(80, sv[:zoom_scale_normal])
    assert_equal(75, sv[:zoom_scale_sheet_layout_view])
    assert_equal(90, sv[:zoom_scale_page_layout_view])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips alignment relativeIndent" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    sid = writer.add_cell_style(alignment: { indent: 2, relative_indent: -1 })
    writer.set_cell("A1", "indented")
    writer.set_cell_style("A1", sid)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    styles = reader.cell_styles
    style = styles.values.find { |s| s[:alignment] && s[:alignment][:relative_indent] }
    assert_not_nil(style)
    assert_equal(-1, style[:alignment][:relative_indent])
    assert_equal(2, style[:alignment][:indent])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips row thickTop and thickBot" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_row_thick_top(1)
    writer.set_row_thick_bot(1)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    ra = reader.row_attributes
    assert_equal(true, ra[1][:thick_top])
    assert_equal(true, ra[1][:thick_bot])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips workbook properties extended attributes" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_workbook_property(:code_name, "ThisWorkbook")
    writer.set_workbook_property(:filter_privacy, true)
    writer.set_workbook_property(:auto_compress_pictures, false)
    writer.set_workbook_property(:backup_file, true)
    writer.set_workbook_property(:show_objects, "placeholders")
    writer.set_workbook_property(:update_links, "never")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    wp = reader.workbook_properties
    assert_equal("ThisWorkbook", wp[:code_name])
    assert_equal(true, wp[:filter_privacy])
    assert_equal(false, wp[:auto_compress_pictures])
    assert_equal(true, wp[:backup_file])
    assert_equal("placeholders", wp[:show_objects])
    assert_equal("never", wp[:update_links])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips workbook view extended attributes" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_workbook_view(:show_horizontal_scroll, false)
    writer.set_workbook_view(:show_vertical_scroll, false)
    writer.set_workbook_view(:show_sheet_tabs, false)
    writer.set_workbook_view(:minimized, true)
    writer.set_workbook_view(:x_window, 100)
    writer.set_workbook_view(:y_window, 200)
    writer.set_workbook_view(:window_width, 20_000)
    writer.set_workbook_view(:window_height, 10_000)
    writer.set_workbook_view(:tab_ratio, 800)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    wv = reader.workbook_views
    assert_equal(false, wv[:show_horizontal_scroll])
    assert_equal(false, wv[:show_vertical_scroll])
    assert_equal(false, wv[:show_sheet_tabs])
    assert_equal(true, wv[:minimized])
    assert_equal(100, wv[:x_window])
    assert_equal(200, wv[:y_window])
    assert_equal(20_000, wv[:window_width])
    assert_equal(10_000, wv[:window_height])
    assert_equal(800, wv[:tab_ratio])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips calc properties extended attributes" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_calc_property(:full_precision, false)
    writer.set_calc_property(:concurrent_calc, false)
    writer.set_calc_property(:concurrent_manual_count, 4)
    writer.set_calc_property(:force_full_calc, true)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cp = reader.calc_properties
    assert_equal(false, cp[:full_precision])
    assert_equal(false, cp[:concurrent_calc])
    assert_equal(4, cp[:concurrent_manual_count])
    assert_equal(true, cp[:force_full_calc])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips sheet properties extended attributes" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_sheet_property(:filter_mode, true)
    writer.set_sheet_property(:published, false)
    writer.set_sheet_property(:fit_to_page, true)
    writer.set_sheet_property(:auto_page_breaks, false)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    sp = reader.sheet_properties
    assert_equal(true, sp[:filter_mode])
    assert_equal(false, sp[:published])
    assert_equal(true, sp[:fit_to_page])
    assert_equal(false, sp[:auto_page_breaks])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips border outline attribute" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    brd_id = writer.add_border(
      left: { style: "thin" },
      outline: false
    )
    style_id = writer.add_cell_style(border_id: brd_id)
    writer.set_cell("A1", "no-outline")
    writer.set_cell_style("A1", style_id)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cs = reader.cell_styles
    assert(cs.key?("A1"), "A1 should have a style")
    border = cs["A1"][:border]
    assert_not_nil(border, "border should be present")
    assert_equal(false, border[:outline])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips font auto color" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    font_id = writer.add_font(auto: true, size: 11, name: "Calibri")
    style_id = writer.add_cell_style(font_id: font_id)
    writer.set_cell("A1", "auto-color")
    writer.set_cell_style("A1", style_id)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cs = reader.cell_styles
    assert(cs.key?("A1"), "A1 should have a style")
    font = cs["A1"][:font]
    assert_not_nil(font, "font should be present")
    assert_equal(true, font[:auto])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips headerFooter scaleWithDoc and alignWithMargins" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_header_footer(:odd_header, "&CHello")
    writer.set_header_footer(:scale_with_doc, false)
    writer.set_header_footer(:align_with_margins, false)
    writer.set_cell("A1", "hf")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    hf = reader.header_footer
    assert_equal(false, hf[:scale_with_doc])
    assert_equal(false, hf[:align_with_margins])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips pageSetup copies, paperHeight, paperWidth, errors" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_page_setup(:copies, 3)
    writer.set_page_setup(:paper_height, "297mm")
    writer.set_page_setup(:paper_width, "210mm")
    writer.set_page_setup(:errors, "blank")
    writer.set_page_setup(:use_printer_defaults, false)
    writer.set_cell("A1", "ps")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    ps = reader.page_setup
    assert_equal(3, ps[:copies])
    assert_equal("297mm", ps[:paper_height])
    assert_equal("210mm", ps[:paper_width])
    assert_equal("blank", ps[:errors])
    assert_equal(false, ps[:use_printer_defaults])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips printOptions gridLinesSet" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_print_option(:grid_lines_set, false)
    writer.set_cell("A1", "po")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    po = reader.print_options
    assert_equal(false, po[:grid_lines_set])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips conditional format stdDev" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 100)
    writer.add_conditional_format("A1:A10",
                                  type: :above_average,
                                  std_dev: 2,
                                  format_id: writer.add_dxf(font: { bold: true }))
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    rules = reader.conditional_formats
    rule = rules.find { |r| r[:type] == "aboveAverage" }
    assert_not_nil(rule, "should find above_average rule")
    assert_equal(2, rule[:std_dev])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips cfvo gte attribute" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_conditional_format("A1:A10",
                                  type: :color_scale,
                                  color_scale: {
                                    cfvo: [{ type: "min" }, { type: "num", val: "50", gte: false }, { type: "max" }],
                                    colors: [{ rgb: "FFFF0000" }, { rgb: "FFFFFF00" }, { rgb: "FF00FF00" }]
                                  })
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    rules = reader.conditional_formats
    rule = rules.find { |r| r[:type] == "colorScale" }
    assert_not_nil(rule)
    cfvos = rule[:color_scale][:cfvo]
    assert_equal(false, cfvos[1][:gte])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips column phonetic attribute" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_column_attribute("A", :phonetic, true)
    writer.set_cell("A1", "test")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    ca = reader.column_attributes
    a_attrs = ca["A"]
    assert_not_nil(a_attrs, "column A should have attributes")
    assert_equal(true, a_attrs[:phonetic])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips named cell style iLevel, hidden, customBuiltin" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.add_named_cell_style(
      name: "Heading 1",
      builtin_id: 16,
      i_level: 0,
      hidden: true,
      custom_builtin: true
    )
    writer.set_cell("A1", "test")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    ncs = reader.named_cell_styles
    heading = ncs.find { |cs| cs[:name] == "Heading 1" }
    assert_not_nil(heading, "should find Heading 1")
    assert_equal(0, heading[:i_level])
    assert_equal(true, heading[:hidden])
    assert_equal(true, heading[:custom_builtin])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips xf pivotButton attribute" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    style_id = writer.add_cell_style(pivot_button: true)
    writer.set_cell("A1", "pivot")
    writer.set_cell_style("A1", style_id)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cs = reader.cell_styles
    assert(cs.key?("A1"))
    assert_equal(true, cs["A1"][:pivot_button])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips sheetFormatPr thickTop and thickBottom" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_sheet_format(:default_row_height, 15)
    writer.set_sheet_format(:thick_top, true)
    writer.set_sheet_format(:thick_bottom, true)
    writer.set_cell("A1", "test")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    sfp = reader.sheet_format
    assert_equal(true, sfp[:thick_top])
    assert_equal(true, sfp[:thick_bottom])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips dataValidations container options" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_data_validations_option(:disable_prompts, true)
    writer.set_data_validations_option(:x_window, 100)
    writer.set_data_validations_option(:y_window, 200)
    writer.add_data_validation("A1:A10", type: "whole", formula1: "1", formula2: "100")
    writer.set_cell("A1", 50)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    dvo = reader.data_validations_options
    assert_equal(true, dvo[:disable_prompts])
    assert_equal(100, dvo[:x_window])
    assert_equal(200, dvo[:y_window])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips border vertical and horizontal sides" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    brd_id = writer.add_border(
      vertical: { style: "thin", color: "FF00FF00" },
      horizontal: { style: "dashed", color: "FF0000FF" }
    )
    style_id = writer.add_cell_style(border_id: brd_id)
    writer.set_cell("A1", "vh")
    writer.set_cell_style("A1", style_id)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cs = reader.cell_styles
    border = cs["A1"][:border]
    assert_not_nil(border)
    assert_equal("thin", border[:vertical][:style])
    assert_equal("FF00FF00", border[:vertical][:color])
    assert_equal("dashed", border[:horizontal][:style])
    assert_equal("FF0000FF", border[:horizontal][:color])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips break extended attributes" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.add_row_break({ id: 10, min: 2, max: 8, man: true, pt: true })
    writer.set_cell("A1", "brk")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    brks = reader.row_breaks
    assert_equal(1, brks.size)
    brk = brks.first
    assert_equal(10, brk[:id])
    assert_equal(2, brk[:min])
    assert_equal(8, brk[:max])
    assert_equal(true, brk[:man])
    assert_equal(true, brk[:pt])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips row phonetic attribute" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_row_phonetic(1)
    writer.set_cell("A1", "test")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    ra = reader.row_attributes
    assert_equal(true, ra[1][:ph])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips definedName extended attributes" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.add_defined_name("MyName", "Sheet1!$A$1",
                            comment: "A comment", description: "A desc",
                            function: true, shortcut_key: "B")
    writer.set_cell("A1", "test")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    dns = reader.defined_names
    dn = dns.find { |d| d[:name] == "MyName" }
    assert_not_nil(dn)
    assert_equal("A comment", dn[:comment])
    assert_equal("A desc", dn[:description])
    assert_equal(true, dn[:function])
    assert_equal("B", dn[:shortcut_key])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips fileVersion element" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_file_version(:app_name, "xl")
    writer.set_file_version(:last_edited, "7")
    writer.set_file_version(:rup_build, "27425")
    writer.set_cell("A1", "test")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    fv = reader.file_version
    assert_equal("xl", fv[:app_name])
    assert_equal("7", fv[:last_edited])
    assert_equal("27425", fv[:rup_build])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips fileSharing element" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_file_sharing(:read_only_recommended, true)
    writer.set_file_sharing(:user_name, "TestUser")
    writer.set_cell("A1", "test")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    fs = reader.file_sharing
    assert_equal(true, fs[:read_only_recommended])
    assert_equal("TestUser", fs[:user_name])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips protectedRanges element" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_protected_range(name: "EditArea", sqref: "A1:B10")
    writer.add_protected_range(name: "SecureRange", sqref: "C1:D5", algorithm_name: "SHA-512",
                               hash_value: "abc123", salt_value: "salt456", spin_count: 100_000)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    ranges = reader.protected_ranges
    assert_equal(2, ranges.size)

    assert_equal("A1:B10", ranges[0][:sqref])
    assert_equal("EditArea", ranges[0][:name])

    assert_equal("C1:D5", ranges[1][:sqref])
    assert_equal("SecureRange", ranges[1][:name])
    assert_equal("SHA-512", ranges[1][:algorithm_name])
    assert_equal("abc123", ranges[1][:hash_value])
    assert_equal("salt456", ranges[1][:salt_value])
    assert_equal(100_000, ranges[1][:spin_count])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips indexedColors and mruColors" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_indexed_colors(%w[FF000000 FFFFFFFF FFFF0000])
    writer.set_mru_colors([{ rgb: "FF00FF00" }, { theme: 3, tint: 0.4 }])
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    ic = reader.indexed_colors
    assert_equal(%w[FF000000 FFFFFFFF FFFF0000], ic)

    mru = reader.mru_colors
    assert_equal(2, mru.size)
    assert_equal("FF00FF00", mru[0][:rgb])
    assert_equal(3, mru[1][:theme])
    assert_in_delta(0.4, mru[1][:tint])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips tableStyles in stylesheet" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    dxf_id = writer.add_dxf(font: { bold: true })
    writer.set_table_styles_option(:default_table_style, "TableStyleMedium2")
    writer.set_table_styles_option(:default_pivot_style, "PivotStyleLight16")
    writer.add_table_style(name: "MyStyle", elements: [
                             { type: "wholeTable", dxf_id: dxf_id },
                             { type: "headerRow", dxf_id: dxf_id, size: 2 }
                           ], pivot: false)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    ts = reader.table_styles
    assert_equal("TableStyleMedium2", ts[:default_table_style])
    assert_equal("PivotStyleLight16", ts[:default_pivot_style])
    assert_equal(1, ts[:styles].size)

    style = ts[:styles][0]
    assert_equal("MyStyle", style[:name])
    assert_equal(false, style[:pivot])
    assert_equal(2, style[:elements].size)
    assert_equal("wholeTable", style[:elements][0][:type])
    assert_equal(dxf_id, style[:elements][0][:dxf_id])
    assert_equal("headerRow", style[:elements][1][:type])
    assert_equal(2, style[:elements][1][:size])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips cellWatches element" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 100)
    writer.set_cell("B2", 200)
    writer.add_cell_watch("A1")
    writer.add_cell_watch("B2")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    watches = reader.cell_watches
    assert_equal(%w[A1 B2], watches)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips dataConsolidate element" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_data_consolidate(
      function: "average", start_labels: true, left_labels: true, link: true,
      data_refs: [{ ref: "A1:B10", sheet: "Sheet1" }, { ref: "C1:D10", name: "Range2" }]
    )
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    dc = reader.data_consolidate
    assert_equal("average", dc[:function])
    assert_equal(true, dc[:start_labels])
    assert_equal(true, dc[:left_labels])
    assert_equal(true, dc[:link])
    assert_equal(2, dc[:data_refs].size)
    assert_equal("A1:B10", dc[:data_refs][0][:ref])
    assert_equal("Sheet1", dc[:data_refs][0][:sheet])
    assert_equal("Range2", dc[:data_refs][1][:name])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips sheetPr sync and transition attributes" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_sheet_property(:sync_horizontal, true)
    writer.set_sheet_property(:sync_vertical, true)
    writer.set_sheet_property(:sync_ref, "A1")
    writer.set_sheet_property(:transition_evaluation, true)
    writer.set_sheet_property(:transition_entry, true)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    props = reader.sheet_properties
    assert_equal(true, props[:sync_horizontal])
    assert_equal(true, props[:sync_vertical])
    assert_equal("A1", props[:sync_ref])
    assert_equal(true, props[:transition_evaluation])
    assert_equal(true, props[:transition_entry])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips workbookPr extended attributes" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_workbook_property(:prompted_solutions, true)
    writer.set_workbook_property(:show_pivot_chart_filter, true)
    writer.set_workbook_property(:allow_refresh_query, true)
    writer.set_workbook_property(:publish_items, true)
    writer.set_workbook_property(:save_external_link_values, false)
    writer.set_workbook_property(:show_border_unselected_tables, false)
    writer.set_workbook_property(:date_compatibility, false)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    wp = reader.workbook_properties
    assert_equal(true, wp[:prompted_solutions])
    assert_equal(true, wp[:show_pivot_chart_filter])
    assert_equal(true, wp[:allow_refresh_query])
    assert_equal(true, wp[:publish_items])
    assert_equal(false, wp[:save_external_link_values])
    assert_equal(false, wp[:show_border_unselected_tables])
    assert_equal(false, wp[:date_compatibility])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips scenarios element" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 100)
    writer.set_scenarios(
      current: 0, show: 0,
      scenarios: [
        {
          name: "Best Case", user: "Admin", comment: "Optimistic",
          input_cells: [{ r: "A1", val: "200" }, { r: "B1", val: "300" }]
        }
      ]
    )
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    sc = reader.scenarios
    assert_equal(0, sc[:current])
    assert_equal(0, sc[:show])
    assert_equal(1, sc[:scenarios].size)

    scenario = sc[:scenarios][0]
    assert_equal("Best Case", scenario[:name])
    assert_equal("Admin", scenario[:user])
    assert_equal("Optimistic", scenario[:comment])
    assert_equal(2, scenario[:input_cells].size)
    assert_equal("A1", scenario[:input_cells][0][:r])
    assert_equal("200", scenario[:input_cells][0][:val])
    assert_equal("B1", scenario[:input_cells][1][:r])
    assert_equal("300", scenario[:input_cells][1][:val])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips outlinePr applyStyles attribute" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_sheet_property(:apply_styles, true)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    props = reader.sheet_properties
    assert_equal(true, props[:apply_styles])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips sheetCalcPr fullCalcOnLoad" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_sheet_property(:full_calc_on_load, true)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    props = reader.sheet_properties
    assert_equal(true, props[:full_calc_on_load])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips table extended attributes" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Name")
    writer.set_cell("B1", "Age")
    writer.add_table("A1:B5", columns: %w[Name Age],
                              header_row_count: 0, published: true, comment: "My table",
                              insert_row: true, insert_row_shift: true,
                              header_row_dxf_id: 1, data_dxf_id: 2, totals_row_dxf_id: 3)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    tbls = reader.tables
    assert_equal(1, tbls.size)
    tbl = tbls[0]
    assert_equal(0, tbl[:header_row_count])
    assert_equal(true, tbl[:published])
    assert_equal("My table", tbl[:comment])
    assert_equal(true, tbl[:insert_row])
    assert_equal(true, tbl[:insert_row_shift])
    assert_equal(1, tbl[:header_row_dxf_id])
    assert_equal(2, tbl[:data_dxf_id])
    assert_equal(3, tbl[:totals_row_dxf_id])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips table border dxfId, cellStyle, tableType, connectionId" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Name")
    writer.set_cell("B1", "Age")
    writer.add_table("A1:B5", columns: %w[Name Age],
                              header_row_border_dxf_id: 10,
                              table_border_dxf_id: 11,
                              totals_row_border_dxf_id: 12,
                              header_row_cell_style: "HeaderStyle",
                              totals_row_cell_style: "TotalsStyle",
                              table_type: "queryTable")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    tbls = reader.tables
    tbl = tbls[0]
    assert_equal(10, tbl[:header_row_border_dxf_id])
    assert_equal(11, tbl[:table_border_dxf_id])
    assert_equal(12, tbl[:totals_row_border_dxf_id])
    assert_equal("HeaderStyle", tbl[:header_row_cell_style])
    assert_equal("TotalsStyle", tbl[:totals_row_cell_style])
    assert_equal("queryTable", tbl[:table_type])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips ignoredErrors" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "123")
    writer.add_ignored_error(sqref: "A1", number_stored_as_text: true, formula_range: true)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    errors = reader.ignored_errors
    assert_equal(1, errors.size)
    assert_equal("A1", errors[0][:sqref])
    assert_equal(true, errors[0][:number_stored_as_text])
    assert_equal(true, errors[0][:formula_range])
    assert_nil(errors[0][:eval_error])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips definedName functionGroupId customMenu help statusBar" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_defined_name("Func1", "Sheet1!$A$1",
                            function_group_id: 3, custom_menu: "CM",
                            help: "H", status_bar: "SB")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    dns = reader.defined_names
    dn = dns.find { |d| d[:name] == "Func1" }
    assert_not_nil(dn)
    assert_equal(3, dn[:function_group_id])
    assert_equal("CM", dn[:custom_menu])
    assert_equal("H", dn[:help])
    assert_equal("SB", dn[:status_bar])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips conditionalFormatting pivot attribute" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_conditional_format("A1:A5", type: :cell_is, operator: "greaterThan",
                                           formula: "3", format_id: 0, priority: 1, pivot: true)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cfs = reader.conditional_formats
    assert_equal(1, cfs.size)
    assert_equal(true, cfs[0][:pivot])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips comment guid and shapeId" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "val")
    writer.add_comment("A1", "Note", guid: "{AABBCCDD-1122-3344-5566-778899AABBCC}", shape_id: 2048)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cmnts = reader.comments
    assert_equal(1, cmnts.size)
    assert_equal("{AABBCCDD-1122-3344-5566-778899AABBCC}", cmnts[0][:guid])
    assert_equal(2048, cmnts[0][:shape_id])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips autoFilter extended attributes" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Name")
    writer.set_auto_filter("A1:B10")
    writer.add_filter_column(0, { type: :filters, values: %w[X],
                                  date_group_items: [{ date_time_grouping: "month", year: 2023, month: 6 }],
                                  hidden_button: true })
    writer.add_filter_column(1, { type: :dynamic, dynamic_type: "aboveAverage",
                                  val: 50.0, max_val: 100.0 })
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    fcs = reader.filter_columns
    assert_equal(true, fcs[0][:hidden_button])
    assert_equal(1, fcs[0][:date_group_items].size)
    dg = fcs[0][:date_group_items][0]
    assert_equal("month", dg[:date_time_grouping])
    assert_equal(2023, dg[:year])
    assert_equal(6, dg[:month])
    assert_equal(50.0, fcs[1][:val])
    assert_equal(100.0, fcs[1][:max_val])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips selection pane and activeCellId" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_selection("C3", sqref: "C3", pane: "bottomLeft", active_cell_id: 2)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    sel = reader.selection
    assert_equal("C3", sel[:active_cell])
    assert_equal("bottomLeft", sel[:pane])
    assert_equal(2, sel[:active_cell_id])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips pane activePane attribute" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_freeze_pane(row: 2, col: 1)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    pane = reader.freeze_pane
    assert_equal(:frozen, pane[:state])
    assert_equal("bottomRight", pane[:active_pane])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips sortState extended attributes" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "X")
    writer.set_auto_filter("A1:B10")
    writer.set_sort_state("A2:B10",
                          [{ ref: "A2:A10", sort_by: "value", custom_list: "a,b,c" }],
                          column_sort: true, sort_method: "stroke")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    ss = reader.sort_state
    assert_equal(true, ss[:column_sort])
    assert_equal("stroke", ss[:sort_method])
    sc = ss[:sort_conditions][0]
    assert_equal("value", sc[:sort_by])
    assert_equal("a,b,c", sc[:custom_list])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "reader parses border start and end elements from Strict format" do
    # Simulate reading styles XML with start/end border elements (ISO 29500 Strict)
    require "xlsxrb/reader"
    styles_xml = <<~XML
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <borders count="1">
          <border>
            <left/><right/><top/><bottom/><diagonal/>
            <start style="thin"><color rgb="FF0000FF"/></start>
            <end style="double"><color rgb="FFFF0000"/></end>
          </border>
        </borders>
      </styleSheet>
    XML
    listener = Xlsxrb::Reader::StylesListener.new
    parser = REXML::Parsers::SAX2Parser.new(styles_xml)
    parser.listen(listener)
    parser.parse
    bdr = listener.borders[0]
    assert_equal("thin", bdr[:start][:style])
    assert_equal("double", bdr[:end][:style])
  end

  test "reader parses xf applyXxx attributes" do
    require "xlsxrb/reader"
    styles_xml = <<~XML
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <cellXfs count="1">
          <xf numFmtId="0" fontId="0" fillId="0" borderId="0"
              applyNumberFormat="1" applyFont="0" applyFill="1" applyBorder="0"
              applyAlignment="1" applyProtection="0"/>
        </cellXfs>
      </styleSheet>
    XML
    listener = Xlsxrb::Reader::StylesListener.new
    parser = REXML::Parsers::SAX2Parser.new(styles_xml)
    parser.listen(listener)
    parser.parse
    xf = listener.cell_xfs[0]
    assert_equal(true, xf[:apply_number_format])
    assert_equal(false, xf[:apply_font])
    assert_equal(true, xf[:apply_fill])
    assert_equal(false, xf[:apply_border])
    assert_equal(true, xf[:apply_alignment])
    assert_equal(false, xf[:apply_protection])
  end

  test "round-trips iconSet percent attribute" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_conditional_format("A1:A5", type: :icon_set, priority: 1,
                                           icon_set: { icon_set: "3Arrows", percent: false,
                                                       cfvo: [{ type: "min" }, { type: "num", val: "33" }, { type: "num", val: "67" }] })
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cfs = reader.conditional_formats
    assert_equal(1, cfs.size)
    assert_equal(false, cfs[0][:icon_set][:percent])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips workbook conformance attribute" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_workbook_property(:conformance, "transitional")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    assert_equal("transitional", reader.conformance)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "parses location rowPageCount and colPageCount from pivot table XML" do
    xml = <<~XML
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PT1">
        <location ref="E1:F5" firstHeaderRow="1" firstDataRow="1" firstDataCol="1" rowPageCount="2" colPageCount="3"/>
        <pivotFields count="1"><pivotField axis="axisRow" showAll="1"/></pivotFields>
        <rowFields count="1"><field x="0"/></rowFields>
        <dataFields count="0"/>
      </pivotTableDefinition>
    XML
    listener = Xlsxrb::Reader::PivotTableListener.new
    parser = REXML::Parsers::SAX2Parser.new(xml)
    parser.listen(listener)
    parser.parse
    pt = listener.pivot_table
    assert_equal 2, pt[:row_page_count]
    assert_equal 3, pt[:col_page_count]
  end

  test "round-trips chart grouping and barDir" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-chart-grp", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "a")
    writer.add_chart(type: :bar, title: "Stacked",
                     cat_ref: "Sheet1!$A$1:$A$1", val_ref: "Sheet1!$A$1:$A$1",
                     grouping: "stacked", bar_dir: "bar")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    charts = reader.charts
    assert_equal(1, charts.size)
    assert_equal("stacked", charts[0][:grouping])
    assert_equal("bar", charts[0][:bar_dir])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips pivotCacheDefinition optional attributes" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-pivot-cache", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Cat")
    writer.set_cell("B1", "Val")
    writer.set_cell("A2", "X")
    writer.set_cell("B2", 1)
    writer.add_pivot_table("Sheet1!A1:B2",
                           row_fields: [0],
                           data_fields: [{ fld: 1, name: "Sum", subtotal: "sum" }],
                           cache_save_data: false, cache_enable_refresh: false,
                           cache_refreshed_by: "Bot", cache_refreshed_version: 6,
                           cache_created_version: 5, cache_record_count: 99,
                           cache_optimize_memory: true)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    pts = reader.pivot_tables
    assert_equal(1, pts.size)
    cache = pts[0][:cache]
    assert_equal(false, cache[:save_data])
    assert_equal(false, cache[:enable_refresh])
    assert_equal("Bot", cache[:refreshed_by])
    assert_equal(6, cache[:refreshed_version])
    assert_equal(5, cache[:created_version])
    assert_equal(99, cache[:record_count])
    assert_equal(true, cache[:optimize_memory])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips cacheField caption, formula, and numFmtId" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-cache-field", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Cat")
    writer.set_cell("B1", "Val")
    writer.set_cell("A2", "X")
    writer.set_cell("B2", 1)
    writer.add_pivot_table("Sheet1!A1:B2",
                           row_fields: [0],
                           data_fields: [{ fld: 1, name: "Sum", subtotal: "sum" }],
                           field_names: %w[Cat Val],
                           field_attrs: {
                             0 => { cache_caption: "Category", cache_formula: "='Sheet1'!A1", cache_num_fmt_id: 49 }
                           })
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    pts = reader.pivot_tables
    cache = pts[0][:cache]
    assert_equal(2, cache[:fields].size)
    assert_equal("Cat", cache[:fields][0][:name])
    assert_equal("Category", cache[:fields][0][:caption])
    assert_equal("='Sheet1'!A1", cache[:fields][0][:formula])
    assert_equal(49, cache[:fields][0][:num_fmt_id])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips worksheetSource name attribute" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-ws-source", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Cat")
    writer.set_cell("B1", "Val")
    writer.set_cell("A2", "X")
    writer.set_cell("B2", 1)
    writer.add_pivot_table("Sheet1!A1:B2",
                           row_fields: [0],
                           data_fields: [{ fld: 1, name: "Sum", subtotal: "sum" }],
                           source_name: "MyNamedRange")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    pts = reader.pivot_tables
    cache = pts[0][:cache]
    assert_equal("Sheet1", cache[:source_sheet])
    assert_equal("A1:B2", cache[:source_ref])
    assert_equal("MyNamedRange", cache[:source_name])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips cacheField sharedItems" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-shared-items", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Category")
    writer.set_cell("B1", "Region")
    writer.set_cell("C1", "Amount")
    writer.add_pivot_table("Sheet1!A1:C4",
                           row_fields: [0],
                           col_fields: [1],
                           data_fields: [{ fld: 2, name: "Sum of Amount", subtotal: "sum" }],
                           field_names: %w[Category Region Amount],
                           items: { 0 => %w[A B C], 1 => %w[East West] })
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    pts = reader.pivot_tables
    cache = pts[0][:cache]
    assert_equal(3, cache[:fields].size)
    assert_equal(%w[A B C], cache[:fields][0][:shared_items])
    assert_equal(%w[East West], cache[:fields][1][:shared_items])
    assert_nil(cache[:fields][2][:shared_items])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips cacheSource type attribute" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-cache-source", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Cat")
    writer.set_cell("B1", "Val")
    writer.set_cell("A2", "X")
    writer.set_cell("B2", 1)
    writer.add_pivot_table("Sheet1!A1:B2",
                           row_fields: [0],
                           data_fields: [{ fld: 1, name: "Sum", subtotal: "sum" }])
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    pts = reader.pivot_tables
    cache = pts[0][:cache]
    assert_equal("worksheet", cache[:source_type])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "parses sharedItems date, missing, and error elements" do
    xml = <<~XML
      <?xml version="1.0" encoding="UTF-8"?>
      <pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <cacheFields count="1">
          <cacheField name="Mixed" numFmtId="0">
            <sharedItems count="5">
              <s v="Hello"/>
              <n v="42"/>
              <d v="2024-01-15T00:00:00"/>
              <m/>
              <e v="#N/A"/>
            </sharedItems>
          </cacheField>
        </cacheFields>
      </pivotCacheDefinition>
    XML

    parser = REXML::Parsers::SAX2Parser.new(xml)
    listener = Xlsxrb::Reader::PivotCacheDefinitionListener.new
    parser.listen(listener)
    parser.parse
    items = listener.cache_definition[:fields][0][:shared_items]
    assert_equal(5, items.size)
    assert_equal("Hello", items[0])
    assert_equal(42.0, items[1])
    assert_equal("2024-01-15T00:00:00", items[2])
    assert_nil(items[3])
    assert_equal("#N/A", items[4])
  end

  test "round-trips pivotCacheRecords" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-cache-records", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Category")
    writer.set_cell("B1", "Region")
    writer.set_cell("C1", "Amount")
    writer.add_pivot_table("Sheet1!A1:C4",
                           row_fields: [0],
                           col_fields: [1],
                           data_fields: [{ fld: 2, name: "Sum of Amount", subtotal: "sum" }],
                           field_names: %w[Category Region Amount],
                           items: { 0 => %w[A B C], 1 => %w[East West] })
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    pts = reader.pivot_tables
    cache = pts[0][:cache]
    records = cache[:records]
    assert_not_nil(records)
    assert_equal(3, records.size)
    # Each record has 3 fields: x (field item index), x, n (number)
    assert_equal(3, records[0].size)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "parses pivotCacheRecords with mixed element types" do
    xml = <<~XML
      <?xml version="1.0" encoding="UTF-8"?>
      <pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2">
        <r>
          <x v="0"/>
          <s v="Hello"/>
          <n v="42"/>
          <b v="1"/>
          <d v="2024-01-15T00:00:00"/>
          <m/>
          <e v="#REF!"/>
        </r>
        <r>
          <x v="1"/>
          <s v="World"/>
          <n v="99.5"/>
          <b v="0"/>
          <d v="2024-06-30T00:00:00"/>
          <m/>
          <e v="#N/A"/>
        </r>
      </pivotCacheRecords>
    XML

    parser = REXML::Parsers::SAX2Parser.new(xml)
    listener = Xlsxrb::Reader::PivotCacheRecordsListener.new
    parser.listen(listener)
    parser.parse
    records = listener.records
    assert_equal(2, records.size)
    assert_equal(7, records[0].size)
    assert_equal({ x: 0 }, records[0][0])
    assert_equal("Hello", records[0][1])
    assert_equal(42.0, records[0][2])
    assert_equal(true, records[0][3])
    assert_equal("2024-01-15T00:00:00", records[0][4])
    assert_nil(records[0][5])
    assert_equal("#REF!", records[0][6])
    assert_equal({ x: 1 }, records[1][0])
    assert_equal(false, records[1][3])
  end

  test "round-trips fonts from stylesheet" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_font(sz: 14, name: "Arial", bold: true, color: "FFFF0000")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    fonts = reader.fonts
    assert_operator(fonts.size, :>=, 2)
    custom = fonts.find { |f| f[:name] == "Arial" && f[:sz] == 14.0 } # rubocop:disable Lint/FloatComparison
    assert_not_nil(custom)
    assert_equal(true, custom[:bold])
    assert_equal("FFFF0000", custom[:color])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips fills from stylesheet" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_fill(pattern: "solid", fg_color: "FFFFFF00")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    fills = reader.fills
    assert_operator(fills.size, :>=, 3)
    custom = fills.find { |f| f[:fg_color] == "FFFFFF00" }
    assert_not_nil(custom)
    assert_equal("solid", custom[:pattern])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips borders from stylesheet" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_border(left: { style: "thin", color: "FF000000" }, right: { style: "medium" })
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    borders = reader.borders
    assert_operator(borders.size, :>=, 2)
    custom = borders.find { |b| b[:left] && b[:left][:style] == "thin" }
    assert_not_nil(custom)
    assert_equal("FF000000", custom[:left][:color])
    assert_equal("medium", custom[:right][:style])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips num_fmts from stylesheet" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    fmt_id = writer.add_number_format("#,##0.00")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    nf = reader.num_fmts
    assert_equal("#,##0.00", nf[fmt_id])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips cell_xfs from stylesheet" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    font_id = writer.add_font(sz: 12, name: "Arial", bold: true)
    style_id = writer.add_cell_style(font_id: font_id, num_fmt_id: 0)
    writer.set_cell_style("A1", style_id)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    xfs = reader.cell_xfs
    assert_operator(xfs.size, :>=, 2)
    custom = xfs.find { |xf| xf[:font_id] == font_id }
    assert_not_nil(custom)
    assert_equal(0, custom[:num_fmt_id])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips pivotField defaultSubtotal, insertBlankRow, insertPageBreak, includeNewItemsInFilter" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Cat")
    writer.set_cell("B1", "Val")
    writer.set_cell("A2", "X")
    writer.set_cell("B2", 1)
    writer.add_pivot_table("Sheet1!A1:B2",
                           row_fields: [0],
                           data_fields: [{ fld: 1, name: "Sum", subtotal: "sum" }],
                           field_attrs: {
                             0 => { default_subtotal: false, insert_blank_row: true,
                                    insert_page_break: true, include_new_items_in_filter: true }
                           })
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    pt = reader.pivot_tables.first
    field = pt[:fields].first
    assert_equal(false, field[:default_subtotal])
    assert_equal(true, field[:insert_blank_row])
    assert_equal(true, field[:insert_page_break])
    assert_equal(true, field[:include_new_items_in_filter])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips pivotTableDefinition extended display and layout attributes" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Cat")
    writer.set_cell("B1", "Val")
    writer.set_cell("A2", "X")
    writer.set_cell("B2", 1)
    writer.add_pivot_table("Sheet1!A1:B2",
                           row_fields: [0],
                           data_fields: [{ fld: 1, name: "Sum", subtotal: "sum" }],
                           multiple_field_filters: false, show_drill: false,
                           show_data_tips: false, enable_drill: false,
                           show_member_property_tips: false,
                           item_print_titles: true, field_print_titles: true,
                           preserve_formatting: false,
                           page_over_then_down: true, page_wrap: 3)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    pt = reader.pivot_tables.first
    assert_equal(false, pt[:multiple_field_filters])
    assert_equal(false, pt[:show_drill])
    assert_equal(false, pt[:show_data_tips])
    assert_equal(false, pt[:enable_drill])
    assert_equal(false, pt[:show_member_property_tips])
    assert_equal(true, pt[:item_print_titles])
    assert_equal(true, pt[:field_print_titles])
    assert_equal(false, pt[:preserve_formatting])
    assert_equal(true, pt[:page_over_then_down])
    assert_equal(3, pt[:page_wrap])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips chart plotVisOnly and dispBlanksAs attributes" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     plot_vis_only: true, disp_blanks_as: "zero")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    chart = reader.charts.first
    assert_equal(true, chart[:plot_vis_only])
    assert_equal("zero", chart[:disp_blanks_as])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips chart varyColors attribute" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :pie,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     vary_colors: true)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    chart = reader.charts.first
    assert_equal(true, chart[:vary_colors])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips chart style attribute" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     style: 26)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    chart = reader.charts.first
    assert_equal(26, chart[:style])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips legend overlay attribute" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     legend: { position: "b", overlay: true })
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    chart = reader.charts.first
    assert_equal("b", chart[:legend][:position])
    assert_equal(true, chart[:legend][:overlay])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "round-trips autoTitleDeleted attribute" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-roundtrip", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     auto_title_deleted: true)
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    chart = reader.charts.first
    assert_equal(true, chart[:auto_title_deleted])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end
end
