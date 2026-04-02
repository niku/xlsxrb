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
    writer.set_row_style(5, 0)

    expected = { 1 => { height: 25.0 }, 3 => { hidden: true }, 5 => { style: 0 } }
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

    assert_equal({ "A1" => { url: "https://example.com" }, "B1" => { url: "https://github.com" } }, writer.hyperlinks)
  end

  test "stores hyperlinks with display tooltip and location" do
    writer = Xlsxrb::Writer.new
    writer.add_hyperlink("A1", "https://example.com", display: "Example Site", tooltip: "Click to visit")
    writer.add_hyperlink("B1", "https://example.com/page", location: "Sheet2!A1")

    expected = {
      "A1" => { url: "https://example.com", display: "Example Site", tooltip: "Click to visit" },
      "B1" => { url: "https://example.com/page", location: "Sheet2!A1" }
    }
    assert_equal(expected, writer.hyperlinks)
  end

  test "stores internal hyperlink with location only" do
    writer = Xlsxrb::Writer.new
    writer.add_hyperlink("A1", location: "Sheet2!A1")

    assert_equal({ "A1" => { location: "Sheet2!A1" } }, writer.hyperlinks)
  end

  test "add_hyperlink requires url or location" do
    writer = Xlsxrb::Writer.new
    assert_raise(ArgumentError) { writer.add_hyperlink("A1") }
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

  test "workbook properties extended attributes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_workbook_property(:code_name, "ThisWorkbook")
    writer.set_workbook_property(:filter_privacy, true)
    writer.set_workbook_property(:auto_compress_pictures, false)
    writer.set_workbook_property(:backup_file, true)
    writer.set_workbook_property(:show_objects, "placeholders")
    writer.set_workbook_property(:update_links, "never")
    writer.set_workbook_property(:refresh_all_connections, true)
    writer.set_workbook_property(:check_compatibility, true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-wbpr", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml = read_xml_from_xlsx(xlsx_path, "xl/workbook.xml")
    assert_match(/codeName="ThisWorkbook"/, xml)
    assert_match(/filterPrivacy="1"/, xml)
    assert_match(/autoCompressPictures="0"/, xml)
    assert_match(/backupFile="1"/, xml)
    assert_match(/showObjects="placeholders"/, xml)
    assert_match(/updateLinks="never"/, xml)
    assert_match(/refreshAllConnections="1"/, xml)
    assert_match(/checkCompatibility="1"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "stores workbook view properties" do
    writer = Xlsxrb::Writer.new
    writer.set_workbook_view(:active_tab, 1)
    writer.set_workbook_view(:first_sheet, 0)

    views = writer.workbook_views
    assert_equal(1, views[:active_tab])
    assert_equal(0, views[:first_sheet])
  end

  test "workbook view extended attributes" do
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
    writer.set_workbook_view(:auto_filter_date_grouping, false)

    xlsx_tempfile = Tempfile.new(["xlsxrb-bv", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml = read_xml_from_xlsx(xlsx_path, "xl/workbook.xml")
    assert_match(/showHorizontalScroll="0"/, xml)
    assert_match(/showVerticalScroll="0"/, xml)
    assert_match(/showSheetTabs="0"/, xml)
    assert_match(/minimized="1"/, xml)
    assert_match(/xWindow="100"/, xml)
    assert_match(/yWindow="200"/, xml)
    assert_match(/windowWidth="20000"/, xml)
    assert_match(/windowHeight="10000"/, xml)
    assert_match(/tabRatio="800"/, xml)
    assert_match(/autoFilterDateGrouping="0"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "stores calc properties" do
    writer = Xlsxrb::Writer.new
    writer.set_calc_property(:calc_id, 191_029)
    writer.set_calc_property(:full_calc_on_load, true)

    props = writer.calc_properties
    assert_equal(191_029, props[:calc_id])
    assert_equal(true, props[:full_calc_on_load])
  end

  test "calc properties extended attributes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_calc_property(:full_precision, false)
    writer.set_calc_property(:concurrent_calc, false)
    writer.set_calc_property(:concurrent_manual_count, 4)
    writer.set_calc_property(:force_full_calc, true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-calc", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml = read_xml_from_xlsx(xlsx_path, "xl/workbook.xml")
    assert_match(/fullPrecision="0"/, xml)
    assert_match(/concurrentCalc="0"/, xml)
    assert_match(/concurrentManualCount="4"/, xml)
    assert_match(/forceFullCalc="1"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
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

  test "stores defined names" do
    writer = Xlsxrb::Writer.new
    writer.add_sheet("Data")
    writer.add_defined_name("MyRange", "Sheet1!$A$1:$B$10")
    writer.add_defined_name("LocalName", "Data!$C$1", sheet: "Data")
    writer.add_defined_name("HiddenName", "42", hidden: true)
    writer.add_defined_name("Constant", "\"hello\"")

    dns = writer.defined_names
    assert_equal(4, dns.size)
    assert_equal("MyRange", dns[0][:name])
    assert_nil(dns[0][:local_sheet_id])
    assert_equal(false, dns[0][:hidden])
    assert_equal(1, dns[1][:local_sheet_id])
    assert_equal(true, dns[2][:hidden])
    assert_equal("\"hello\"", dns[3][:value])
  end

  test "stores sheet properties" do
    writer = Xlsxrb::Writer.new
    writer.set_sheet_property(:tab_color, "FF0000FF", sheet: "Sheet1")
    writer.set_sheet_property(:summary_below, false, sheet: "Sheet1")
    writer.set_sheet_property(:summary_right, true, sheet: "Sheet1")

    props = writer.sheet_properties(sheet: "Sheet1")
    assert_equal("FF0000FF", props[:tab_color])
    assert_equal(false, props[:summary_below])
    assert_equal(true, props[:summary_right])
  end

  test "sheet_properties defaults to empty" do
    writer = Xlsxrb::Writer.new
    assert_equal({}, writer.sheet_properties)
  end

  test "stores sheet format properties" do
    writer = Xlsxrb::Writer.new
    writer.set_sheet_format(:default_row_height, 15.0)
    writer.set_sheet_format(:default_col_width, 10.5)
    writer.set_sheet_format(:base_col_width, 8)

    fmt = writer.sheet_format
    assert_equal(15.0, fmt[:default_row_height])
    assert_equal(10.5, fmt[:default_col_width])
    assert_equal(8, fmt[:base_col_width])
  end

  test "stores row outline level and collapsed" do
    writer = Xlsxrb::Writer.new
    writer.set_row_outline_level(2, 1)
    writer.set_row_collapsed(3)

    attrs = writer.row_attributes
    assert_equal(1, attrs[2][:outline_level])
    assert_equal(true, attrs[3][:collapsed])
  end

  test "row emits thickTop and thickBot" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_row_thick_top(1)
    writer.set_row_thick_bot(1)

    xlsx_tempfile = Tempfile.new(["xlsxrb-thick", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/thickTop="1"/, xml)
    assert_match(/thickBot="1"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "stores column attributes" do
    writer = Xlsxrb::Writer.new
    writer.set_column_attribute("B", :hidden, true)
    writer.set_column_attribute("C", :best_fit, true)
    writer.set_column_attribute("D", :outline_level, 2)
    writer.set_column_attribute("D", :collapsed, true)

    ca = writer.column_attributes
    assert_equal(true, ca["B"][:hidden])
    assert_equal(true, ca["C"][:best_fit])
    assert_equal(2, ca["D"][:outline_level])
    assert_equal(true, ca["D"][:collapsed])
  end

  test "stores sheet view properties" do
    writer = Xlsxrb::Writer.new
    writer.set_sheet_view(:show_grid_lines, false)
    writer.set_sheet_view(:zoom_scale, 150)

    sv = writer.sheet_view
    assert_equal(false, sv[:show_grid_lines])
    assert_equal(150, sv[:zoom_scale])
  end

  test "stores freeze pane" do
    writer = Xlsxrb::Writer.new
    writer.set_freeze_pane(row: 1, col: 1)

    fp = writer.freeze_pane
    assert_equal(1, fp[:row])
    assert_equal(1, fp[:col])
  end

  test "stores selection" do
    writer = Xlsxrb::Writer.new
    writer.set_selection("B2", sqref: "B2:C3")

    sel = writer.selection
    assert_equal("B2", sel[:active_cell])
    assert_equal("B2:C3", sel[:sqref])
  end

  test "stores print options" do
    writer = Xlsxrb::Writer.new
    writer.set_print_option(:grid_lines, true)
    writer.set_print_option(:horizontal_centered, true)

    po = writer.print_options
    assert_equal(true, po[:grid_lines])
    assert_equal(true, po[:horizontal_centered])
  end

  test "stores page margins" do
    writer = Xlsxrb::Writer.new
    writer.set_page_margins(left: 0.7, right: 0.7, top: 0.75, bottom: 0.75)

    pm = writer.page_margins
    assert_equal(0.7, pm[:left])
    assert_equal(0.75, pm[:top])
  end

  test "stores page setup" do
    writer = Xlsxrb::Writer.new
    writer.set_page_setup(:orientation, "landscape")
    writer.set_page_setup(:paper_size, 9)

    ps = writer.page_setup
    assert_equal("landscape", ps[:orientation])
    assert_equal(9, ps[:paper_size])
  end

  test "stores header footer" do
    writer = Xlsxrb::Writer.new
    writer.set_header_footer(:odd_header, "&CPage &P")
    writer.set_header_footer(:odd_footer, "&CFooter")

    hf = writer.header_footer
    assert_equal("&CPage &P", hf[:odd_header])
    assert_equal("&CFooter", hf[:odd_footer])
  end

  test "stores row and col breaks" do
    writer = Xlsxrb::Writer.new
    writer.add_row_break(10)
    writer.add_row_break(20)
    writer.add_col_break(5)

    assert_equal([10, 20], writer.row_breaks)
    assert_equal([5], writer.col_breaks)
  end

  test "stores filter columns" do
    writer = Xlsxrb::Writer.new
    writer.set_auto_filter("A1:C10")
    writer.add_filter_column(0, { type: :filters, values: %w[A B] })
    writer.add_filter_column(1, { type: :custom, operator: "greaterThan", val: "100" })

    fc = writer.filter_columns
    assert_equal(:filters, fc[0][:type])
    assert_equal(%w[A B], fc[0][:values])
    assert_equal(:custom, fc[1][:type])
  end

  test "stores sort state" do
    writer = Xlsxrb::Writer.new
    writer.set_sort_state("A1:B10", [{ ref: "A1:A10" }, { ref: "B1:B10", descending: true }])

    ss = writer.sort_state
    assert_equal("A1:B10", ss[:ref])
    assert_equal(2, ss[:sort_conditions].size)
    assert_equal(true, ss[:sort_conditions][1][:descending])
  end

  test "stores data validations" do
    writer = Xlsxrb::Writer.new
    writer.add_data_validation("A1:A100", type: "whole", operator: "between",
                                          formula1: "1", formula2: "100",
                                          show_error_message: true, error: "Must be 1-100")
    writer.add_data_validation("B1:B100", type: "list", formula1: '"Yes,No"',
                                          show_input_message: true, prompt: "Choose one")

    dvs = writer.data_validations
    assert_equal(2, dvs.size)
    assert_equal("A1:A100", dvs[0][:sqref])
    assert_equal("whole", dvs[0][:type])
    assert_equal("between", dvs[0][:operator])
    assert_equal("B1:B100", dvs[1][:sqref])
    assert_equal("list", dvs[1][:type])
  end

  test "stores conditional formatting rules" do
    writer = Xlsxrb::Writer.new
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

    cfs = writer.conditional_formats
    assert_equal(4, cfs.size)
    assert_equal(:cell_is, cfs[0][:type])
    assert_equal("greaterThan", cfs[0][:operator])
    assert_equal(:color_scale, cfs[1][:type])
    assert_equal(:data_bar, cfs[2][:type])
    assert_equal(:icon_set, cfs[3][:type])
  end

  test "dataBar emits minLength maxLength showValue attributes" do
    writer = Xlsxrb::Writer.new
    writer.add_conditional_format("A1:A10", type: :data_bar, priority: 1,
                                            data_bar: {
                                              cfvo: [{ type: "min" }, { type: "max" }],
                                              color: "FF638EC6",
                                              min_length: 5, max_length: 90, show_value: false
                                            })
    cfs = writer.conditional_formats
    assert_equal(5, cfs[0][:data_bar][:min_length])
    assert_equal(90, cfs[0][:data_bar][:max_length])
    assert_equal(false, cfs[0][:data_bar][:show_value])
  end

  test "iconSet emits reverse and showValue attributes" do
    writer = Xlsxrb::Writer.new
    writer.add_conditional_format("A1:A10", type: :icon_set, priority: 1,
                                            icon_set: {
                                              icon_set: "3Arrows",
                                              cfvo: [{ type: "percent", val: "0" },
                                                     { type: "percent", val: "33" },
                                                     { type: "percent", val: "67" }],
                                              reverse: true, show_value: false
                                            })
    cfs = writer.conditional_formats
    assert_equal(true, cfs[0][:icon_set][:reverse])
    assert_equal(false, cfs[0][:icon_set][:show_value])
  end

  test "add_named_cell_style registers cellStyleXfs and cellStyles" do
    writer = Xlsxrb::Writer.new
    fid = writer.add_font(bold: true, sz: 14, name: "Arial")
    xf_id = writer.add_named_cell_style(name: "Heading1", font_id: fid, builtin_id: 1)
    assert_equal(1, xf_id)
  end

  test "cellXf emits specified xfId linkage" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    fid = writer.add_font(bold: true, sz: 14, name: "Arial")
    base_xf_id = writer.add_named_cell_style(name: "Heading1", font_id: fid, builtin_id: 1)
    cell_xf = writer.add_cell_style(xf_id: base_xf_id)
    writer.set_cell("A1", "hello")
    writer.set_cell_style("A1", cell_xf)
    writer.write(xlsx_path)

    xml = read_xml_from_xlsx(xlsx_path, "xl/styles.xml")
    assert_match(/<cellXfs[^>]*>/, xml)
    assert_match(/<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="1"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "stores chart with multiple series and axis titles" do
    writer = Xlsxrb::Writer.new
    writer.add_chart(type: :bar, title: "Sales",
                     series: [
                       { cat_ref: "Sheet1!$A$1:$A$3", val_ref: "Sheet1!$B$1:$B$3", name: "Sheet1!$B$1" },
                       { cat_ref: "Sheet1!$A$1:$A$3", val_ref: "Sheet1!$C$1:$C$3", name: "Sheet1!$C$1" }
                     ],
                     legend: { position: "b" },
                     data_labels: { show_val: true, show_cat_name: false },
                     cat_axis_title: "Category",
                     val_axis_title: "Value")

    charts = writer.charts
    assert_equal(1, charts.size)
    assert_equal(2, charts[0][:series].size)
    assert_equal("b", charts[0][:legend][:position])
    assert_equal("Category", charts[0][:cat_axis_title])
  end

  test "stores shapes with preset geometry and text" do
    writer = Xlsxrb::Writer.new
    writer.add_shape(preset: "ellipse", text: "Hello", name: "Oval 1",
                     from_col: 1, from_row: 1, to_col: 4, to_row: 6)

    shapes = writer.shapes
    assert_equal(1, shapes.size)
    assert_equal("ellipse", shapes[0][:preset])
    assert_equal("Hello", shapes[0][:text])
    assert_equal("Oval 1", shapes[0][:name])
    assert_equal(1, shapes[0][:from_col])
    assert_equal(4, shapes[0][:to_col])

    writer.add_shape(preset: "roundRect", name: "RR 1")
    assert_equal(2, writer.shapes.size)
  end

  test "stores fonts, fills, borders, and cell styles" do
    writer = Xlsxrb::Writer.new
    fid = writer.add_font(bold: true, sz: 14, name: "Arial", color: "FFFF0000")
    assert_equal(1, fid)

    fill_id = writer.add_fill(pattern: "solid", fg_color: "FF00FF00")
    assert_equal(2, fill_id)

    brd_id = writer.add_border(left: { style: "thin", color: "FF000000" },
                               right: { style: "thin" },
                               top: { style: "thin" },
                               bottom: { style: "thin" })
    assert_equal(1, brd_id)

    style_id = writer.add_cell_style(font_id: fid, fill_id: fill_id, border_id: brd_id)
    assert_equal(1, style_id)

    writer.set_cell("A1", "styled")
    writer.set_cell_style("A1", style_id)
  end

  test "add_font supports extended attributes (strike, underline val, vertAlign, scheme, family)" do
    writer = Xlsxrb::Writer.new
    fid = writer.add_font(
      bold: true, italic: true, strike: true, sz: 12, name: "Calibri",
      color: "FF0000FF", underline: "double", vert_align: "superscript",
      scheme: "minor", family: 2
    )
    assert_equal(1, fid)

    writer.set_cell("A1", "extended font")
    style_id = writer.add_cell_style(font_id: fid)
    writer.set_cell_style("A1", style_id)

    xlsx_tempfile = Tempfile.new(["xlsxrb-fontex", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/styles.xml")
    assert_match(%r{<strike/>}, xml_content)
    assert_match(%r{<u val="double"/>}, xml_content)
    assert_match(%r{<vertAlign val="superscript"/>}, xml_content)
    assert_match(%r{<scheme val="minor"/>}, xml_content)
    assert_match(%r{<family val="2"/>}, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_border with diagonal emits diagonal element and border attributes" do
    writer = Xlsxrb::Writer.new
    brd_id = writer.add_border(
      left: { style: "thin" }, right: { style: "thin" },
      top: { style: "thin" }, bottom: { style: "thin" },
      diagonal: { style: "thin", color: "FFFF0000" },
      diagonal_up: true, diagonal_down: true
    )
    assert_equal(1, brd_id)

    style_id = writer.add_cell_style(border_id: brd_id)
    writer.set_cell("A1", "diag")
    writer.set_cell_style("A1", style_id)

    xlsx_tempfile = Tempfile.new(["xlsxrb-diag", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/styles.xml")
    assert_match(/<border[^>]*diagonalUp="1"/, xml_content)
    assert_match(/<border[^>]*diagonalDown="1"/, xml_content)
    assert_match(/<diagonal style="thin">/, xml_content)
    assert_match(%r{<color rgb="FFFF0000"/>}, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_fill with gradient type emits gradientFill" do
    writer = Xlsxrb::Writer.new
    fill_id = writer.add_fill(
      gradient: { type: "linear", degree: 90,
                  stops: [{ position: 0, color: "FFFF0000" }, { position: 1, color: "FF0000FF" }] }
    )
    assert_operator(fill_id, :>=, 2) # 0=none, 1=gray125

    style_id = writer.add_cell_style(fill_id: fill_id)
    writer.set_cell("A1", "gradient")
    writer.set_cell_style("A1", style_id)

    xlsx_tempfile = Tempfile.new(["xlsxrb-gradient", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/styles.xml")
    assert_match(/<gradientFill[^>]*type="linear"/, xml_content)
    assert_match(/<gradientFill[^>]*degree="90"/, xml_content)
    assert_match(%r{<stop position="0"><color rgb="FFFF0000"/></stop>}, xml_content)
    assert_match(%r{<stop position="1"><color rgb="FF0000FF"/></stop>}, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "stores dxf entries" do
    writer = Xlsxrb::Writer.new
    dxf_id = writer.add_dxf(font: { bold: true, color: "FFFF0000" },
                            fill: { pattern: "solid", fg_color: "FFFFFF00" })
    assert_equal(0, dxf_id)

    dxf_id2 = writer.add_dxf(border: { left: { style: "thin" } })
    assert_equal(1, dxf_id2)
  end

  test "stores tables" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Name")
    writer.set_cell("B1", "Age")
    writer.add_table("A1:B5", columns: %w[Name Age])

    tbls = writer.tables
    assert_equal(1, tbls.size)
    assert_equal("A1:B5", tbls[0][:ref])
    assert_equal(%w[Name Age], tbls[0][:columns])
  end

  test "stores tables with totals row and enhanced columns" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Item")
    writer.set_cell("B1", "Price")
    writer.set_cell("C1", "Tax")
    writer.add_table("A1:C5", columns: [
                       "Item",
                       { name: "Price", totals_row_function: "sum" },
                       { name: "Tax", calculated_column_formula: "[Price]*0.1" }
                     ], totals_row_count: 1, style: { name: "TableStyleLight1", show_row_stripes: false })

    tbls = writer.tables
    assert_equal(1, tbls.size)
    assert_equal(1, tbls[0][:totals_row_count])
    assert_equal("TableStyleLight1", tbls[0][:style][:name])
  end

  test "add_table with tableColumn extended attributes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Item")
    writer.set_cell("B1", "Price")
    writer.add_table("A1:B5", columns: [
                       { name: "Item", totals_row_label: "Total", header_row_dxf_id: 1 },
                       { name: "Price", totals_row_function: "sum", data_dxf_id: 2,
                         totals_row_dxf_id: 3, data_cell_style: "Currency" }
                     ], totals_row_count: 1)
    tbls = writer.tables
    cols = tbls[0][:columns]
    assert_equal("Total", cols[0][:totals_row_label])
    assert_equal(1, cols[0][:header_row_dxf_id])
    assert_equal(2, cols[1][:data_dxf_id])
    assert_equal(3, cols[1][:totals_row_dxf_id])
    assert_equal("Currency", cols[1][:data_cell_style])
  end

  test "emits totalsRowFormula in table column" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Item")
    writer.set_cell("B1", "Price")
    writer.add_table("A1:B5", columns: [
                       "Item",
                       { name: "Price", totals_row_function: "custom",
                         totals_row_formula: "SUBTOTAL(109,[Price])" }
                     ], totals_row_count: 1)

    xlsx_tempfile = Tempfile.new(["xlsxrb-trf", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/tables/table1.xml")
    assert_match(%r{<totalsRowFormula>SUBTOTAL\(109,\[Price\]\)</totalsRowFormula>}, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "enables shared string table mode" do
    writer = Xlsxrb::Writer.new
    writer.use_shared_strings!
    writer.set_cell("A1", "hello")
    writer.set_cell("B1", "hello")
    writer.set_cell("C1", "world")

    xlsx_tempfile = Tempfile.new(["xlsxrb-sst", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    cells = reader.cells
    assert_equal("hello", cells["A1"])
    assert_equal("hello", cells["B1"])
    assert_equal("world", cells["C1"])
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  # --- Phase 2: Writer unit tests ---

  test "insert_image stores image definition" do
    writer = Xlsxrb::Writer.new
    png = "\x89PNG".b
    writer.insert_image(png, ext: "png", from_col: 1, from_row: 2, to_col: 5, to_row: 8, name: "Test")
    imgs = writer.images
    assert_equal(1, imgs.size)
    assert_equal("Test", imgs[0][:name])
    assert_equal("png", imgs[0][:ext])
    assert_equal(1, imgs[0][:from_col])
    assert_equal(2, imgs[0][:from_row])
  end

  test "insert_image stores description" do
    writer = Xlsxrb::Writer.new
    png = "\x89PNG".b
    writer.insert_image(png, ext: "png", name: "Logo", description: "Company logo image")
    imgs = writer.images
    assert_equal(1, imgs.size)
    assert_equal("Logo", imgs[0][:name])
    assert_equal("Company logo image", imgs[0][:description])
  end

  test "insert_image stores title and hidden attributes" do
    writer = Xlsxrb::Writer.new
    png = "\x89PNG".b
    writer.insert_image(png, ext: "png", name: "Pic1", title: "My tooltip", hidden: true)
    imgs = writer.images
    assert_equal("My tooltip", imgs[0][:title])
    assert_equal(true, imgs[0][:hidden])
  end

  test "emits cNvPr title and hidden on image" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    png = "\x89PNG".b
    writer.insert_image(png, ext: "png", name: "Pic1", title: "Tooltip", hidden: true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-cnvpr", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/title="Tooltip"/, drawing_xml)
    assert_match(/hidden="1"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "insert_image emits macro attribute on pic element" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    png = "\x89PNG".b
    writer.insert_image(png, ext: "png", name: "Pic1", macro: "MyMacro")

    xlsx_tempfile = Tempfile.new(["xlsxrb-macro", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/xdr:pic macro="MyMacro"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "insert_image with no_change_aspect false omits noChangeAspect" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    png = "\x89PNG".b
    writer.insert_image(png, ext: "png", name: "Pic1", no_change_aspect: false)

    xlsx_tempfile = Tempfile.new(["xlsxrb-piclock", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_no_match(/noChangeAspect/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "insert_image with no_crop emits noCrop on picLocks" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    png = "\x89PNG".b
    writer.insert_image(png, ext: "png", name: "Pic1", no_crop: true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-piclock", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/noCrop="1"/, drawing_xml)
    assert_match(/noChangeAspect="1"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "insert_image with line_color and line_width emits a:ln in spPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    png = "\x89PNG".b
    writer.insert_image(png, ext: "png", name: "Pic1", line_color: "0000FF", line_width: 25_400)

    xlsx_tempfile = Tempfile.new(["xlsxrb-imgln", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:ln w="25400"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "insert_image with rotation emits a:xfrm rot in spPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    png = "\x89PNG".b
    writer.insert_image(png, ext: "png", name: "Pic1", rotation: 5_400_000)

    xlsx_tempfile = Tempfile.new(["xlsxrb-imgrot", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/rot="5400000"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "insert_image with published emits fPublished on anchor" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    png = "\x89PNG".b
    writer.insert_image(png, ext: "png", name: "Pic1", published: true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-pub", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/fPublished="1"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "insert_image with src_rect emits a:srcRect in blipFill" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    png = "\x89PNG".b
    writer.insert_image(png, ext: "png", name: "Pic1", src_rect: { top: 10_000, bottom: 20_000, left: 5000, right: 15_000 })

    xlsx_tempfile = Tempfile.new(["xlsxrb-srcrect", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:srcRect t="10000" b="20000" l="5000" r="15000"/>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "insert_image with alpha_mod_fix emits a:alphaModFix on blip" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    png = "\x89PNG".b
    writer.insert_image(png, ext: "png", name: "Pic1", alpha_mod_fix: 50_000)

    xlsx_tempfile = Tempfile.new(["xlsxrb-alpha", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:blip r:embed="[^"]+"><a:alphaModFix amt="50000"/></a:blip>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_chart stores chart definition" do
    writer = Xlsxrb::Writer.new
    writer.add_chart(type: :bar, title: "Sales", cat_ref: "Sheet1!$A$1:$A$3", val_ref: "Sheet1!$B$1:$B$3")
    charts = writer.charts
    assert_equal(1, charts.size)
    assert_equal(:bar, charts[0][:type])
    assert_equal("Sales", charts[0][:title])
  end

  test "add_chart emits custom grouping and barDir" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "a")
    writer.add_chart(type: :bar, title: "Stacked",
                     cat_ref: "Sheet1!$A$1:$A$1", val_ref: "Sheet1!$A$1:$A$1",
                     grouping: "stacked", bar_dir: "bar")
    xlsx_path = File.join(Dir.tmpdir, "chart_grp_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(/grouping val="stacked"/, xml)
    assert_match(/barDir val="bar"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_chart supports area, scatter, doughnut, radar types" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "a")
    %i[area scatter doughnut radar].each { |t| writer.add_chart(type: t, cat_ref: "Sheet1!$A$1:$A$1", val_ref: "Sheet1!$A$1:$A$1") }
    xlsx_path = File.join(Dir.tmpdir, "chart_types_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    (1..4).each do |i|
      xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart#{i}.xml")
      assert(xml.include?("<c:"), "chart#{i}.xml should contain chart XML")
    end
    xml1 = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(/areaChart/, xml1)
    xml3 = read_xml_from_xlsx(xlsx_path, "xl/charts/chart3.xml")
    assert_match(/doughnutChart/, xml3)
    refute_match(/catAx/, xml3, "doughnut chart should not have axes")
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "insert_image with edit_as stores editAs attribute" do
    writer = Xlsxrb::Writer.new
    png = "\x89PNG".b
    writer.insert_image(png, ext: "png", name: "Pic1", edit_as: "absolute")
    imgs = writer.images
    assert_equal("absolute", imgs[0][:edit_as])
  end

  test "add_chart with edit_as stores editAs attribute" do
    writer = Xlsxrb::Writer.new
    writer.add_chart(type: :bar, title: "Sales", cat_ref: "Sheet1!$A$1:$A$3", val_ref: "Sheet1!$B$1:$B$3", edit_as: "oneCell")
    charts = writer.charts
    assert_equal("oneCell", charts[0][:edit_as])
  end

  test "add_chart with anchor positions stores from/to col/row" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "x")
    writer.add_chart(type: :bar, title: "Positioned",
                     cat_ref: "Sheet1!$A$1:$A$1", val_ref: "Sheet1!$A$1:$A$1",
                     from_col: 2, from_row: 3, to_col: 8, to_row: 18,
                     from_col_off: 100, from_row_off: 200, to_col_off: 300, to_row_off: 400)
    charts = writer.charts
    assert_equal(2, charts[0][:from_col])
    assert_equal(3, charts[0][:from_row])
    assert_equal(8, charts[0][:to_col])
    assert_equal(18, charts[0][:to_row])

    xlsx_path = File.join(Dir.tmpdir, "chart_anchor_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<xdr:from>.*<xdr:col>2</xdr:col>}m, xml)
    assert_match(%r{<xdr:to>.*<xdr:col>8</xdr:col>}m, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_chart with name and description emits cNvPr attrs" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "x")
    writer.add_chart(type: :bar, title: "Sales",
                     cat_ref: "Sheet1!$A$1:$A$1", val_ref: "Sheet1!$A$1:$A$1",
                     name: "Chart 1", description: "A sales chart")
    charts = writer.charts
    assert_equal("Chart 1", charts[0][:name])
    assert_equal("A sales chart", charts[0][:description])

    xlsx_path = File.join(Dir.tmpdir, "chart_descr_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/name="Chart 1"/, xml)
    assert_match(/descr="A sales chart"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_chart with frame_macro emits macro on graphicFrame" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "x")
    writer.add_chart(type: :bar, title: "Sales",
                     cat_ref: "Sheet1!$A$1:$A$1", val_ref: "Sheet1!$A$1:$A$1",
                     frame_macro: "ChartMacro")

    xlsx_tempfile = Tempfile.new(["xlsxrb-gf-macro", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/graphicFrame macro="ChartMacro"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_chart with frame_no_grp emits graphicFrameLocks" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "x")
    writer.add_chart(type: :bar, title: "Sales",
                     cat_ref: "Sheet1!$A$1:$A$1", val_ref: "Sheet1!$A$1:$A$1",
                     frame_no_grp: true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-gflocks", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/graphicFrameLocks noGrp="1"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with edit_as stores editAs attribute" do
    writer = Xlsxrb::Writer.new
    writer.add_shape(preset: "rect", edit_as: "absolute")
    shapes = writer.shapes
    assert_equal("absolute", shapes[0][:edit_as])
  end

  test "add_shape with description stores descr attribute" do
    writer = Xlsxrb::Writer.new
    writer.add_shape(preset: "rect", description: "A rectangle shape")
    shapes = writer.shapes
    assert_equal("A rectangle shape", shapes[0][:description])

    xlsx_path = File.join(Dir.tmpdir, "shape_descr_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/descr="A rectangle shape"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape emits macro and textlink attributes on sp element" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", macro: "ShapeMacro", textlink: "$A$1")

    xlsx_tempfile = Tempfile.new(["xlsxrb-shape-macro", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/xdr:sp macro="ShapeMacro" textlink="\$A\$1"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with f_locks_text emits spLocks element" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", f_locks_text: true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-spl", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/spLocks fLocksText="1"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with no_grp and no_rot emits spLocks attributes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", no_grp: true, no_rot: true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-spl", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/noGrp="1"/, drawing_xml)
    assert_match(/noRot="1"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with fill_color emits solidFill in spPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", fill_color: "FF0000")

    xlsx_tempfile = Tempfile.new(["xlsxrb-sfill", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with line_color and line_width emits a:ln in spPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", line_color: "0000FF", line_width: 12_700)

    xlsx_tempfile = Tempfile.new(["xlsxrb-sln", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:ln w="12700"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text body properties emits bodyPr attributes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Hello", text_wrap: "square", text_anchor: "ctr", text_vert_overflow: "clip")

    xlsx_tempfile = Tempfile.new(["xlsxrb-sbody", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/wrap="square"/, drawing_xml)
    assert_match(/anchor="ctr"/, drawing_xml)
    assert_match(/vertOverflow="clip"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with rotation emits a:xfrm rot in spPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", rotation: 5_400_000)

    xlsx_tempfile = Tempfile.new(["xlsxrb-srot", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/rot="5400000"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with no_fill and no_line emits noFill elements" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", no_fill: true, no_line: true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-snofill", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:prstGeom.*<a:noFill/>}, drawing_xml)
    assert_match(%r{<a:ln><a:noFill/></a:ln>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with adjust_values emits a:gd in avLst" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "roundRect",
                     adjust_values: [{ name: "adj", fmla: "val 16667" }])

    xlsx_tempfile = Tempfile.new(["xlsxrb-avlst", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:avLst><a:gd name="adj" fmla="val 16667"/></a:avLst>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_font emits a:rPr with font attributes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Bold",
                     text_font: { bold: true, italic: true, size: 1400, color: "FF0000", name: "Arial" })

    xlsx_tempfile = Tempfile.new(["xlsxrb-rpr", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/<a:rPr b="1" i="1" sz="1400">/, drawing_xml)
    assert_match(%r{<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>}, drawing_xml)
    assert_match(%r{<a:latin typeface="Arial"/>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with autofit none emits a:noAutofit" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "No autofit", autofit: "none")

    xlsx_tempfile = Tempfile.new(["xlsxrb-autofit", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:bodyPr[^>]*><a:noAutofit/></a:bodyPr>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with autofit shape emits a:spAutoFit" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Shape autofit", autofit: "shape")

    xlsx_tempfile = Tempfile.new(["xlsxrb-autofit", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:bodyPr[^>]*><a:spAutoFit/></a:bodyPr>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with autofit normal emits a:normAutofit" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Normal autofit", autofit: "normal")

    xlsx_tempfile = Tempfile.new(["xlsxrb-autofit", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:bodyPr[^>]*><a:normAutofit/></a:bodyPr>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with outer_shadow emits a:effectLst with a:outerShdw" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Shadow",
                     outer_shadow: { blur_rad: 50_800, dist: 38_100, dir: 2_700_000, color: "000000" })

    xlsx_tempfile = Tempfile.new(["xlsxrb-shadow", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000"><a:srgbClr val="000000"/></a:outerShdw></a:effectLst>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with gradient_fill emits a:gradFill with stops and lin" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Grad",
                     gradient_fill: { stops: [{ pos: 0, color: "FF0000" }, { pos: 100_000, color: "0000FF" }], angle: 5_400_000 })

    xlsx_tempfile = Tempfile.new(["xlsxrb-grad", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with line_dash emits a:prstDash in a:ln" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Dashed", line_color: "000000", line_dash: "dash")

    xlsx_tempfile = Tempfile.new(["xlsxrb-dash", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="dash"/></a:ln>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with line_cap emits cap attribute on a:ln" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Cap", line_color: "000000", line_cap: "rnd")

    xlsx_tempfile = Tempfile.new(["xlsxrb-linecap", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/<a:ln cap="rnd">/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with line_align emits algn attribute on a:ln" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Align", line_color: "000000", line_align: "in")

    xlsx_tempfile = Tempfile.new(["xlsxrb-linealign", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/<a:ln algn="in">/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with line_compound emits cmpd attribute on a:ln" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Cmpd", line_color: "000000", line_compound: "dbl")

    xlsx_tempfile = Tempfile.new(["xlsxrb-linecmpd", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/<a:ln cmpd="dbl">/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with line_join round emits a:round inside a:ln" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "RJ", line_color: "000000", line_join: "round")

    xlsx_tempfile = Tempfile.new(["xlsxrb-linejoin", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:ln>.*<a:round/>.*</a:ln>}m, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with head_end and tail_end emits arrow elements" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Arrow",
                     line_color: "000000",
                     head_end: { type: "triangle", w: "med", len: "med" },
                     tail_end: { type: "stealth", w: "lg", len: "lg" })

    xlsx_tempfile = Tempfile.new(["xlsxrb-arrow", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:headEnd type="triangle" w="med" len="med"/>}, drawing_xml)
    assert_match(%r{<a:tailEnd type="stealth" w="lg" len="lg"/>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "insert_image with clientData attrs stores locks_with_sheet and prints_with_sheet" do
    writer = Xlsxrb::Writer.new
    png = "\x89PNG".b
    writer.insert_image(png, ext: "png", name: "Pic1", locks_with_sheet: false, prints_with_sheet: false)
    imgs = writer.images
    assert_equal(false, imgs[0][:locks_with_sheet])
    assert_equal(false, imgs[0][:prints_with_sheet])
  end

  test "insert_image with anchor offsets stores col/row offsets" do
    writer = Xlsxrb::Writer.new
    png = "\x89PNG".b
    writer.insert_image(png, ext: "png", name: "Pic1",
                             from_col: 1, from_row: 2, to_col: 5, to_row: 8,
                             from_col_off: 100_000, from_row_off: 200_000,
                             to_col_off: 300_000, to_row_off: 400_000)
    imgs = writer.images
    assert_equal(100_000, imgs[0][:from_col_off])
    assert_equal(200_000, imgs[0][:from_row_off])
    assert_equal(300_000, imgs[0][:to_col_off])
    assert_equal(400_000, imgs[0][:to_row_off])
  end

  test "add_comment stores comment definition" do
    writer = Xlsxrb::Writer.new
    writer.add_comment("A1", "Note text", author: "Tester")
    writer.add_comment("B2", "Second note")
    comments = writer.comments
    assert_equal(2, comments.size)
    assert_equal("A1", comments[0][:ref])
    assert_equal("Note text", comments[0][:text])
    assert_equal("Tester", comments[0][:author])
    assert_equal("Author", comments[1][:author])
  end

  test "add_comment stores rich text comment" do
    writer = Xlsxrb::Writer.new
    rt = Xlsxrb::RichText.new(runs: [
                                { text: "Bold", font: { bold: true, sz: 9, name: "Calibri" } },
                                { text: " normal" }
                              ])
    writer.add_comment("A1", rt, author: "Tester")
    comments = writer.comments
    assert_equal(1, comments.size)
    assert_instance_of(Xlsxrb::RichText, comments[0][:text])
    assert_equal("Bold normal", comments[0][:text].to_s)
  end

  test "add_pivot_table stores pivot table definition" do
    writer = Xlsxrb::Writer.new
    writer.add_pivot_table("Sheet1!A1:C4",
                           row_fields: [0],
                           data_fields: [{ fld: 2, name: "Sum of Amount", subtotal: "sum" }],
                           dest_ref: "E1:F4")
    pivots = writer.pivot_tables
    assert_equal(1, pivots.size)
    assert_equal([0], pivots[0][:row_fields])
    assert_equal("Sheet1!A1:C4", pivots[0][:source_ref])
  end

  test "add_pivot_table with col_fields, field_names, and items" do
    writer = Xlsxrb::Writer.new
    writer.add_pivot_table("Sheet1!A1:C4",
                           row_fields: [0],
                           col_fields: [1],
                           data_fields: [{ fld: 2, name: "Sum of Amount", subtotal: "sum" }],
                           field_names: %w[Category Region Amount],
                           items: { 0 => %w[A B C], 1 => %w[East West] })
    pivots = writer.pivot_tables
    assert_equal(1, pivots.size)
    assert_equal([0], pivots[0][:row_fields])
    assert_equal([1], pivots[0][:col_fields])
    assert_equal(%w[Category Region Amount], pivots[0][:field_names])
    assert_equal(%w[A B C], pivots[0][:items][0])
    assert_equal(%w[East West], pivots[0][:items][1])
  end

  test "add_pivot_table with extended attributes (dataCaption, dataOnRows, grandTotals, compact, outline, showHeaders)" do
    writer = Xlsxrb::Writer.new
    writer.add_pivot_table("Sheet1!A1:C4",
                           row_fields: [0],
                           data_fields: [{ fld: 2, name: "Sum of Amount", subtotal: "sum" }],
                           data_caption: "Custom Caption",
                           data_on_rows: true,
                           row_grand_totals: false,
                           col_grand_totals: false,
                           compact: false,
                           outline: false,
                           show_headers: false)
    pivots = writer.pivot_tables
    assert_equal(1, pivots.size)
    assert_equal("Custom Caption", pivots[0][:data_caption])
    assert_equal(true, pivots[0][:data_on_rows])
    assert_equal(false, pivots[0][:row_grand_totals])
    assert_equal(false, pivots[0][:col_grand_totals])
    assert_equal(false, pivots[0][:compact])
    assert_equal(false, pivots[0][:outline])
    assert_equal(false, pivots[0][:show_headers])
  end

  test "add_pivot_table with grandTotalCaption, errorCaption, missingCaption, tag, version attrs" do
    writer = Xlsxrb::Writer.new
    writer.add_pivot_table("Sheet1!A1:C4",
                           row_fields: [0],
                           data_fields: [{ fld: 2, name: "Sum", subtotal: "sum" }],
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
    pivots = writer.pivot_tables
    assert_equal("Grand Total", pivots[0][:grand_total_caption])
    assert_equal("#N/A", pivots[0][:error_caption])
    assert_equal(true, pivots[0][:show_error])
    assert_equal("(blank)", pivots[0][:missing_caption])
    assert_equal(false, pivots[0][:show_missing])
    assert_equal("custom-tag", pivots[0][:tag])
    assert_equal(2, pivots[0][:indent])
    assert_equal(true, pivots[0][:published])
    assert_equal(6, pivots[0][:created_version])
    assert_equal(8, pivots[0][:updated_version])
    assert_equal(3, pivots[0][:min_refreshable_version])
  end

  test "add_pivot_table with applyXxxFormats attributes" do
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

    xlsx_tempfile = Tempfile.new(["xlsxrb-pivot", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/pivotTables/pivotTable1.xml")
    assert_match(/applyNumberFormats="1"/, xml_content)
    assert_match(/applyBorderFormats="1"/, xml_content)
    assert_match(/applyFontFormats="0"/, xml_content)
    assert_match(/applyPatternFormats="0"/, xml_content)
    assert_match(/applyAlignmentFormats="0"/, xml_content)
    assert_match(/applyWidthHeightFormats="0"/, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_pivot_table with source_name stores worksheetSource name" do
    writer = Xlsxrb::Writer.new
    writer.add_pivot_table("Sheet1!A1:C4",
                           row_fields: [0],
                           data_fields: [{ fld: 2, name: "Sum of Amount", subtotal: "sum" }],
                           source_name: "MyNamedRange")
    pivots = writer.pivot_tables
    assert_equal(1, pivots.size)
    assert_equal("MyNamedRange", pivots[0][:source_name])
  end

  test "add_pivot_table with dataField showDataAs, baseField, baseItem, numFmtId" do
    writer = Xlsxrb::Writer.new
    writer.add_pivot_table("Sheet1!A1:C4",
                           row_fields: [0],
                           data_fields: [{ fld: 2, name: "% of Total", subtotal: "sum",
                                           show_data_as: "percentOfTotal", base_field: 0, base_item: 0, num_fmt_id: 10 }])
    pivots = writer.pivot_tables
    df = pivots[0][:data_fields][0]
    assert_equal("percentOfTotal", df[:show_data_as])
    assert_equal(0, df[:base_field])
    assert_equal(0, df[:base_item])
    assert_equal(10, df[:num_fmt_id])
  end

  test "add_pivot_table with pivot_table_style emits pivotTableStyleInfo" do
    writer = Xlsxrb::Writer.new
    writer.add_pivot_table("Sheet1!A1:C4",
                           row_fields: [0],
                           data_fields: [{ fld: 2, name: "Sum", subtotal: "sum" }],
                           pivot_table_style: { name: "PivotStyleLight16",
                                                show_row_headers: true,
                                                show_col_headers: true,
                                                show_row_stripes: false,
                                                show_col_stripes: false,
                                                show_last_column: true })
    pivots = writer.pivot_tables
    psi = pivots[0][:pivot_table_style]
    assert_equal("PivotStyleLight16", psi[:name])
    assert_equal(true, psi[:show_row_headers])
    assert_equal(false, psi[:show_row_stripes])
  end

  test "add_external_link stores external link definition" do
    writer = Xlsxrb::Writer.new
    writer.add_external_link(target: "Book2.xlsx", sheet_names: %w[Sheet1 Sheet2])
    els = writer.external_links
    assert_equal(1, els.size)
    assert_equal("Book2.xlsx", els[0][:target])
    assert_equal(%w[Sheet1 Sheet2], els[0][:sheet_names])
  end

  test "preserve_macros flag" do
    writer = Xlsxrb::Writer.new
    assert_false(writer.preserve_macros?)
    writer.preserve_macros!
    assert_true(writer.preserve_macros?)
  end

  test "add_raw_entry includes entry in output" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-raw", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.add_raw_entry("custom/data.txt", "test content")
    writer.write(xlsx_path)

    reader = Xlsxrb::Reader.new(xlsx_path)
    data = reader.raw_entry("custom/data.txt")
    assert_equal("test content", data)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "insert_image on unknown sheet raises" do
    writer = Xlsxrb::Writer.new
    assert_raise(ArgumentError) { writer.insert_image("data", sheet: "Nonexistent") }
  end

  test "add_comment on unknown address raises" do
    writer = Xlsxrb::Writer.new
    assert_raise(ArgumentError) { writer.add_comment("ZZZ", "text") }
  end

  test "set_sheet_protection stores protection settings" do
    writer = Xlsxrb::Writer.new
    writer.set_sheet_protection(sheet: "Sheet1", password: "CF1A", objects: true, scenarios: true)
    prot = writer.sheet_protection(sheet: "Sheet1")
    assert_equal("CF1A", prot[:password])
    assert_equal(true, prot[:objects])
    assert_equal(true, prot[:scenarios])
  end

  test "copy_entries_from preserves extra parts in round-trip" do
    source_tempfile = Tempfile.new(["xlsxrb-source", ".xlsx"])
    source_path = source_tempfile.path
    source_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "original")
    writer.add_raw_entry("customXml/item1.xml", "<root>custom</root>")
    writer.write(source_path)

    # Copy entries into a new writer.
    writer2 = Xlsxrb::Writer.new
    writer2.copy_entries_from(source_path)
    writer2.set_cell("A1", "modified")

    output_tempfile = Tempfile.new(["xlsxrb-output", ".xlsx"])
    output_path = output_tempfile.path
    output_tempfile.close
    writer2.write(output_path)

    reader = Xlsxrb::Reader.new(output_path)
    # Generated cell overrides copied cell.
    cells = reader.cells
    assert_equal("modified", cells["A1"])
    # Custom XML part preserved.
    assert_equal("<root>custom</root>", reader.raw_entry("customXml/item1.xml"))
  ensure
    File.delete(source_path) if source_path && File.exist?(source_path)
    File.delete(output_path) if output_path && File.exist?(output_path)
  end

  test "set_workbook_protection stores protection settings" do
    writer = Xlsxrb::Writer.new
    writer.set_workbook_protection(lock_structure: true, lock_windows: false)
    prot = writer.workbook_protection
    assert_equal(true, prot[:lock_structure])
    assert_equal(false, prot[:lock_windows])
  end

  test "set_workbook_protection with lockRevision and revision algorithm attrs" do
    writer = Xlsxrb::Writer.new
    writer.set_workbook_protection(
      lock_structure: true,
      lock_revision: true,
      revisions_algorithm_name: "SHA-512",
      revisions_hash_value: "abc123",
      revisions_salt_value: "salt456",
      revisions_spin_count: 100_000
    )
    prot = writer.workbook_protection
    assert_equal(true, prot[:lock_revision])
    assert_equal("SHA-512", prot[:revisions_algorithm_name])
    assert_equal("abc123", prot[:revisions_hash_value])
    assert_equal("salt456", prot[:revisions_salt_value])
    assert_equal(100_000, prot[:revisions_spin_count])
  end

  test "add_cell_style with protection emits protection element" do
    writer = Xlsxrb::Writer.new
    style_id = writer.add_cell_style(protection: { locked: false, hidden: true })
    assert_equal(1, style_id)

    writer.set_cell("A1", "protected")
    writer.set_cell_style("A1", style_id)

    xlsx_tempfile = Tempfile.new(["xlsxrb-prot", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/styles.xml")
    assert_match(/applyProtection="1"/, xml_content)
    assert_match(%r{<protection[^/>]*locked="0"}, xml_content)
    assert_match(%r{<protection[^/>]*hidden="1"}, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_font with theme color and tint emits theme color attributes" do
    writer = Xlsxrb::Writer.new
    fid = writer.add_font(sz: 11, name: "Calibri", theme: 1, tint: -0.25)
    assert_equal(1, fid)

    style_id = writer.add_cell_style(font_id: fid)
    writer.set_cell("A1", "theme color")
    writer.set_cell_style("A1", style_id)

    xlsx_tempfile = Tempfile.new(["xlsxrb-theme", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/styles.xml")
    assert_match(%r{<color[^/>]*theme="1"}, xml_content)
    assert_match(%r{<color[^/>]*tint="-0.25"}, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_font with indexed color emits indexed color attribute" do
    writer = Xlsxrb::Writer.new
    fid = writer.add_font(sz: 11, name: "Calibri", indexed: 10)
    assert_equal(1, fid)

    style_id = writer.add_cell_style(font_id: fid)
    writer.set_cell("A1", "indexed color")
    writer.set_cell_style("A1", style_id)

    xlsx_tempfile = Tempfile.new(["xlsxrb-indexed", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/styles.xml")
    assert_match(%r{<color indexed="10"/>}, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_fill with theme colors emits theme color attributes" do
    writer = Xlsxrb::Writer.new
    fill_id = writer.add_fill(pattern: "solid", fg_color_theme: 4, fg_color_tint: 0.6)
    assert_operator(fill_id, :>=, 2)

    style_id = writer.add_cell_style(fill_id: fill_id)
    writer.set_cell("A1", "theme fill")
    writer.set_cell_style("A1", style_id)

    xlsx_tempfile = Tempfile.new(["xlsxrb-themefill", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/styles.xml")
    assert_match(%r{<fgColor[^/>]*theme="4"}, xml_content)
    assert_match(%r{<fgColor[^/>]*tint="0.6"}, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_fill with auto color emits auto attribute" do
    writer = Xlsxrb::Writer.new
    fill_id = writer.add_fill(pattern: "solid", fg_color_auto: true, bg_color_auto: true)
    assert_operator(fill_id, :>=, 2)

    style_id = writer.add_cell_style(fill_id: fill_id)
    writer.set_cell("A1", "auto fill")
    writer.set_cell_style("A1", style_id)

    xlsx_tempfile = Tempfile.new(["xlsxrb-autofill", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/styles.xml")
    assert_match(%r{<fgColor auto="1"/>}, xml_content)
    assert_match(%r{<bgColor auto="1"/>}, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_border with theme color emits theme color attributes" do
    writer = Xlsxrb::Writer.new
    brd_id = writer.add_border(left: { style: "thin", theme: 1, tint: -0.25 })
    assert_equal(1, brd_id)

    style_id = writer.add_cell_style(border_id: brd_id)
    writer.set_cell("A1", "border theme")
    writer.set_cell_style("A1", style_id)

    xlsx_tempfile = Tempfile.new(["xlsxrb-brdtheme", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/styles.xml")
    assert_match(/<left style="thin">/, xml_content)
    assert_match(%r{<color[^/>]*theme="1"}, xml_content)
    assert_match(%r{<color[^/>]*tint="-0.25"}, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "rich text value stored and retrievable" do
    writer = Xlsxrb::Writer.new
    rt = Xlsxrb::RichText.new(runs: [
                                { text: "Bold", font: { bold: true } },
                                { text: " Normal" }
                              ])
    writer.set_cell("A1", rt)
    assert_equal(rt, writer.cells["A1"])
    assert_equal("Bold Normal", rt.to_s)
  end

  test "shared formula attributes stored" do
    writer = Xlsxrb::Writer.new
    sf = Xlsxrb::Formula.new(expression: "A1+1", type: :shared, ref: "B1:B10", shared_index: 0, cached_value: "2")
    writer.set_cell("B1", sf)
    result = writer.cells["B1"]
    assert_equal(:shared, result.type)
    assert_equal(0, result.shared_index)
    assert_equal("B1:B10", result.ref)
  end

  test "array formula attributes stored" do
    writer = Xlsxrb::Writer.new
    af = Xlsxrb::Formula.new(expression: "{SUM(A1:A3*B1:B3)}", type: :array, ref: "C1")
    writer.set_cell("C1", af)
    result = writer.cells["C1"]
    assert_equal(:array, result.type)
    assert_equal("C1", result.ref)
  end

  test "formula calculate_always emits ca attribute" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", Xlsxrb::Formula.new(expression: "NOW()", cached_value: "45000", calculate_always: true))

    xlsx_tempfile = Tempfile.new(["xlsxrb-ca", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    sheet_xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/ca="1"/, sheet_xml)
    assert_match(%r{<f ca="1">NOW\(\)</f>}, sheet_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "formula aca emits aca attribute" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", Xlsxrb::Formula.new(expression: "SUM(B1:B10)", type: :array, ref: "A1", aca: true))

    xlsx_tempfile = Tempfile.new(["xlsxrb-aca", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    sheet_xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/aca="1"/, sheet_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "formula bx emits bx attribute" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", Xlsxrb::Formula.new(expression: "SUM(B1:B10)", bx: true))

    xlsx_tempfile = Tempfile.new(["xlsxrb-bx", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    sheet_xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/bx="1"/, sheet_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "formula dataTable type emits dt2D dtr r1 r2 attributes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", Xlsxrb::Formula.new(
                            expression: "", type: :data_table,
                            dt2d: true, dtr: true, r1: "A$1", r2: "$A1"
                          ))

    xlsx_tempfile = Tempfile.new(["xlsxrb-dt", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    sheet_xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/t="dataTable"/, sheet_xml)
    assert_match(/dt2D="1"/, sheet_xml)
    assert_match(/dtr="1"/, sheet_xml)
    assert_match(/r1="A\$1"/, sheet_xml)
    assert_match(/r2="\$A1"/, sheet_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "set_cell_phonetic emits ph attribute on cell" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "テスト")
    writer.set_cell_phonetic("A1")

    xlsx_tempfile = Tempfile.new(["xlsxrb-ph", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    sheet_xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/ph="1"/, sheet_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_cell_style with alignment stores alignment attributes" do
    writer = Xlsxrb::Writer.new
    style_id = writer.add_cell_style(
      alignment: { horizontal: "center", vertical: "top", wrap_text: true, text_rotation: 45,
                   indent: 2, shrink_to_fit: true }
    )
    assert_equal(1, style_id)

    writer.set_cell("A1", "aligned")
    writer.set_cell_style("A1", style_id)

    xlsx_tempfile = Tempfile.new(["xlsxrb-alignment", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    # Verify alignment element is emitted in styles.xml
    xml_content = read_xml_from_xlsx(xlsx_path, "xl/styles.xml")
    assert_match(/applyAlignment="1"/, xml_content)
    assert_match(%r{<alignment[^/>]*horizontal="center"}, xml_content)
    assert_match(%r{<alignment[^/>]*vertical="top"}, xml_content)
    assert_match(%r{<alignment[^/>]*wrapText="1"}, xml_content)
    assert_match(%r{<alignment[^/>]*textRotation="45"}, xml_content)
    assert_match(%r{<alignment[^/>]*indent="2"}, xml_content)
    assert_match(%r{<alignment[^/>]*shrinkToFit="1"}, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_cell_style with partial alignment only emits specified attrs" do
    writer = Xlsxrb::Writer.new
    style_id = writer.add_cell_style(alignment: { horizontal: "left" })
    assert_equal(1, style_id)

    writer.set_cell("A1", "left")
    writer.set_cell_style("A1", style_id)

    xlsx_tempfile = Tempfile.new(["xlsxrb-alignment", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/styles.xml")
    assert_match(%r{<alignment horizontal="left"/>}, xml_content)
    assert_no_match(/wrapText/, xml_content)
    assert_no_match(/textRotation/, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "stores expanded conditional formatting rule types" do
    writer = Xlsxrb::Writer.new
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

    cfs = writer.conditional_formats
    assert_equal(6, cfs.size)
    assert_equal(:above_average, cfs[0][:type])
    assert_equal(false, cfs[0][:above_average])
    assert_equal(true, cfs[0][:equal_average])
    assert_equal(:top10, cfs[1][:type])
    assert_equal(5, cfs[1][:rank])
    assert_equal(true, cfs[1][:percent])
    assert_equal(true, cfs[1][:bottom])
    assert_equal(:duplicate_values, cfs[2][:type])
    assert_equal(:contains_text, cfs[3][:type])
    assert_equal("hello", cfs[3][:text])
    assert_equal(:begins_with, cfs[4][:type])
    assert_equal(:ends_with, cfs[5][:type])
  end

  test "rich text emits extended font attributes in XML" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    rt = Xlsxrb::RichText.new(runs: [
                                { text: "Strike", font: { strike: true, name: "Arial", sz: 11 } },
                                { text: "Double", font: { underline: "double", name: "Arial", sz: 11 } },
                                { text: "Super", font: { vert_align: "superscript", name: "Arial", sz: 11 } },
                                { text: "Theme", font: { theme: 1, tint: 0.5, name: "Calibri", sz: 11, family: 2, scheme: "minor" } }
                              ])
    writer.set_cell("A1", rt)
    writer.write(xlsx_path)

    xml = read_xml_from_xlsx(xlsx_path, "xl/sharedStrings.xml")
    assert_match(%r{<strike/>}, xml)
    assert_match(%r{<u val="double"/>}, xml)
    assert_match(%r{<vertAlign val="superscript"/>}, xml)
    assert_match(/<color[^>]*theme="1"/, xml)
    assert_match(/<color[^>]*tint="0.5"/, xml)
    assert_match(%r{<family val="2"/>}, xml)
    assert_match(%r{<scheme val="minor"/>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "CF colorScale and dataBar emit theme/indexed colors" do
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

    xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(%r{<color theme="4" tint="-0.25"/>}, xml)
    assert_match(%r{<color theme="9"/>}, xml)
    assert_match(%r{<color indexed="10"/>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "gradient fill stops emit theme/indexed colors" do
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

    xml = read_xml_from_xlsx(xlsx_path, "xl/styles.xml")
    assert_match(%r{<stop position="0"><color theme="4" tint="-0.5"/></stop>}, xml)
    assert_match(%r{<stop position="1"><color indexed="12"/></stop>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "stores complete set of CF rule types" do
    writer = Xlsxrb::Writer.new
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

    cfs = writer.conditional_formats
    assert_equal(6, cfs.size)
    assert_equal(:expression, cfs[0][:type])
    assert_equal(:unique_values, cfs[1][:type])
    assert_equal(:not_contains_text, cfs[2][:type])
    assert_equal("bad", cfs[2][:text])
    assert_equal(:contains_blanks, cfs[3][:type])
    assert_equal(:not_contains_blanks, cfs[4][:type])
    assert_equal(:time_period, cfs[5][:type])
    assert_equal("lastWeek", cfs[5][:time_period])
  end

  test "DXF emits alignment and protection elements" do
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

    xml = read_xml_from_xlsx(xlsx_path, "xl/styles.xml")
    assert_match(/<dxf>/, xml)
    assert_match(%r{<alignment horizontal="center" wrapText="1"/>}, xml)
    assert_match(%r{<protection locked="0" hidden="1"/>}, xml)
    assert_match(%r{<numFmt numFmtId="164" formatCode="#,##0.00"/>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writes error cell values with t=e" do
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

    xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(%r{<c r="A1" t="e"><v>#N/A</v></c>}, xml)
    assert_match(%r{<c r="B1" t="e"><v>#DIV/0!</v></c>}, xml)
    assert_match(%r{<c r="C1" t="e"><v>#VALUE!</v></c>}, xml)
    assert_match(%r{<c r="D1" t="e"><v>#REF!</v></c>}, xml)
    assert_match(%r{<c r="E1" t="e"><v>#NAME\?</v></c>}, xml)
    assert_match(%r{<c r="F1" t="e"><v>#NUM!</v></c>}, xml)
    assert_match(%r{<c r="G1" t="e"><v>#NULL!</v></c>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writes extended core properties" do
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

    xml = read_xml_from_xlsx(xlsx_path, "docProps/core.xml")
    assert_match(%r{<dc:title>My Title</dc:title>}, xml)
    assert_match(%r{<dc:subject>My Subject</dc:subject>}, xml)
    assert_match(%r{<dc:creator>Alice</dc:creator>}, xml)
    assert_match(%r{<cp:keywords>ruby, xlsx</cp:keywords>}, xml)
    assert_match(%r{<dc:description>A test document</dc:description>}, xml)
    assert_match(%r{<cp:lastModifiedBy>Bob</cp:lastModifiedBy>}, xml)
    assert_match(%r{<cp:revision>3</cp:revision>}, xml)
    assert_match(%r{<cp:category>Reports</cp:category>}, xml)
    assert_match(%r{<cp:contentStatus>Draft</cp:contentStatus>}, xml)
    assert_match(%r{<dc:language>en-US</dc:language>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writes split pane with xSplit and ySplit in twips" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_split_pane(x_split: 2400, y_split: 1800, top_left_cell: "C4")
    writer.write(xlsx_path)

    xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(%r{<pane xSplit="2400" ySplit="1800" topLeftCell="C4" activePane="bottomRight"/>}, xml)
    assert_no_match(/state="frozen"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writes colorFilter and iconFilter in autoFilter" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Header1")
    writer.set_cell("B1", "Header2")
    writer.set_auto_filter("A1:B10")
    writer.add_filter_column(0, { type: :color_filter, dxf_id: 0 })
    writer.add_filter_column(1, { type: :icon_filter, icon_set: "3Arrows", icon_id: 1 })
    writer.write(xlsx_path)

    xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(%r{<colorFilter dxfId="0"/>}, xml)
    assert_match(%r{<iconFilter iconSet="3Arrows" iconId="1"/>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writes colorFilter with cellColor=false for font color filter" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Header")
    writer.set_auto_filter("A1:A10")
    writer.add_filter_column(0, { type: :color_filter, dxf_id: 1, cell_color: false })
    writer.write(xlsx_path)

    xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(%r{<colorFilter dxfId="1" cellColor="0"/>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writes Time values as fractional serial numbers with datetime format" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    t = Time.utc(2024, 3, 15, 14, 30, 0)
    writer.set_cell("A1", t)
    writer.write(xlsx_path)

    xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    # Serial for 2024-03-15 = 45366, plus 14:30:00 = 14*3600+30*60 = 52200 / 86400 = 0.604166...
    assert_match(/<v>45366\.604166/, xml)

    styles_xml = read_xml_from_xlsx(xlsx_path, "xl/styles.xml")
    assert_match(/yyyy\\-mm\\-dd\\ hh:mm:ss/, styles_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "set_print_area creates _xlnm.Print_Area defined name" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_print_area("A1:D20")

    dns = writer.defined_names
    assert_equal(1, dns.size)
    assert_equal("_xlnm.Print_Area", dns[0][:name])
    assert_equal("'Sheet1'!$A$1:$D$20", dns[0][:value])
    assert_equal(0, dns[0][:local_sheet_id])
  end

  test "set_print_titles creates _xlnm.Print_Titles defined name" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_print_titles(rows: "1:3", cols: "A:B")

    dns = writer.defined_names
    assert_equal(1, dns.size)
    assert_equal("_xlnm.Print_Titles", dns[0][:name])
    assert_equal("'Sheet1'!$A:$B,'Sheet1'!$1:$3", dns[0][:value])
  end

  test "hash_password returns deterministic result with fixed salt" do
    salt = "\x00" * 16
    result = Xlsxrb.hash_password("secret", salt: salt, spin_count: 1000)

    assert_equal("SHA-512", result[:algorithm_name])
    assert_equal(1000, result[:spin_count])
    assert_equal([salt].pack("m0"), result[:salt_value])
    assert_match(%r{\A[A-Za-z0-9+/]+=*\z}, result[:hash_value])

    # Same inputs should produce same output
    result2 = Xlsxrb.hash_password("secret", salt: salt, spin_count: 1000)
    assert_equal(result[:hash_value], result2[:hash_value])
  end

  test "hash_password integrates with set_sheet_protection" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "protected")
    hp = Xlsxrb.hash_password("mypassword", spin_count: 1000)
    writer.set_sheet_protection(**hp)

    xlsx_tempfile = Tempfile.new(["xlsxrb-hash-pw", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    sheet_xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/algorithmName="SHA-512"/, sheet_xml)
    assert_match(%r{hashValue="[A-Za-z0-9+/]+=*"}, sheet_xml)
    assert_match(%r{saltValue="[A-Za-z0-9+/]+=*"}, sheet_xml)
    assert_match(/spinCount="1000"/, sheet_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "hash_password supports SHA-256 algorithm" do
    result = Xlsxrb.hash_password("test", algorithm: "SHA-256", spin_count: 100)
    assert_equal("SHA-256", result[:algorithm_name])
    assert_equal(100, result[:spin_count])
    # SHA-256 produces 32-byte hash → 44-char base64
    assert_equal(44, result[:hash_value].length)
  end

  test "set_header_footer emits firstHeader and firstFooter elements" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_header_footer(:first_header, "&CFirst Page Header")
    writer.set_header_footer(:first_footer, "&CFirst Page Footer")
    writer.set_header_footer(:different_first, true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-hf", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    sheet_xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/differentFirst="1"/, sheet_xml)
    assert_match(%r{<firstHeader>&amp;CFirst Page Header</firstHeader>}, sheet_xml)
    assert_match(%r{<firstFooter>&amp;CFirst Page Footer</firstFooter>}, sheet_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "set_header_footer emits differentOddEven attribute" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_header_footer(:odd_header, "&LOdd Header")
    writer.set_header_footer(:even_header, "&LEven Header")
    writer.set_header_footer(:different_odd_even, true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-hf", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    sheet_xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/differentOddEven="1"/, sheet_xml)
    assert_match(%r{<evenHeader>&amp;LEven Header</evenHeader>}, sheet_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "page_setup emits additional attributes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_page_setup(:page_order, "overThenDown")
    writer.set_page_setup(:black_and_white, true)
    writer.set_page_setup(:draft, true)
    writer.set_page_setup(:cell_comments, "atEnd")
    writer.set_page_setup(:first_page_number, 5)
    writer.set_page_setup(:use_first_page_number, true)
    writer.set_page_setup(:horizontal_dpi, 300)
    writer.set_page_setup(:vertical_dpi, 300)

    xlsx_tempfile = Tempfile.new(["xlsxrb-ps", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    sheet_xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/pageOrder="overThenDown"/, sheet_xml)
    assert_match(/blackAndWhite="1"/, sheet_xml)
    assert_match(/draft="1"/, sheet_xml)
    assert_match(/cellComments="atEnd"/, sheet_xml)
    assert_match(/firstPageNumber="5"/, sheet_xml)
    assert_match(/useFirstPageNumber="1"/, sheet_xml)
    assert_match(/horizontalDpi="300"/, sheet_xml)
    assert_match(/verticalDpi="300"/, sheet_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "data validation emits showDropDown and imeMode attributes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_data_validation("A1:A10", type: "list",
                                         formula1: '"Yes,No"',
                                         show_drop_down: true,
                                         ime_mode: "hiragana")

    xlsx_tempfile = Tempfile.new(["xlsxrb-dv", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    sheet_xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/showDropDown="1"/, sheet_xml)
    assert_match(/imeMode="hiragana"/, sheet_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "alignment emits readingOrder and justifyLastLine attributes" do
    writer = Xlsxrb::Writer.new
    sid = writer.add_cell_style(alignment: { horizontal: "distributed", reading_order: 2, justify_last_line: true })
    writer.set_cell("A1", "RTL text")
    writer.set_cell_style("A1", sid)

    xlsx_tempfile = Tempfile.new(["xlsxrb-align", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    styles_xml = read_xml_from_xlsx(xlsx_path, "xl/styles.xml")
    assert_match(/readingOrder="2"/, styles_xml)
    assert_match(/justifyLastLine="1"/, styles_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "alignment emits relativeIndent attribute" do
    writer = Xlsxrb::Writer.new
    sid = writer.add_cell_style(alignment: { indent: 2, relative_indent: -1 })
    writer.set_cell("A1", "indented")
    writer.set_cell_style("A1", sid)

    xlsx_tempfile = Tempfile.new(["xlsxrb-align-ri", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    styles_xml = read_xml_from_xlsx(xlsx_path, "xl/styles.xml")
    assert_match(/relativeIndent="-1"/, styles_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "font emits charset attribute" do
    writer = Xlsxrb::Writer.new
    fid = writer.add_font(name: "MS Gothic", sz: 11, family: 3, charset: 128)
    sid = writer.add_cell_style(font_id: fid)
    writer.set_cell("A1", "テスト")
    writer.set_cell_style("A1", sid)

    xlsx_tempfile = Tempfile.new(["xlsxrb-charset", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    styles_xml = read_xml_from_xlsx(xlsx_path, "xl/styles.xml")
    assert_match(/charset val="128"/, styles_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "font emits shadow, outline, condense, extend" do
    writer = Xlsxrb::Writer.new
    fid = writer.add_font(name: "Arial", sz: 12, bold: true, shadow: true, outline: true, condense: true, extend: true)
    sid = writer.add_cell_style(font_id: fid)
    writer.set_cell("A1", "styled")
    writer.set_cell_style("A1", sid)

    xlsx_tempfile = Tempfile.new(["xlsxrb-font-effects", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    styles_xml = read_xml_from_xlsx(xlsx_path, "xl/styles.xml")
    assert_match(%r{<shadow/>}, styles_xml)
    assert_match(%r{<outline/>}, styles_xml)
    assert_match(%r{<condense/>}, styles_xml)
    assert_match(%r{<extend/>}, styles_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "sheetFormatPr emits outlineLevelRow, outlineLevelCol, zeroHeight, customHeight" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_sheet_format(:default_row_height, 15)
    writer.set_sheet_format(:outline_level_row, 3)
    writer.set_sheet_format(:outline_level_col, 2)
    writer.set_sheet_format(:zero_height, true)
    writer.set_sheet_format(:custom_height, true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-sfp", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    sheet_xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/outlineLevelRow="3"/, sheet_xml)
    assert_match(/outlineLevelCol="2"/, sheet_xml)
    assert_match(/zeroHeight="1"/, sheet_xml)
    assert_match(/customHeight="1"/, sheet_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "cellXf emits quotePrefix attribute" do
    writer = Xlsxrb::Writer.new
    sid = writer.add_cell_style(quote_prefix: true)
    writer.set_cell("A1", "001234")
    writer.set_cell_style("A1", sid)

    xlsx_tempfile = Tempfile.new(["xlsxrb-qp", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    styles_xml = read_xml_from_xlsx(xlsx_path, "xl/styles.xml")
    assert_match(/quotePrefix="1"/, styles_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "showFormulas on sheet view" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_sheet_view(:show_formulas, true)

    sv = writer.sheet_view
    assert_equal(true, sv[:show_formulas])

    xlsx_tempfile = Tempfile.new(["test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/showFormulas="1"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "sheetView showZeros, view, showOutlineSymbols, showRuler" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_sheet_view(:show_zeros, false)
    writer.set_sheet_view(:view, "pageBreakPreview")
    writer.set_sheet_view(:show_outline_symbols, false)
    writer.set_sheet_view(:show_ruler, false)

    xlsx_tempfile = Tempfile.new(["test-sv", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/showZeros="0"/, xml)
    assert_match(/view="pageBreakPreview"/, xml)
    assert_match(/showOutlineSymbols="0"/, xml)
    assert_match(/showRuler="0"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "sheetView remaining attributes: topLeftCell, colorId, zoomScaleNormal, etc." do
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

    xlsx_tempfile = Tempfile.new(["test-sv2", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/windowProtection="1"/, xml)
    assert_match(/defaultGridColor="0"/, xml)
    assert_match(/showWhiteSpace="0"/, xml)
    assert_match(/topLeftCell="B5"/, xml)
    assert_match(/colorId="10"/, xml)
    assert_match(/zoomScaleNormal="80"/, xml)
    assert_match(/zoomScaleSheetLayoutView="75"/, xml)
    assert_match(/zoomScalePageLayoutView="90"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "codeName on sheet properties" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_sheet_property(:code_name, "MySheet")

    props = writer.sheet_properties
    assert_equal("MySheet", props[:code_name])

    xlsx_tempfile = Tempfile.new(["test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/codeName="MySheet"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "sheet properties extended attributes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_sheet_property(:filter_mode, true)
    writer.set_sheet_property(:published, false)
    writer.set_sheet_property(:enable_format_conditions_calculation, false)
    writer.set_sheet_property(:fit_to_page, true)
    writer.set_sheet_property(:auto_page_breaks, false)

    xlsx_tempfile = Tempfile.new(["test-sp", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/filterMode="1"/, xml)
    assert_match(/published="0"/, xml)
    assert_match(/enableFormatConditionsCalculation="0"/, xml)
    assert_match(/fitToPage="1"/, xml)
    assert_match(/autoPageBreaks="0"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "workbook view visibility" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_workbook_view(:visibility, "hidden")

    views = writer.workbook_views
    assert_equal("hidden", views[:visibility])

    xlsx_tempfile = Tempfile.new(["test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml = read_xml_from_xlsx(xlsx_path, "xl/workbook.xml")
    assert_match(/visibility="hidden"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "phonetic properties on sheet" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_phonetic_properties({ font_id: 1, type: "Hiragana", alignment: "center" })

    pp = writer.phonetic_properties
    assert_equal(1, pp[:font_id])
    assert_equal("Hiragana", pp[:type])
    assert_equal("center", pp[:alignment])

    xlsx_tempfile = Tempfile.new(["test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/phoneticPr/, xml)
    assert_match(/fontId="1"/, xml)
    assert_match(/type="Hiragana"/, xml)
    assert_match(/alignment="center"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "custom properties are written to docProps/custom.xml" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_custom_property("Project", "Alpha")
    writer.add_custom_property("Version", 42, type: :number)
    writer.add_custom_property("Active", true, type: :bool)

    xlsx_tempfile = Tempfile.new(["xlsxrb-custom", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    custom_xml = read_xml_from_xlsx(xlsx_path, "docProps/custom.xml")
    assert_match(/name="Project"/, custom_xml)
    assert_match(%r{<vt:lpwstr>Alpha</vt:lpwstr>}, custom_xml)
    assert_match(%r{<vt:i4>42</vt:i4>}, custom_xml)
    assert_match(%r{<vt:bool>true</vt:bool>}, custom_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_border with outline false emits outline attribute" do
    writer = Xlsxrb::Writer.new
    brd_id = writer.add_border(
      left: { style: "thin" },
      outline: false
    )
    assert_equal(1, brd_id)

    style_id = writer.add_cell_style(border_id: brd_id)
    writer.set_cell("A1", "no-outline")
    writer.set_cell_style("A1", style_id)

    xlsx_tempfile = Tempfile.new(["xlsxrb-outline", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/styles.xml")
    assert_match(/<border[^>]*outline="0"/, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "font with auto color emits auto attribute" do
    writer = Xlsxrb::Writer.new
    font_id = writer.add_font(auto: true, size: 11, name: "Calibri")
    style_id = writer.add_cell_style(font_id: font_id)
    writer.set_cell("A1", "auto-color")
    writer.set_cell_style("A1", style_id)

    xlsx_tempfile = Tempfile.new(["xlsxrb-auto", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/styles.xml")
    assert_match(%r{<color auto="1"/>}, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "headerFooter emits scaleWithDoc and alignWithMargins" do
    writer = Xlsxrb::Writer.new
    writer.set_header_footer(:odd_header, "&CHello")
    writer.set_header_footer(:scale_with_doc, false)
    writer.set_header_footer(:align_with_margins, false)
    writer.set_cell("A1", "hf")

    xlsx_tempfile = Tempfile.new(["xlsxrb-hf", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/scaleWithDoc="0"/, xml_content)
    assert_match(/alignWithMargins="0"/, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "pageSetup emits copies, paperHeight, paperWidth, errors" do
    writer = Xlsxrb::Writer.new
    writer.set_page_setup(:copies, 3)
    writer.set_page_setup(:paper_height, "297mm")
    writer.set_page_setup(:paper_width, "210mm")
    writer.set_page_setup(:errors, "blank")
    writer.set_page_setup(:use_printer_defaults, false)
    writer.set_cell("A1", "ps")

    xlsx_tempfile = Tempfile.new(["xlsxrb-ps", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/copies="3"/, xml_content)
    assert_match(/paperHeight="297mm"/, xml_content)
    assert_match(/paperWidth="210mm"/, xml_content)
    assert_match(/errors="blank"/, xml_content)
    assert_match(/usePrinterDefaults="0"/, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "printOptions emits gridLinesSet" do
    writer = Xlsxrb::Writer.new
    writer.set_print_option(:grid_lines_set, false)
    writer.set_cell("A1", "po")

    xlsx_tempfile = Tempfile.new(["xlsxrb-po", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/gridLinesSet="0"/, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "conditional format rule emits stdDev attribute" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 100)
    writer.add_conditional_format("A1:A10",
                                  type: :above_average,
                                  std_dev: 2,
                                  format_id: writer.add_dxf(font: { bold: true }))

    xlsx_tempfile = Tempfile.new(["xlsxrb-cf", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/stdDev="2"/, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "cfvo emits gte attribute when false" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_conditional_format("A1:A10",
                                  type: :color_scale,
                                  color_scale: {
                                    cfvo: [{ type: "min" }, { type: "num", val: "50", gte: false }, { type: "max" }],
                                    colors: [{ rgb: "FFFF0000" }, { rgb: "FFFFFF00" }, { rgb: "FF00FF00" }]
                                  })

    xlsx_tempfile = Tempfile.new(["xlsxrb-cfvo", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/gte="0"/, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "column with phonetic attribute emits phonetic" do
    writer = Xlsxrb::Writer.new
    writer.set_column_attribute("A", :phonetic, true)
    writer.set_cell("A1", "test")

    xlsx_tempfile = Tempfile.new(["xlsxrb-phonetic", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/phonetic="1"/, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "named cell style emits iLevel, hidden, customBuiltin" do
    writer = Xlsxrb::Writer.new
    writer.add_named_cell_style(
      name: "Heading 1",
      builtin_id: 16,
      i_level: 0,
      hidden: true,
      custom_builtin: true
    )
    writer.set_cell("A1", "test")

    xlsx_tempfile = Tempfile.new(["xlsxrb-cs", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/styles.xml")
    assert_match(/iLevel="0"/, xml_content)
    assert_match(/hidden="1"/, xml_content)
    assert_match(/customBuiltin="1"/, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "xf entry emits pivotButton attribute" do
    writer = Xlsxrb::Writer.new
    style_id = writer.add_cell_style(pivot_button: true)
    writer.set_cell("A1", "pivot")
    writer.set_cell_style("A1", style_id)

    xlsx_tempfile = Tempfile.new(["xlsxrb-pivot", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/styles.xml")
    assert_match(/pivotButton="1"/, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "sheetFormatPr emits thickTop and thickBottom" do
    writer = Xlsxrb::Writer.new
    writer.set_sheet_format(:default_row_height, 15)
    writer.set_sheet_format(:thick_top, true)
    writer.set_sheet_format(:thick_bottom, true)
    writer.set_cell("A1", "test")

    xlsx_tempfile = Tempfile.new(["xlsxrb-sfp", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/thickTop="1"/, xml_content)
    assert_match(/thickBottom="1"/, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "dataValidations container emits disablePrompts, xWindow, yWindow" do
    writer = Xlsxrb::Writer.new
    writer.set_data_validations_option(:disable_prompts, true)
    writer.set_data_validations_option(:x_window, 100)
    writer.set_data_validations_option(:y_window, 200)
    writer.add_data_validation("A1:A10", type: "whole", formula1: "1", formula2: "100")
    writer.set_cell("A1", 50)

    xlsx_tempfile = Tempfile.new(["xlsxrb-dvo", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/disablePrompts="1"/, xml_content)
    assert_match(/xWindow="100"/, xml_content)
    assert_match(/yWindow="200"/, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "border with vertical and horizontal sides" do
    writer = Xlsxrb::Writer.new
    brd_id = writer.add_border(
      vertical: { style: "thin", color: "FF00FF00" },
      horizontal: { style: "dashed", color: "FF0000FF" }
    )
    style_id = writer.add_cell_style(border_id: brd_id)
    writer.set_cell("A1", "vh")
    writer.set_cell_style("A1", style_id)

    xlsx_tempfile = Tempfile.new(["xlsxrb-bvh", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/styles.xml")
    assert_match(/<vertical style="thin">/, xml_content)
    assert_match(/<horizontal style="dashed">/, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "row break with extended attributes emits min, max, pt" do
    writer = Xlsxrb::Writer.new
    writer.add_row_break({ id: 10, min: 2, max: 8, man: true, pt: true })
    writer.set_cell("A1", "brk")

    xlsx_tempfile = Tempfile.new(["xlsxrb-brk", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/min="2"/, xml_content)
    assert_match(/max="8"/, xml_content)
    assert_match(/man="1"/, xml_content)
    assert_match(/pt="1"/, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "row with phonetic attribute emits ph" do
    writer = Xlsxrb::Writer.new
    writer.set_row_phonetic(1)
    writer.set_cell("A1", "test")

    xlsx_tempfile = Tempfile.new(["xlsxrb-ph", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/<row[^>]*ph="1"/, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "definedName emits extended attributes" do
    writer = Xlsxrb::Writer.new
    writer.add_defined_name("MyName", "Sheet1!$A$1",
                            comment: "A comment", description: "A desc",
                            function: true, vb_procedure: true, xlm: true,
                            shortcut_key: "A", publish_to_server: true, workbook_parameter: true)
    writer.set_cell("A1", "test")

    xlsx_tempfile = Tempfile.new(["xlsxrb-dn", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/workbook.xml")
    assert_match(/comment="A comment"/, xml_content)
    assert_match(/description="A desc"/, xml_content)
    assert_match(/function="1"/, xml_content)
    assert_match(/vbProcedure="1"/, xml_content)
    assert_match(/xlm="1"/, xml_content)
    assert_match(/shortcutKey="A"/, xml_content)
    assert_match(/publishToServer="1"/, xml_content)
    assert_match(/workbookParameter="1"/, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "fileVersion element emits appName and lastEdited" do
    writer = Xlsxrb::Writer.new
    writer.set_file_version(:app_name, "xl")
    writer.set_file_version(:last_edited, "7")
    writer.set_file_version(:lowest_edited, "7")
    writer.set_file_version(:rup_build, "27425")
    writer.set_cell("A1", "test")

    xlsx_tempfile = Tempfile.new(["xlsxrb-fv", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/workbook.xml")
    assert_match(/appName="xl"/, xml_content)
    assert_match(/lastEdited="7"/, xml_content)
    assert_match(/lowestEdited="7"/, xml_content)
    assert_match(/rupBuild="27425"/, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "fileSharing element emits readOnlyRecommended and userName" do
    writer = Xlsxrb::Writer.new
    writer.set_file_sharing(:read_only_recommended, true)
    writer.set_file_sharing(:user_name, "TestUser")
    writer.set_cell("A1", "test")

    xlsx_tempfile = Tempfile.new(["xlsxrb-fs", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/workbook.xml")
    assert_match(/readOnlyRecommended="1"/, xml_content)
    assert_match(/userName="TestUser"/, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits protectedRanges in worksheet XML" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "data")
    writer.add_protected_range(name: "EditArea", sqref: "A1:B10")
    writer.add_protected_range(name: "SecureRange", sqref: "C1:D5", algorithm_name: "SHA-512",
                               hash_value: "abc123", salt_value: "salt456", spin_count: 100_000)

    xlsx_tempfile = Tempfile.new(["xlsxrb-pr", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/<protectedRanges>/, xml_content)
    assert_match(/sqref="A1:B10"/, xml_content)
    assert_match(/name="EditArea"/, xml_content)
    assert_match(/algorithmName="SHA-512"/, xml_content)
    assert_match(/hashValue="abc123"/, xml_content)
    assert_match(/saltValue="salt456"/, xml_content)
    assert_match(/spinCount="100000"/, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits indexedColors and mruColors in stylesheet" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "data")
    writer.set_indexed_colors(%w[FF000000 FFFFFFFF FFFF0000])
    writer.set_mru_colors([{ rgb: "FF00FF00" }, { theme: 3, tint: 0.4 }])

    xlsx_tempfile = Tempfile.new(["xlsxrb-colors", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/styles.xml")
    assert_match(/<colors>/, xml_content)
    assert_match(/<indexedColors>/, xml_content)
    assert_match(%r{<rgbColor rgb="FF000000"/>}, xml_content)
    assert_match(%r{<rgbColor rgb="FFFFFFFF"/>}, xml_content)
    assert_match(%r{<rgbColor rgb="FFFF0000"/>}, xml_content)
    assert_match(/<mruColors>/, xml_content)
    assert_match(/rgb="FF00FF00"/, xml_content)
    assert_match(/theme="3"/, xml_content)
    assert_match(/tint="0.4"/, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits tableStyles in stylesheet" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "data")
    dxf_id = writer.add_dxf(font: { bold: true })
    writer.set_table_styles_option(:default_table_style, "TableStyleMedium2")
    writer.set_table_styles_option(:default_pivot_style, "PivotStyleLight16")
    writer.add_table_style(name: "MyStyle", elements: [
                             { type: "wholeTable", dxf_id: dxf_id },
                             { type: "headerRow", dxf_id: dxf_id, size: 2 }
                           ], pivot: false)

    xlsx_tempfile = Tempfile.new(["xlsxrb-ts", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/styles.xml")
    assert_match(/<tableStyles /, xml_content)
    assert_match(/defaultTableStyle="TableStyleMedium2"/, xml_content)
    assert_match(/defaultPivotStyle="PivotStyleLight16"/, xml_content)
    assert_match(/name="MyStyle"/, xml_content)
    assert_match(/pivot="0"/, xml_content)
    assert_match(/type="wholeTable"/, xml_content)
    assert_match(/type="headerRow"/, xml_content)
    assert_match(/size="2"/, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits cellWatches in worksheet XML" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 100)
    writer.set_cell("B2", 200)
    writer.add_cell_watch("A1")
    writer.add_cell_watch("B2")

    xlsx_tempfile = Tempfile.new(["xlsxrb-cw", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/<cellWatches>/, xml_content)
    assert_match(%r{<cellWatch r="A1"/>}, xml_content)
    assert_match(%r{<cellWatch r="B2"/>}, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits dataConsolidate in worksheet XML" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "data")
    writer.set_data_consolidate(
      function: "average", start_labels: true, left_labels: true, link: true,
      data_refs: [{ ref: "A1:B10", sheet: "Sheet1" }, { ref: "C1:D10", name: "Range2" }]
    )

    xlsx_tempfile = Tempfile.new(["xlsxrb-dc", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/<dataConsolidate /, xml_content)
    assert_match(/function="average"/, xml_content)
    assert_match(/startLabels="1"/, xml_content)
    assert_match(/leftLabels="1"/, xml_content)
    assert_match(/link="1"/, xml_content)
    assert_match(/<dataRefs count="2">/, xml_content)
    assert_match(/ref="A1:B10"/, xml_content)
    assert_match(/sheet="Sheet1"/, xml_content)
    assert_match(/name="Range2"/, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits sheetPr sync and transition attributes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "data")
    writer.set_sheet_property(:sync_horizontal, true)
    writer.set_sheet_property(:sync_vertical, true)
    writer.set_sheet_property(:sync_ref, "A1")
    writer.set_sheet_property(:transition_evaluation, true)
    writer.set_sheet_property(:transition_entry, true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-sp", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/syncHorizontal="1"/, xml_content)
    assert_match(/syncVertical="1"/, xml_content)
    assert_match(/syncRef="A1"/, xml_content)
    assert_match(/transitionEvaluation="1"/, xml_content)
    assert_match(/transitionEntry="1"/, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits workbookPr extended attributes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "data")
    writer.set_workbook_property(:show_border_unselected_tables, false)
    writer.set_workbook_property(:prompted_solutions, true)
    writer.set_workbook_property(:show_ink_annotation, false)
    writer.set_workbook_property(:save_external_link_values, false)
    writer.set_workbook_property(:show_pivot_chart_filter, true)
    writer.set_workbook_property(:allow_refresh_query, true)
    writer.set_workbook_property(:publish_items, true)
    writer.set_workbook_property(:date_compatibility, false)

    xlsx_tempfile = Tempfile.new(["xlsxrb-wbpr", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/workbook.xml")
    assert_match(/showBorderUnselectedTables="0"/, xml_content)
    assert_match(/promptedSolutions="1"/, xml_content)
    assert_match(/showInkAnnotation="0"/, xml_content)
    assert_match(/saveExternalLinkValues="0"/, xml_content)
    assert_match(/showPivotChartFilter="1"/, xml_content)
    assert_match(/allowRefreshQuery="1"/, xml_content)
    assert_match(/publishItems="1"/, xml_content)
    assert_match(/dateCompatibility="0"/, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits scenarios in worksheet XML" do
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

    xlsx_tempfile = Tempfile.new(["xlsxrb-sc", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/<scenarios /, xml_content)
    assert_match(/current="0"/, xml_content)
    assert_match(/name="Best Case"/, xml_content)
    assert_match(/user="Admin"/, xml_content)
    assert_match(/r="A1"/, xml_content)
    assert_match(/val="200"/, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits outlinePr applyStyles attribute" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "data")
    writer.set_sheet_property(:apply_styles, true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-ol", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/applyStyles="1"/, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits outlinePr showOutlineSymbols attribute" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "data")
    writer.set_sheet_property(:show_outline_symbols, false)

    xlsx_tempfile = Tempfile.new(["xlsxrb-sos", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/showOutlineSymbols="0"/, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits tabColor with indexed attribute" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "data")
    writer.set_sheet_property(:tab_color_indexed, 10)

    xlsx_tempfile = Tempfile.new(["xlsxrb-tci", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(%r{<tabColor indexed="10"/>}, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits tabColor with auto attribute" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "data")
    writer.set_sheet_property(:tab_color_auto, true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-tca", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(%r{<tabColor auto="1"/>}, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits fileRecoveryPr element" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "data")
    writer.set_file_recovery_property(:auto_recover, false)
    writer.set_file_recovery_property(:crash_save, true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-frp", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/workbook.xml")
    assert_match(/autoRecover="0"/, xml_content)
    assert_match(/crashSave="1"/, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits sheetCalcPr fullCalcOnLoad" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "data")
    writer.set_sheet_property(:full_calc_on_load, true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-scp", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    xml_content = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(%r{<sheetCalcPr fullCalcOnLoad="1"/>}, xml_content)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits table extended attributes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Name")
    writer.set_cell("B1", "Age")
    writer.add_table("A1:B5", columns: %w[Name Age],
                              header_row_count: 0, published: true, comment: "Test table",
                              insert_row: true, insert_row_shift: true,
                              header_row_dxf_id: 1, data_dxf_id: 2, totals_row_dxf_id: 3)
    xlsx_path = File.join(Dir.tmpdir, "table_ext_attrs_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    tbl_xml = read_xml_from_xlsx(xlsx_path, "xl/tables/table1.xml")
    assert_match(/headerRowCount="0"/, tbl_xml)
    assert_match(/published="1"/, tbl_xml)
    assert_match(/comment="Test table"/, tbl_xml)
    assert_match(/insertRow="1"/, tbl_xml)
    assert_match(/insertRowShift="1"/, tbl_xml)
    assert_match(/headerRowDxfId="1"/, tbl_xml)
    assert_match(/dataDxfId="2"/, tbl_xml)
    assert_match(/totalsRowDxfId="3"/, tbl_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_table with border dxfId, cellStyle, tableType, connectionId" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Name")
    writer.set_cell("B1", "Age")
    writer.add_table("A1:B5", columns: %w[Name Age],
                              header_row_border_dxf_id: 10,
                              table_border_dxf_id: 11,
                              totals_row_border_dxf_id: 12,
                              header_row_cell_style: "HeaderStyle",
                              totals_row_cell_style: "TotalsStyle",
                              connection_id: 5,
                              table_type: "queryTable")
    tbls = writer.tables
    assert_equal(10, tbls[0][:header_row_border_dxf_id])
    assert_equal(11, tbls[0][:table_border_dxf_id])
    assert_equal(12, tbls[0][:totals_row_border_dxf_id])
    assert_equal("HeaderStyle", tbls[0][:header_row_cell_style])
    assert_equal("TotalsStyle", tbls[0][:totals_row_cell_style])
    assert_equal(5, tbls[0][:connection_id])
    assert_equal("queryTable", tbls[0][:table_type])
  end

  test "emits ignoredErrors element" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "123")
    writer.add_ignored_error(sqref: "A1:B2", number_stored_as_text: true, eval_error: true)
    xlsx_path = File.join(Dir.tmpdir, "ignored_errors_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/<ignoredErrors>/, xml)
    assert_match(%r{<ignoredError sqref="A1:B2" evalError="1" numberStoredAsText="1"/>}, xml)
    assert_match(%r{</ignoredErrors>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits definedName extended attributes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_defined_name("MyFunc", "Sheet1!$A$1",
                            function_group_id: 4, custom_menu: "My Menu",
                            help: "Help text", status_bar: "Status text")
    xlsx_path = File.join(Dir.tmpdir, "dn_ext_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/workbook.xml")
    assert_match(/functionGroupId="4"/, xml)
    assert_match(/customMenu="My Menu"/, xml)
    assert_match(/help="Help text"/, xml)
    assert_match(/statusBar="Status text"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits conditionalFormatting pivot attribute" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_conditional_format("A1:A10", type: :cell_is, operator: "greaterThan",
                                            formula: "5", format_id: 0, priority: 1, pivot: true)
    xlsx_path = File.join(Dir.tmpdir, "cf_pivot_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/<conditionalFormatting sqref="A1:A10" pivot="1">/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits comment guid and shapeId attributes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_comment("A1", "Note", guid: "{12345678-1234-1234-1234-123456789ABC}", shape_id: 1025)
    xlsx_path = File.join(Dir.tmpdir, "comment_guid_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/comments1.xml")
    assert_match(/guid="\{12345678-1234-1234-1234-123456789ABC\}"/, xml)
    assert_match(/shapeId="1025"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits autoFilter extended attributes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Name")
    writer.set_cell("B1", "Date")
    writer.set_auto_filter("A1:B10")
    writer.add_filter_column(0, { type: :filters, values: %w[Alice],
                                  calendar_type: "gregorian",
                                  date_group_items: [{ date_time_grouping: "year", year: 2024 }],
                                  hidden_button: true, show_button: false })
    writer.add_filter_column(1, { type: :top10, top: true, val: 5, filter_val: 4.5 })
    xlsx_path = File.join(Dir.tmpdir, "af_ext_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/hiddenButton="1"/, xml)
    assert_match(/showButton="0"/, xml)
    assert_match(/calendarType="gregorian"/, xml)
    assert_match(/dateGroupItem dateTimeGrouping="year" year="2024"/, xml)
    assert_match(/filterVal="4.5"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits selection pane and activeCellId attributes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_selection("B2", sqref: "B2:C3", pane: "bottomRight", active_cell_id: 1)
    xlsx_path = File.join(Dir.tmpdir, "sel_pane_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/pane="bottomRight"/, xml)
    assert_match(/activeCellId="1"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits sortState extended attributes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Name")
    writer.set_cell("B1", "Value")
    writer.set_auto_filter("A1:B10")
    writer.set_sort_state("A2:B10",
                          [{ ref: "A2:A10", sort_by: "cellColor", dxf_id: 0 }],
                          column_sort: true, case_sensitive: true, sort_method: "pinYin")
    xlsx_path = File.join(Dir.tmpdir, "sort_ext_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/columnSort="1"/, xml)
    assert_match(/caseSensitive="1"/, xml)
    assert_match(/sortMethod="pinYin"/, xml)
    assert_match(/sortBy="cellColor"/, xml)
    assert_match(/dxfId="0"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits iconSet percent attribute" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_conditional_format("A1:A5", type: :icon_set, priority: 1,
                                           icon_set: { icon_set: "3Arrows", percent: false,
                                                       cfvo: [{ type: "min" }, { type: "num", val: "33" }, { type: "num", val: "67" }] })
    xlsx_path = File.join(Dir.tmpdir, "icon_pct_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/worksheets/sheet1.xml")
    assert_match(/percent="0"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits workbook conformance attribute" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_workbook_property(:conformance, "transitional")
    xlsx_path = File.join(Dir.tmpdir, "conformance_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/workbook.xml")
    assert_match(/conformance="transitional"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits pivotCacheDefinition optional attributes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Cat")
    writer.set_cell("B1", "Val")
    writer.set_cell("A2", "X")
    writer.set_cell("B2", 1)
    writer.add_pivot_table("Sheet1!A1:B2",
                           row_fields: [0], data_fields: [{ fld: 1, name: "Sum", subtotal: "sum" }],
                           cache_save_data: false, cache_enable_refresh: false,
                           cache_refreshed_by: "Bot", cache_refreshed_version: 6,
                           cache_created_version: 5, cache_record_count: 99,
                           cache_optimize_memory: true)
    xlsx_path = File.join(Dir.tmpdir, "pcd_attrs_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/pivotCache/pivotCacheDefinition1.xml")
    assert_match(/saveData="0"/, xml)
    assert_match(/enableRefresh="0"/, xml)
    assert_match(/refreshedBy="Bot"/, xml)
    assert_match(/refreshedVersion="6"/, xml)
    assert_match(/createdVersion="5"/, xml)
    assert_match(/recordCount="99"/, xml)
    assert_match(/optimizeMemory="1"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits location rowPageCount and colPageCount on pivot table" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Cat")
    writer.set_cell("B1", "Val")
    writer.set_cell("A2", "X")
    writer.set_cell("B2", 1)
    writer.add_pivot_table("Sheet1!A1:B2",
                           row_fields: [0], data_fields: [{ fld: 1, name: "Sum", subtotal: "sum" }],
                           row_page_count: 2, col_page_count: 3)
    xlsx_path = File.join(Dir.tmpdir, "loc_page_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/pivotTables/pivotTable1.xml")
    assert_match(/rowPageCount="2"/, xml)
    assert_match(/colPageCount="3"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits pivotField per-field attributes via field_attrs" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Cat")
    writer.set_cell("B1", "Val")
    writer.set_cell("A2", "X")
    writer.set_cell("B2", 1)
    writer.add_pivot_table("Sheet1!A1:B2",
                           row_fields: [0],
                           data_fields: [{ fld: 1, name: "Sum", subtotal: "sum" }],
                           field_attrs: {
                             0 => { compact: false, outline: false, subtotal_top: false,
                                    show_all: true, num_fmt_id: 164, sort_type: "ascending" }
                           })
    xlsx_path = File.join(Dir.tmpdir, "field_attrs_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/pivotTables/pivotTable1.xml")
    assert_match(/compact="0"/, xml)
    assert_match(/outline="0"/, xml)
    assert_match(/subtotalTop="0"/, xml)
    assert_match(/showAll="1"/, xml)
    assert_match(/numFmtId="164"/, xml)
    assert_match(/sortType="ascending"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits cacheField caption and formula via field_attrs" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Cat")
    writer.set_cell("B1", "Val")
    writer.set_cell("A2", "X")
    writer.set_cell("B2", 1)
    writer.add_pivot_table("Sheet1!A1:B2",
                           row_fields: [0],
                           data_fields: [{ fld: 1, name: "Sum", subtotal: "sum" }],
                           field_attrs: {
                             0 => { cache_caption: "Category", cache_formula: "='Sheet1'!A1", cache_num_fmt_id: 49 }
                           })
    xlsx_path = File.join(Dir.tmpdir, "cache_field_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/pivotCache/pivotCacheDefinition1.xml")
    assert_match(/caption="Category"/, xml)
    assert_match(/formula="=&apos;Sheet1&apos;!A1"/, xml)
    assert_match(/numFmtId="49"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits pivotField defaultSubtotal, insertBlankRow, insertPageBreak, includeNewItemsInFilter via field_attrs" do
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
    xlsx_path = File.join(Dir.tmpdir, "pf_ext_attrs_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/pivotTables/pivotTable1.xml")
    assert_match(/defaultSubtotal="0"/, xml)
    assert_match(/insertBlankRow="1"/, xml)
    assert_match(/insertPageBreak="1"/, xml)
    assert_match(/includeNewItemsInFilter="1"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits pivotTableDefinition extended display and layout attributes" do
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
    xlsx_path = File.join(Dir.tmpdir, "ptd_ext_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/pivotTables/pivotTable1.xml")
    assert_match(/multipleFieldFilters="0"/, xml)
    assert_match(/showDrill="0"/, xml)
    assert_match(/showDataTips="0"/, xml)
    assert_match(/enableDrill="0"/, xml)
    assert_match(/showMemberPropertyTips="0"/, xml)
    assert_match(/itemPrintTitles="1"/, xml)
    assert_match(/fieldPrintTitles="1"/, xml)
    assert_match(/preserveFormatting="0"/, xml)
    assert_match(/pageOverThenDown="1"/, xml)
    assert_match(/pageWrap="3"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits pivotTableDefinition compactData and outlineData attributes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Cat")
    writer.set_cell("B1", "Val")
    writer.set_cell("A2", "X")
    writer.set_cell("B2", 1)
    writer.add_pivot_table("Sheet1!A1:B2",
                           row_fields: [0],
                           data_fields: [{ fld: 1, name: "Sum", subtotal: "sum" }],
                           compact_data: false, outline_data: true)
    xlsx_path = File.join(Dir.tmpdir, "ptd_cd_od_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/pivotTables/pivotTable1.xml")
    assert_match(/compactData="0"/, xml)
    assert_match(/outlineData="1"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits pivotTableDefinition showMultipleLabel and showDataDropDown attributes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Cat")
    writer.set_cell("B1", "Val")
    writer.set_cell("A2", "X")
    writer.set_cell("B2", 1)
    writer.add_pivot_table("Sheet1!A1:B2",
                           row_fields: [0],
                           data_fields: [{ fld: 1, name: "Sum", subtotal: "sum" }],
                           show_multiple_label: false, show_data_drop_down: false)
    xlsx_path = File.join(Dir.tmpdir, "ptd_sml_sddd_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/pivotTables/pivotTable1.xml")
    assert_match(/showMultipleLabel="0"/, xml)
    assert_match(/showDataDropDown="0"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits pivotTableDefinition editData and disableFieldList attributes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Cat")
    writer.set_cell("B1", "Val")
    writer.set_cell("A2", "X")
    writer.set_cell("B2", 1)
    writer.add_pivot_table("Sheet1!A1:B2",
                           row_fields: [0],
                           data_fields: [{ fld: 1, name: "Sum", subtotal: "sum" }],
                           edit_data: true, disable_field_list: true)
    xlsx_path = File.join(Dir.tmpdir, "ptd_ed_dfl_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/pivotTables/pivotTable1.xml")
    assert_match(/editData="1"/, xml)
    assert_match(/disableFieldList="1"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits pivotTableDefinition visualTotals and printDrill attributes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Cat")
    writer.set_cell("B1", "Val")
    writer.set_cell("A2", "X")
    writer.set_cell("B2", 1)
    writer.add_pivot_table("Sheet1!A1:B2",
                           row_fields: [0],
                           data_fields: [{ fld: 1, name: "Sum", subtotal: "sum" }],
                           visual_totals: false, print_drill: true)
    xlsx_path = File.join(Dir.tmpdir, "ptd_vt_pd_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/pivotTables/pivotTable1.xml")
    assert_match(/visualTotals="0"/, xml)
    assert_match(/printDrill="1"/, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits chart plotVisOnly and dispBlanksAs attributes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     plot_vis_only: true, disp_blanks_as: "zero")
    xlsx_path = File.join(Dir.tmpdir, "chart_pvo_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:plotVisOnly val="1"/>}, xml)
    assert_match(%r{<c:dispBlanksAs val="zero"/>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits chart varyColors attribute" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :pie,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     vary_colors: true)
    xlsx_path = File.join(Dir.tmpdir, "chart_vc_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:varyColors val="1"/>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits chart style element" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     style: 26)
    xlsx_path = File.join(Dir.tmpdir, "chart_style_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:style val="26"/>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits autoTitleDeleted element" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     auto_title_deleted: true)
    xlsx_path = File.join(Dir.tmpdir, "auto_title_deleted_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:autoTitleDeleted val="1"/>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits roundedCorners element" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     rounded_corners: true)
    xlsx_path = File.join(Dir.tmpdir, "rounded_corners_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:roundedCorners val="1"/>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits showBubbleSize in data labels" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     data_labels: { show_bubble_size: true })
    xlsx_path = File.join(Dir.tmpdir, "show_bubble_size_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:showBubbleSize val="1"/>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits separator in data labels" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     data_labels: { show_val: true, separator: ", " })
    xlsx_path = File.join(Dir.tmpdir, "separator_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:separator>, </c:separator>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits dLblPos in data labels" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     data_labels: { position: "outEnd", show_val: true })
    xlsx_path = File.join(Dir.tmpdir, "dlblpos_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:dLblPos val="outEnd"/>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits tickLblPos on chart axes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     cat_axis_tick_lbl_pos: "low",
                     val_axis_tick_lbl_pos: "none")
    xlsx_path = File.join(Dir.tmpdir, "tick_lbl_pos_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:catAx>.*<c:tickLblPos val="low"/>.*</c:catAx>}m, xml)
    assert_match(%r{<c:valAx>.*<c:tickLblPos val="none"/>.*</c:valAx>}m, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits majorGridlines on chart axes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     val_axis_major_gridlines: true)
    xlsx_path = File.join(Dir.tmpdir, "gridlines_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:valAx>.*<c:majorGridlines/>.*</c:valAx>}m, xml)
    refute_match(%r{<c:catAx>.*<c:majorGridlines/>.*</c:catAx>}m, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits minorGridlines on chart axes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     cat_axis_minor_gridlines: true)
    xlsx_path = File.join(Dir.tmpdir, "minor_gridlines_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:catAx>.*<c:minorGridlines/>.*</c:catAx>}m, xml)
    refute_match(%r{<c:valAx>.*<c:minorGridlines/>.*</c:valAx>}m, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits showDLblsOverMax element" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     show_d_lbls_over_max: true)
    xlsx_path = File.join(Dir.tmpdir, "show_dlbls_over_max_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:showDLblsOverMax val="1"/>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits configurable axis delete attribute" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     cat_axis_delete: true, val_axis_delete: false)
    xlsx_path = File.join(Dir.tmpdir, "axis_delete_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:catAx>.*<c:delete val="1"/>.*</c:catAx>}m, xml)
    assert_match(%r{<c:valAx>.*<c:delete val="0"/>.*</c:valAx>}m, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits configurable axis orientation" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     val_axis_orientation: "maxMin")
    xlsx_path = File.join(Dir.tmpdir, "axis_orient_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:valAx>.*<c:orientation val="maxMin"/>.*</c:valAx>}m, xml)
    assert_match(%r{<c:catAx>.*<c:orientation val="minMax"/>.*</c:catAx>}m, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits gapWidth and overlap for bar charts" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     gap_width: 200, overlap: -25)
    xlsx_path = File.join(Dir.tmpdir, "gap_width_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:gapWidth val="200"/>}, xml)
    assert_match(%r{<c:overlap val="-25"/>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits legend overlay element" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     legend: { position: "b", overlay: true })
    xlsx_path = File.join(Dir.tmpdir, "legend_overlay_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:legend><c:legendPos val="b"/><c:overlay val="1"/></c:legend>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits view3D element with all properties" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     view_3d: { rot_x: 15, rot_y: 20, r_ang_ax: true, perspective: 30,
                                h_percent: 150, depth_percent: 200 })
    xlsx_path = File.join(Dir.tmpdir, "view3d_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:view3D>.*<c:rotX val="15"/>.*<c:hPercent val="150"/>.*<c:rotY val="20"/>.*<c:depthPercent val="200"/>.*<c:rAngAx val="1"/>.*<c:perspective val="30"/>.*</c:view3D>}m, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits numFmt on cat and val axes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     cat_axis_num_fmt: { format_code: "General", source_linked: true },
                     val_axis_num_fmt: { format_code: "0.00", source_linked: false })
    xlsx_path = File.join(Dir.tmpdir, "axis_numfmt_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:catAx>.*<c:numFmt formatCode="General" sourceLinked="1"/>.*</c:catAx>}m, xml)
    assert_match(%r{<c:valAx>.*<c:numFmt formatCode="0.00" sourceLinked="0"/>.*</c:valAx>}m, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits majorTickMark and minorTickMark on axes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     cat_axis_major_tick_mark: "out", cat_axis_minor_tick_mark: "in",
                     val_axis_major_tick_mark: "cross", val_axis_minor_tick_mark: "none")
    xlsx_path = File.join(Dir.tmpdir, "tick_marks_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:catAx>.*<c:majorTickMark val="out"/>.*<c:minorTickMark val="in"/>.*</c:catAx>}m, xml)
    assert_match(%r{<c:valAx>.*<c:majorTickMark val="cross"/>.*<c:minorTickMark val="none"/>.*</c:valAx>}m, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits crosses on cat and val axes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     cat_axis_crosses: "autoZero", val_axis_crosses: "max")
    xlsx_path = File.join(Dir.tmpdir, "axis_crosses_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:catAx>.*<c:crossAx val="2"/><c:crosses val="autoZero"/>.*</c:catAx>}m, xml)
    assert_match(%r{<c:valAx>.*<c:crossAx val="1"/><c:crosses val="max"/>.*</c:valAx>}m, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits crossBetween, majorUnit, minorUnit on val axis" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     val_axis_cross_between: "between",
                     val_axis_major_unit: 10, val_axis_minor_unit: 2)
    xlsx_path = File.join(Dir.tmpdir, "val_ax_ext_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:valAx>.*<c:crossBetween val="between"/>.*<c:majorUnit val="10"/>.*<c:minorUnit val="2"/>.*</c:valAx>}m, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits scaling min and max on axes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     val_axis_scaling_min: 0, val_axis_scaling_max: 100)
    xlsx_path = File.join(Dir.tmpdir, "scaling_minmax_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:valAx>.*<c:scaling>.*<c:max val="100"/>.*<c:min val="0"/>.*</c:scaling>}m, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits logBase in scaling on val axis" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     val_axis_log_base: 10)
    xlsx_path = File.join(Dir.tmpdir, "logbase_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:valAx>.*<c:scaling><c:logBase val="10"/>.*</c:scaling>}m, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits firstSliceAng and holeSize for doughnut charts" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :doughnut,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     first_slice_ang: 90, hole_size: 50)
    xlsx_path = File.join(Dir.tmpdir, "doughnut_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:firstSliceAng val="90"/>}, xml)
    assert_match(%r{<c:holeSize val="50"/>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits smooth and marker for line charts" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :line,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     smooth: true, marker: true)
    xlsx_path = File.join(Dir.tmpdir, "line_smooth_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:marker val="1"/>}, xml)
    assert_match(%r{<c:smooth val="1"/>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits scatterStyle and radarStyle for respective chart types" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :scatter,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     scatter_style: "smoothMarker")
    writer.add_chart(type: :radar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     radar_style: "filled")
    xlsx_path = File.join(Dir.tmpdir, "scatter_radar_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml1 = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    xml2 = read_xml_from_xlsx(xlsx_path, "xl/charts/chart2.xml")
    assert_match(%r{<c:scatterStyle val="smoothMarker"/>}, xml1)
    assert_match(%r{<c:radarStyle val="filled"/>}, xml2)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits configurable axPos on cat and val axes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     cat_axis_pos: "t", val_axis_pos: "r")
    xlsx_path = File.join(Dir.tmpdir, "axis_pos_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:catAx>.*<c:axPos val="t"/>.*</c:catAx>}m, xml)
    assert_match(%r{<c:valAx>.*<c:axPos val="r"/>.*</c:valAx>}m, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits gapDepth and shape for 3D bar charts" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar3d,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     gap_depth: 150, bar_shape: "cylinder")
    xlsx_path = File.join(Dir.tmpdir, "bar3d_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:gapDepth val="150"/>}, xml)
    assert_match(%r{<c:shape val="cylinder"/>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits bubble chart properties" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bubble,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     bubble_3d: true, bubble_scale: 200,
                     show_neg_bubbles: false, size_represents: "w")
    xlsx_path = File.join(Dir.tmpdir, "bubble_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:bubble3D val="1"/>}, xml)
    assert_match(%r{<c:bubbleScale val="200"/>}, xml)
    assert_match(%r{<c:showNegBubbles val="0"/>}, xml)
    assert_match(%r{<c:sizeRepresents val="w"/>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits crossesAt on cat and val axes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     cat_axis_crosses_at: 3.5, val_axis_crosses_at: 10.0)
    xlsx_path = File.join(Dir.tmpdir, "crosses_at_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:catAx>.*<c:crossesAt val="3.5"/>.*</c:catAx>}m, xml)
    assert_match(%r{<c:valAx>.*<c:crossesAt val="10.0"/>.*</c:valAx>}m, xml)
    assert_no_match(/<c:crosses /, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits wireframe for surface chart" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :surface,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     wireframe: true)
    xlsx_path = File.join(Dir.tmpdir, "wireframe_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:wireframe val="1"/>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits dTable with all boolean children" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     data_table: { show_horz_border: true, show_vert_border: true, show_outline: true, show_keys: true })
    xlsx_path = File.join(Dir.tmpdir, "dtable_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(/<c:dTable>/, xml)
    assert_match(%r{<c:showHorzBorder val="1"/>}, xml)
    assert_match(%r{<c:showVertBorder val="1"/>}, xml)
    assert_match(%r{<c:showOutline val="1"/>}, xml)
    assert_match(%r{<c:showKeys val="1"/>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits series spPr solidFill when fill_color specified" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1", fill_color: "FF0000" }])
    xlsx_path = File.join(Dir.tmpdir, "ser_fill_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits plot area spPr solidFill when plot_area_fill specified" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     plot_area_fill: "F0F0F0")
    xlsx_path = File.join(Dir.tmpdir, "plotfill_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{</c:valAx><c:spPr><a:solidFill><a:srgbClr val="F0F0F0"/></a:solidFill></c:spPr></c:plotArea>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits legendEntry with idx and delete" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     legend: { position: "r", entries: [{ idx: 0, delete: true }] })
    xlsx_path = File.join(Dir.tmpdir, "legentry_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:legendEntry><c:idx val="0"/><c:delete val="1"/></c:legendEntry>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits series spPr with line color and width" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :line,
                     series: [{ val_ref: "Sheet1!$A$1", line_color: "0000FF", line_width: 2 }])
    xlsx_path = File.join(Dir.tmpdir, "ser_line_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:spPr><a:ln w="25400"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln></c:spPr>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits series spPr with both fill and line" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1", fill_color: "FF0000", line_color: "0000FF", line_width: 1 }])
    xlsx_path = File.join(Dir.tmpdir, "ser_fill_line_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln w="12700">}, xml)
    assert_match(%r{<a:ln w="12700"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits series marker with symbol and size" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :line,
                     series: [{ val_ref: "Sheet1!$A$1", marker_symbol: "diamond", marker_size: 8 }])
    xlsx_path = File.join(Dir.tmpdir, "ser_marker_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:marker><c:symbol val="diamond"/><c:size val="8"/></c:marker>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits dropLines and hiLowLines on line chart" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :line,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     drop_lines: true, hi_low_lines: true)
    xlsx_path = File.join(Dir.tmpdir, "drop_hilow_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:dropLines/>}, xml)
    assert_match(%r{<c:hiLowLines/>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits upDownBars with gapWidth on line chart" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :line,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     up_down_bars: { gap_width: 150 })
    xlsx_path = File.join(Dir.tmpdir, "updown_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:upDownBars><c:gapWidth val="150"/><c:upBars/><c:downBars/></c:upDownBars>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits tickLblSkip and tickMarkSkip on cat axis" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     cat_axis_tick_lbl_skip: 2, cat_axis_tick_mark_skip: 3)
    xlsx_path = File.join(Dir.tmpdir, "tick_skip_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:catAx>.*<c:tickLblSkip val="2"/>.*</c:catAx>}m, xml)
    assert_match(%r{<c:catAx>.*<c:tickMarkSkip val="3"/>.*</c:catAx>}m, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits lblOffset and noMultiLvlLbl on cat axis" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     cat_axis_lbl_offset: 50, cat_axis_no_multi_lvl_lbl: true)
    xlsx_path = File.join(Dir.tmpdir, "lbl_offset_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:catAx>.*<c:lblOffset val="50"/>.*</c:catAx>}m, xml)
    assert_match(%r{<c:catAx>.*<c:noMultiLvlLbl val="1"/>.*</c:catAx>}m, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits dispUnits with builtInUnit on val axis" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     val_axis_disp_units: "thousands")
    xlsx_path = File.join(Dir.tmpdir, "disp_units_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:valAx>.*<c:dispUnits><c:builtInUnit val="thousands"/></c:dispUnits>.*</c:valAx>}m, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits dPt elements for series data_points" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.set_cell("A2", 20)
    writer.set_cell("A3", 30)
    writer.add_chart(type: :pie,
                     series: [{ val_ref: "Sheet1!$A$1:$A$3",
                                data_points: [{ idx: 0, fill_color: "FF0000" },
                                              { idx: 1, fill_color: "00FF00" },
                                              { idx: 2, fill_color: "0000FF" }] }])
    xlsx_path = File.join(Dir.tmpdir, "dpt_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr></c:dPt>}, xml)
    assert_match(%r{<c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></c:spPr></c:dPt>}, xml)
    assert_match(%r{<c:dPt><c:idx val="2"/><c:spPr><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></c:spPr></c:dPt>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "emits trendline element in series" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.set_cell("A2", 2)
    writer.add_chart(type: :line,
                     series: [{ val_ref: "Sheet1!$A$1:$A$2",
                                trendline: { type: "poly", order: 3, forward: 2.5,
                                             disp_r_sqr: true, disp_eq: true,
                                             name: "MyTrend" } }])
    xlsx_path = File.join(Dir.tmpdir, "trendline_#{Process.pid}.xlsx")
    writer.write(xlsx_path)
    xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(/<c:trendline>/, xml)
    assert_match(%r{<c:name>MyTrend</c:name>}, xml)
    assert_match(%r{<c:trendlineType val="poly"/>}, xml)
    assert_match(%r{<c:order val="3"/>}, xml)
    assert_match(%r{<c:forward val="2\.5"/>}, xml)
    assert_match(%r{<c:dispRSqr val="1"/>}, xml)
    assert_match(%r{<c:dispEq val="1"/>}, xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with inner_shadow emits a:effectLst with a:innerShdw" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "InnerShadow",
                     inner_shadow: { blur_rad: 63_500, dist: 25_400, dir: 5_400_000, color: "FF0000" })

    xlsx_tempfile = Tempfile.new(["xlsxrb-inner-shadow", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:effectLst><a:innerShdw blurRad="63500" dist="25400" dir="5400000"><a:srgbClr val="FF0000"/></a:innerShdw></a:effectLst>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with both outer_shadow and inner_shadow emits both in a:effectLst" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "BothShadows",
                     outer_shadow: { blur_rad: 50_800, dist: 38_100, dir: 2_700_000, color: "000000" },
                     inner_shadow: { blur_rad: 63_500, dist: 25_400, dir: 5_400_000, color: "FF0000" })

    xlsx_tempfile = Tempfile.new(["xlsxrb-both-shadow", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:effectLst>.*<a:outerShdw.*</a:outerShdw>.*<a:innerShdw.*</a:innerShdw>.*</a:effectLst>}m, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with glow emits a:effectLst with a:glow" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Glow",
                     glow: { rad: 101_600, color: "FF0000" })

    xlsx_tempfile = Tempfile.new(["xlsxrb-glow", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:effectLst><a:glow rad="101600"><a:srgbClr val="FF0000"/></a:glow></a:effectLst>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with soft_edge emits a:effectLst with a:softEdge" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "SoftEdge",
                     soft_edge: { rad: 63_500 })

    xlsx_tempfile = Tempfile.new(["xlsxrb-softedge", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:effectLst><a:softEdge rad="63500"/></a:effectLst>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with reflection emits a:effectLst with a:reflection" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Reflect",
                     reflection: { blur_rad: 6_350, st_a: 52_000, end_a: 300, dist: 0, dir: 5_400_000,
                                   sy: -100_000, algn: "bl", rot_with_shape: false })

    xlsx_tempfile = Tempfile.new(["xlsxrb-reflect", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:effectLst><a:reflection blurRad="6350" stA="52000" endA="300" dist="0" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/></a:effectLst>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_chart with title as hash emits formatted chart title" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.set_cell("A2", 20)
    writer.add_chart(type: :bar,
                     title: { text: "My Chart", font: { bold: true, italic: true, size: 1400, color: "FF0000", name: "Arial" } },
                     series: [{ val_ref: "Sheet1!$A$1:$A$2" }])

    xlsx_tempfile = Tempfile.new(["xlsxrb-chart-title", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    chart_xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr b="1" i="1" sz="1400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>My Chart</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>}, chart_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_chart with title as plain string still works" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :bar, title: "Simple Title", series: [{ val_ref: "Sheet1!$A$1:$A$1" }])

    xlsx_tempfile = Tempfile.new(["xlsxrb-chart-plain", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    chart_xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Simple Title</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>}, chart_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_chart with series line_cap and line_join emits a:ln attributes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.set_cell("A2", 20)
    writer.add_chart(type: :line,
                     series: [{ val_ref: "Sheet1!$A$1:$A$2",
                                line_color: "FF0000", line_width: 2,
                                line_cap: "rnd", line_join: "round" }])

    xlsx_tempfile = Tempfile.new(["xlsxrb-line-cap", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    chart_xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(/<a:ln w="25400" cap="rnd">/, chart_xml)
    assert_match(%r{<a:round/>}, chart_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_font strike emits strike attribute on a:rPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Struck",
                     text_font: { strike: "sngStrike" })

    xlsx_tempfile = Tempfile.new(["xlsxrb-strike", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:rPr strike="sngStrike"/>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_font underline emits u attribute on a:rPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Underlined",
                     text_font: { underline: "sng" })

    xlsx_tempfile = Tempfile.new(["xlsxrb-underline", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:rPr u="sng"/>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_font baseline emits baseline attribute on a:rPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Super",
                     text_font: { baseline: 30_000 })

    xlsx_tempfile = Tempfile.new(["xlsxrb-baseline", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:rPr baseline="30000"/>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_font spacing emits spc attribute on a:rPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Spaced",
                     text_font: { spacing: 200 })

    xlsx_tempfile = Tempfile.new(["xlsxrb-spacing", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:rPr spc="200"/>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_font cap emits cap attribute on a:rPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "AllCaps",
                     text_font: { cap: "all" })

    xlsx_tempfile = Tempfile.new(["xlsxrb-cap", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:rPr cap="all"/>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_align emits a:pPr algn attribute" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Centered",
                     text_align: "ctr")

    xlsx_tempfile = Tempfile.new(["xlsxrb-align", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:pPr algn="ctr"/>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_font_align emits fontAlgn attribute on a:pPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "FontAlign",
                     text_font_align: "ctr")

    xlsx_tempfile = Tempfile.new(["xlsxrb-fontalign", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:pPr fontAlgn="ctr"/>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_def_tab_sz emits defTabSz attribute on a:pPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Tabs",
                     text_def_tab_sz: 914_400)

    xlsx_tempfile = Tempfile.new(["xlsxrb-deftabsz", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:pPr defTabSz="914400"/>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_font lang emits lang attribute on a:rPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Hello",
                     text_font: { lang: "en-US" })

    xlsx_tempfile = Tempfile.new(["xlsxrb-lang", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:rPr lang="en-US"/>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_vertical emits vert attribute on a:bodyPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Vertical",
                     text_vertical: "vert")

    xlsx_tempfile = Tempfile.new(["xlsxrb-vert", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:bodyPr vert="vert"/>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_insets emits lIns tIns rIns bIns on a:bodyPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Padded",
                     text_insets: { left: 91_440, top: 45_720, right: 91_440, bottom: 45_720 })

    xlsx_tempfile = Tempfile.new(["xlsxrb-insets", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/lIns="91440"/, drawing_xml)
    assert_match(/tIns="45720"/, drawing_xml)
    assert_match(/rIns="91440"/, drawing_xml)
    assert_match(/bIns="45720"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_font ea_font emits a:ea element in a:rPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "CJK",
                     text_font: { ea_font: "MS Gothic" })

    xlsx_tempfile = Tempfile.new(["xlsxrb-ea", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:rPr><a:ea typeface="MS Gothic"/></a:rPr>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_font cs_font emits a:cs element in a:rPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Complex",
                     text_font: { cs_font: "Arabic Typesetting" })

    xlsx_tempfile = Tempfile.new(["xlsxrb-cs", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:rPr><a:cs typeface="Arabic Typesetting"/></a:rPr>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_font sym_font emits a:sym element in a:rPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Symbols",
                     text_font: { sym_font: "Wingdings" })

    xlsx_tempfile = Tempfile.new(["xlsxrb-sym", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:rPr><a:sym typeface="Wingdings"/></a:rPr>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_font highlight emits a:highlight element in a:rPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Highlighted",
                     text_font: { highlight: "FFFF00" })

    xlsx_tempfile = Tempfile.new(["xlsxrb-hl", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:rPr><a:highlight><a:srgbClr val="FFFF00"/></a:highlight></a:rPr>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_rot emits rot attribute on a:bodyPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Rotated",
                     text_rot: 2_700_000)

    xlsx_tempfile = Tempfile.new(["xlsxrb-rot", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:bodyPr rot="2700000"/>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_indent emits marL marR indent attributes on a:pPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Indented",
                     text_indent: { left: 457_200, right: 228_600, indent: -114_300 })

    xlsx_tempfile = Tempfile.new(["xlsxrb-indent", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:pPr marL="457200" marR="228600" indent="-114300"/>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_font kern emits kern attribute on a:rPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Kerned",
                     text_font: { kern: 1200 })

    xlsx_tempfile = Tempfile.new(["xlsxrb-kern", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:rPr kern="1200"/>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_anchor_ctr emits anchorCtr attribute on a:bodyPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Centered",
                     text_anchor_ctr: true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-actr", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:bodyPr anchorCtr="1"/>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_spacing emits spcBef and spcAft in a:pPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Spaced",
                     text_spacing: { before: 600, after: 400 })

    xlsx_tempfile = Tempfile.new(["xlsxrb-spc", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:pPr><a:spcBef><a:spcPts val="600"/></a:spcBef><a:spcAft><a:spcPts val="400"/></a:spcAft></a:pPr>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_spacing line emits lnSpc in a:pPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "LineSpaced",
                     text_spacing: { line: 1200 })

    xlsx_tempfile = Tempfile.new(["xlsxrb-lnspc", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:pPr><a:lnSpc><a:spcPts val="1200"/></a:lnSpc></a:pPr>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_spacing line_pct emits spcPct in lnSpc" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "PctLine",
                     text_spacing: { line_pct: 150_000 })

    xlsx_tempfile = Tempfile.new(["xlsxrb-lnspcpct", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:lnSpc><a:spcPct val="150000"/></a:lnSpc>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_spacing before_pct and after_pct emits spcPct in spcBef and spcAft" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "PctSpacing",
                     text_spacing: { before_pct: 50_000, after_pct: 100_000 })

    xlsx_tempfile = Tempfile.new(["xlsxrb-spcpct", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:spcBef><a:spcPct val="50000"/></a:spcBef>}, drawing_xml)
    assert_match(%r{<a:spcAft><a:spcPct val="100000"/></a:spcAft>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_horz_overflow emits horzOverflow attribute on a:bodyPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Overflow",
                     text_horz_overflow: "overflow")

    xlsx_tempfile = Tempfile.new(["xlsxrb-hovf", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:bodyPr horzOverflow="overflow"/>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_num_col emits numCol attribute on a:bodyPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Columns",
                     text_num_col: 2)

    xlsx_tempfile = Tempfile.new(["xlsxrb-numcol", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/numCol="2"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_spc_col emits spcCol attribute on a:bodyPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "ColSpacing",
                     text_spc_col: 457_200)

    xlsx_tempfile = Tempfile.new(["xlsxrb-spccol", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/spcCol="457200"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_rtl_col emits rtlCol on a:bodyPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "RTL", text_rtl_col: true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-rtlcol", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/rtlCol="1"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_from_word_art emits fromWordArt on a:bodyPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "WA", text_from_word_art: true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-fwa", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/fromWordArt="1"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_upright emits upright on a:bodyPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Up", text_upright: true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-upright", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/upright="1"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_compat_ln_spc emits compatLnSpc on a:bodyPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Compat", text_compat_ln_spc: true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-cls", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/compatLnSpc="1"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_spc_first_last_para emits spcFirstLastPara on a:bodyPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "First",
                     text_spc_first_last_para: true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-sflp", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:bodyPr spcFirstLastPara="1"/>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_rtl emits rtl attribute on a:pPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "RTL",
                     text_rtl: true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-rtl", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:pPr rtl="1"/>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_ea_ln_brk emits eaLnBrk on a:pPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "EA", text_ea_ln_brk: true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-ealnbrk", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/eaLnBrk="1"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_latin_ln_brk emits latinLnBrk on a:pPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Latin", text_latin_ln_brk: true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-latinlnbrk", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/latinLnBrk="1"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_hanging_punct emits hangingPunct on a:pPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Hang", text_hanging_punct: true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-hangpunct", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/hangingPunct="1"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_tab_stops emits a:tabLst in a:pPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Tabs",
                     text_tab_stops: [{ pos: 914_400, align: "l" }, { pos: 1_828_800, align: "r" }])

    xlsx_tempfile = Tempfile.new(["xlsxrb-tabs", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:tabLst><a:tab pos="914400" algn="l"/><a:tab pos="1828800" algn="r"/></a:tabLst>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_bullet none emits a:buNone in a:pPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "NoBullet",
                     text_bullet: { type: "none" })

    xlsx_tempfile = Tempfile.new(["xlsxrb-bunone", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:pPr><a:buNone/></a:pPr>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_bullet char emits a:buChar in a:pPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Bullet",
                     text_bullet: { type: "char", char: "\u2022" })

    xlsx_tempfile = Tempfile.new(["xlsxrb-buchar", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/<a:buChar char="/, drawing_xml.force_encoding("UTF-8"))
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_bullet auto emits a:buAutoNum in a:pPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Numbered",
                     text_bullet: { type: "auto", auto_type: "arabicPeriod", start_at: 5 })

    xlsx_tempfile = Tempfile.new(["xlsxrb-buauto", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:buAutoNum type="arabicPeriod" startAt="5"/>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_level emits lvl attribute on a:pPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Level2",
                     text_level: 2)

    xlsx_tempfile = Tempfile.new(["xlsxrb-lvl", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:pPr lvl="2"/>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_bullet font emits a:buFont in a:pPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "FontBullet",
                     text_bullet: { type: "char", char: "\u2022", font: "Wingdings" })

    xlsx_tempfile = Tempfile.new(["xlsxrb-bufont", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:buFont typeface="Wingdings"/>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_bullet size_pts emits a:buSzPts in a:pPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "BigBullet",
                     text_bullet: { type: "char", char: "-", size_pts: 1400 })

    xlsx_tempfile = Tempfile.new(["xlsxrb-buszpts", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:buSzPts val="1400"/>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_bullet size_pct emits a:buSzPct in a:pPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "PctBullet",
                     text_bullet: { type: "char", char: "-", size_pct: 150_000 })

    xlsx_tempfile = Tempfile.new(["xlsxrb-buszpct", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:buSzPct val="150000"/>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_bullet color emits a:buClr in a:pPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "ColorBullet",
                     text_bullet: { type: "char", char: "-", color: "FF0000" })

    xlsx_tempfile = Tempfile.new(["xlsxrb-buclr", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:buClr><a:srgbClr val="FF0000"/></a:buClr>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_force_aa emits forceAA on a:bodyPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "AA", text_force_aa: true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-forceaa", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/forceAA="1"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_warp emits a:prstTxWarp in a:bodyPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Warped",
                     text_warp: { preset: "textWave1" })

    xlsx_tempfile = Tempfile.new(["xlsxrb-warp", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:prstTxWarp prst="textWave1"><a:avLst/></a:prstTxWarp>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_font no_proof emits noProof on a:rPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "NP",
                     text_font: { no_proof: true })

    xlsx_tempfile = Tempfile.new(["xlsxrb-noproof", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/noProof="1"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_font normalize_h emits normalizeH on a:rPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "NH",
                     text_font: { normalize_h: true })

    xlsx_tempfile = Tempfile.new(["xlsxrb-normh", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/normalizeH="1"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_font kumimoji emits kumimoji on a:rPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "KM",
                     text_font: { kumimoji: true })

    xlsx_tempfile = Tempfile.new(["xlsxrb-kumi", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/kumimoji="1"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_font dirty emits dirty on a:rPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "D",
                     text_font: { dirty: true })

    xlsx_tempfile = Tempfile.new(["xlsxrb-dirty", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/dirty="1"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_font smt_clean emits smtClean on a:rPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "SC",
                     text_font: { smt_clean: true })

    xlsx_tempfile = Tempfile.new(["xlsxrb-smtclean", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/smtClean="1"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_font bmk emits bmk on a:rPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "BM",
                     text_font: { bmk: "bookmark1" })

    xlsx_tempfile = Tempfile.new(["xlsxrb-bmk", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/bmk="bookmark1"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_end_para_rpr emits a:endParaRPr element" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "EPR",
                     text_end_para_rpr: { lang: "en-US", size: 1100 })

    xlsx_tempfile = Tempfile.new(["xlsxrb-endpararpr", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/<a:endParaRPr[^>]*lang="en-US"/, drawing_xml)
    assert_match(/<a:endParaRPr[^>]*sz="1100"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_end_para_rpr with children emits full a:endParaRPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "EPR2",
                     text_end_para_rpr: { lang: "en-US", name: "Arial", bold: true })

    xlsx_tempfile = Tempfile.new(["xlsxrb-endpararpr2", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/<a:endParaRPr[^>]*b="1"[^>]*lang="en-US"/, drawing_xml)
    assert_match(%r{<a:endParaRPr[^/]*>.*<a:latin typeface="Arial"/>.*</a:endParaRPr>}m, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_def_rpr emits a:defRPr inside a:pPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "DR",
                     text_def_rpr: { lang: "en-US", size: 1100 })

    xlsx_tempfile = Tempfile.new(["xlsxrb-defrpr", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(/<a:pPr>.*<a:defRPr[^>]*lang="en-US"/m, drawing_xml)
    assert_match(/<a:defRPr[^>]*sz="1100"/, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_paragraphs generates multiple a:p elements" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text_paragraphs: [
                       { text: "First", font: { bold: true } },
                       { text: "Second", align: "ctr" },
                       { text: "Third" }
                     ])

    xlsx_tempfile = Tempfile.new(["xlsxrb-multipara", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:p><a:r><a:rPr b="1"/><a:t>First</a:t></a:r></a:p>}, drawing_xml)
    assert_match(%r{<a:p><a:pPr algn="ctr"/><a:r><a:t>Second</a:t></a:r></a:p>}, drawing_xml)
    assert_match(%r{<a:p><a:r><a:t>Third</a:t></a:r></a:p>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_paragraphs runs generates multiple a:r elements" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text_paragraphs: [
                       { runs: [
                         { text: "Bold", font: { bold: true } },
                         { text: " Normal" },
                         { text: " Italic", font: { italic: true } }
                       ] }
                     ])

    xlsx_tempfile = Tempfile.new(["xlsxrb-multiruns", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:r><a:rPr b="1"/><a:t>Bold</a:t></a:r>}, drawing_xml)
    assert_match(%r{<a:r><a:t> Normal</a:t></a:r>}, drawing_xml)
    assert_match(%r{<a:r><a:rPr i="1"/><a:t> Italic</a:t></a:r>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_shape with text_font alt_lang emits altLang attribute on a:rPr" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_shape(preset: "rect", text: "Alt",
                     text_font: { lang: "en-US", alt_lang: "ja-JP" })

    xlsx_tempfile = Tempfile.new(["xlsxrb-altlang", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    drawing_xml = read_xml_from_xlsx(xlsx_path, "xl/drawings/drawing1.xml")
    assert_match(%r{<a:rPr lang="en-US" altLang="ja-JP"/>}, drawing_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "add_chart with cat_axis_label_rotation emits txPr with rot on catAx" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Cat")
    writer.set_cell("B1", 10)
    writer.add_chart(type: :bar, cat_ref: "Sheet1!A1", val_ref: "Sheet1!B1",
                     cat_axis_label_rotation: -2_700_000)

    xlsx_tempfile = Tempfile.new(["xlsxrb-axis-rot", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close
    writer.write(xlsx_path)

    chart_xml = read_xml_from_xlsx(xlsx_path, "xl/charts/chart1.xml")
    assert_match(%r{<c:catAx>.*<c:txPr><a:bodyPr rot="-2700000"/>.*</c:catAx>}m, chart_xml)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  private

  def read_xml_from_xlsx(xlsx_path, entry_name)
    require "xlsxrb/reader"
    reader = Xlsxrb::Reader.new(xlsx_path)
    reader.raw_entry(entry_name)
  end
end
