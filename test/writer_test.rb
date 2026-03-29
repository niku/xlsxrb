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

  test "add_chart stores chart definition" do
    writer = Xlsxrb::Writer.new
    writer.add_chart(type: :bar, title: "Sales", cat_ref: "Sheet1!$A$1:$A$3", val_ref: "Sheet1!$B$1:$B$3")
    charts = writer.charts
    assert_equal(1, charts.size)
    assert_equal(:bar, charts[0][:type])
    assert_equal("Sales", charts[0][:title])
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

  test "set_workbook_protection stores protection settings" do
    writer = Xlsxrb::Writer.new
    writer.set_workbook_protection(lock_structure: true, lock_windows: false)
    prot = writer.workbook_protection
    assert_equal(true, prot[:lock_structure])
    assert_equal(false, prot[:lock_windows])
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
end
