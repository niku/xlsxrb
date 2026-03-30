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

  private

  # ensure zlib loaded
  def read_xml_from_xlsx(xlsx_path, entry_name)
    require "xlsxrb/reader"
    reader = Xlsxrb::Reader.new(xlsx_path)
    reader.raw_entry(entry_name)
  end
end
