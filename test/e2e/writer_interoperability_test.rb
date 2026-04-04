# frozen_string_literal: true

require "test_helper"
require "open3"
require "tempfile"

class WriterInteroperabilityTest < Test::Unit::TestCase
  SCENARIO_DIR = File.expand_path("../fixtures/sdk_scenarios", __dir__)

  test "writer output passes Open XML SDK validation and value checks" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_string_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output preserves multiple inline strings in the same row" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("B1", "world")
    writer.set_cell("A1", "hello")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_same_row_multiple_strings_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores numeric cells correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 42)
    writer.set_cell("B1", 3.14)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_numeric_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores boolean cells correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", true)
    writer.set_cell("B1", false)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_boolean_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores formula cells correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.set_cell("A2", 20)
    writer.set_cell("A3", Xlsxrb::Formula.new(expression: "SUM(A1:A2)", cached_value: "30"))
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_formula_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores multiple sheets correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.add_sheet("Data")
    writer.set_cell("A1", "main", sheet: "Sheet1")
    writer.set_cell("A1", "data", sheet: "Data")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_multi_sheet_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores column widths correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_column_width("A", 20.0)
    writer.set_column_width("C", 15.5)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_column_width_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores row attributes correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_row_height(1, 25.0)
    writer.set_row_hidden(3)
    writer.set_row_style(5, 0)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_row_attributes_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores cellXf xfId linkage correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    style_xf_id = writer.add_named_cell_style(name: "Heading1", num_fmt_id: 0)
    cell_xf_id = writer.add_cell_style(xf_id: style_xf_id)
    writer.set_cell_style("A1", cell_xf_id)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_cell_xf_xfid_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores merge cells correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "merged")
    writer.merge_cells("A1:B2")
    writer.merge_cells("C3:D4")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_merge_cells_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores hyperlinks correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Example")
    writer.add_hyperlink("A1", "https://example.com")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_hyperlink_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores hyperlinks with display tooltip and location" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
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

    assert_openxml_sdk_scenario_passes("writer_hyperlink_deep_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores styles correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    fmt_id = writer.add_number_format("0.00")
    writer.set_cell("A1", 3.14)
    writer.set_cell_format("A1", fmt_id)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_styles_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores date cells correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", Date.new(2024, 1, 15))
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_date_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores auto filter correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Name")
    writer.set_cell("B1", "Age")
    writer.set_auto_filter("A1:B10")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_auto_filter_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores core properties correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_core_property(:title, "My Workbook")
    writer.set_core_property(:creator, "Test User")
    writer.set_core_property(:created, "2024-01-15T00:00:00Z")
    writer.set_core_property(:modified, "2024-01-16T12:00:00Z")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_core_properties_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores app properties correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_app_property(:application, "Xlsxrb")
    writer.set_app_property(:app_version, "1.0.0")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_app_properties_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores workbook properties correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.add_sheet("Data")
    writer.set_cell("A1", "main", sheet: "Sheet1")
    writer.set_cell("A1", "data", sheet: "Data")
    writer.set_workbook_property(:date1904, false)
    writer.set_workbook_property(:default_theme_version, 166_925)
    writer.set_workbook_view(:active_tab, 1)
    writer.set_calc_property(:calc_id, 191_029)
    writer.set_calc_property(:full_calc_on_load, true)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_workbook_properties_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores sheet states correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.add_sheet("Hidden")
    writer.add_sheet("VeryHidden")
    writer.set_cell("A1", "main", sheet: "Sheet1")
    writer.set_sheet_state("Hidden", :hidden)
    writer.set_sheet_state("VeryHidden", :very_hidden)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_sheet_state_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores defined names correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.add_sheet("Data")
    writer.set_cell("A1", "main", sheet: "Sheet1")
    writer.add_defined_name("MyRange", "Sheet1!$A$1:$B$10")
    writer.add_defined_name("LocalName", "Data!$C$1", sheet: "Data")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_defined_names_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores sheet properties correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_sheet_property(:tab_color, "FF0000FF")
    writer.set_sheet_property(:summary_below, false)
    writer.set_sheet_property(:summary_right, true)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_sheet_properties_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores dimension correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("B2", "hello")
    writer.set_cell("D5", "world")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_dimension_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores sheet format properties correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_sheet_format(:default_row_height, 18.0)
    writer.set_sheet_format(:default_col_width, 12.5)
    writer.set_sheet_format(:base_col_width, 10)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_sheet_format_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores row and column attributes correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_row_outline_level(2, 1)
    writer.set_row_collapsed(3)
    writer.set_column_attribute("B", :hidden, true)
    writer.set_column_attribute("C", :outline_level, 2)
    writer.set_column_attribute("C", :collapsed, true)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_row_col_attrs_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores sheet view, freeze pane, and selection correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_sheet_view(:show_grid_lines, false)
    writer.set_sheet_view(:zoom_scale, 150)
    writer.set_freeze_pane(row: 1, col: 1)
    writer.set_selection("C5")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_sheet_view_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores print options, page setup, and breaks correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_print_option(:grid_lines, true)
    writer.set_page_margins(left: 0.7, right: 0.7, top: 0.75, bottom: 0.75)
    writer.set_page_setup(:orientation, "landscape")
    writer.set_page_setup(:paper_size, 9)
    writer.set_header_footer(:odd_header, "&CPage &P")
    writer.add_row_break(10)
    writer.add_row_break(20)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_print_page_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores filter columns and sort state correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_auto_filter("A1:C10")
    writer.add_filter_column(0, { type: :filters, values: %w[A B] })
    writer.add_filter_column(1, { type: :custom, operator: "greaterThan", val: "100" })
    writer.set_sort_state("A1:B10", [{ ref: "A1:A10" }, { ref: "B1:B10", descending: true }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_filter_sort_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores data validations correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
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

    assert_openxml_sdk_scenario_passes("writer_data_validation_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores conditional formatting correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
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

    assert_openxml_sdk_scenario_passes("writer_conditional_format_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores expanded styles correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    fid = writer.add_font(bold: true, sz: 14, name: "Arial", color: "FFFF0000")
    fill_id = writer.add_fill(pattern: "solid", fg_color: "FF00FF00")
    brd_id = writer.add_border(left: { style: "thin" }, right: { style: "thin" },
                               top: { style: "thin" }, bottom: { style: "thin" })
    style_id = writer.add_cell_style(font_id: fid, fill_id: fill_id, border_id: brd_id)
    writer.add_dxf(font: { bold: true, color: "FFFF0000" })
    writer.set_cell("A1", "styled")
    writer.set_cell_style("A1", style_id)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_expanded_styles_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores tables and shared strings correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.use_shared_strings!
    writer.set_cell("A1", "Name")
    writer.set_cell("B1", "Age")
    writer.set_cell("A2", "Alice")
    writer.set_cell("B2", 30)
    writer.add_table("A1:B5", columns: %w[Name Age])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_table_sst_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  # --- Phase 2: Steps 141-145 ---

  test "writer output stores images correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    # Minimal 1x1 white pixel PNG.
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
    writer.set_cell("A1", "with image")
    writer.insert_image(png_bytes, ext: "png", from_col: 0, from_row: 0, to_col: 5, to_row: 10)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_image_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores image description correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
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
    writer.set_cell("A1", "with image")
    writer.insert_image(png_bytes, ext: "png", name: "Logo", description: "Company logo image")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_image_descr_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores charts correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Category")
    writer.set_cell("B1", "Value")
    writer.set_cell("A2", "X")
    writer.set_cell("B2", 10)
    writer.add_chart(type: :bar, title: "Test Chart", cat_ref: "Sheet1!$A$2:$A$2", val_ref: "Sheet1!$B$2:$B$2")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_chart_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores comments correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "data")
    writer.add_comment("A1", "Hello comment", author: "Tester")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_comment_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores rich text comments correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
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

    assert_openxml_sdk_scenario_passes("writer_comment_rich_text_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores pivot tables correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Category")
    writer.set_cell("B1", "Amount")
    writer.set_cell("A2", "A")
    writer.set_cell("B2", 100)
    writer.add_sheet("PivotSheet")
    writer.add_pivot_table("Sheet1!A1:B2",
                           row_fields: [0],
                           data_fields: [{ fld: 1, name: "Sum of Amount", subtotal: "sum" }],
                           dest_ref: "A1:B3",
                           sheet: "PivotSheet")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_pivot_table_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer VBA guard raises without preserve_macros" do
    writer = Xlsxrb::Writer.new
    assert_false(writer.preserve_macros?)
    writer.preserve_macros!
    assert_true(writer.preserve_macros?)
  end

  test "writer generates valid sheet protection" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Protected data")
    writer.set_sheet_protection(password: "CF1A")

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_sheet_protection_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid workbook protection" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_workbook_protection(lock_structure: true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_workbook_protection_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid rich text in shared strings" do
    writer = Xlsxrb::Writer.new
    writer.use_shared_strings!
    rt = Xlsxrb::RichText.new(runs: [
                                { text: "Bold", font: { bold: true } },
                                { text: " Normal" }
                              ])
    writer.set_cell("A1", rt)

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_rich_text_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid shared and array formulas" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.set_cell("A2", 20)
    writer.set_cell("B1", Xlsxrb::Formula.new(expression: "A1*2", type: :shared, ref: "B1:B2", shared_index: 0, cached_value: "20"))
    writer.set_cell("B2", Xlsxrb::Formula.new(expression: "", type: :shared, shared_index: 0, cached_value: "40"))
    writer.set_cell("C1", Xlsxrb::Formula.new(expression: "SUM(A1:A2)", type: :array, ref: "C1", cached_value: "30"))

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_formulas_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid CF dataBar and iconSet deep attributes" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 50)
    writer.add_conditional_format("A1:A10", type: :data_bar, priority: 1,
                                            data_bar: {
                                              cfvo: [{ type: "min" }, { type: "max" }],
                                              color: "FF638EC6",
                                              min_length: 5, max_length: 90, show_value: false
                                            })
    writer.add_conditional_format("B1:B10", type: :icon_set, priority: 2,
                                            icon_set: {
                                              icon_set: "3Arrows",
                                              cfvo: [{ type: "percent", val: "0" },
                                                     { type: "percent", val: "33" },
                                                     { type: "percent", val: "67" }],
                                              reverse: true, show_value: false
                                            })

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_cf_deep_attrs_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid DV with deep attributes" do
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

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_dv_deep_attrs_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid cellStyleXfs and cellStyles" do
    writer = Xlsxrb::Writer.new
    fid = writer.add_font(bold: true, sz: 14, name: "Arial")
    writer.add_named_cell_style(name: "Heading1", font_id: fid, builtin_id: 1)
    writer.set_cell("A1", "Hello")

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_styles_deep_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid chart with multiple series and axis titles" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Cat")
    writer.set_cell("B1", 10)
    writer.set_cell("C1", 20)
    writer.add_chart(type: :bar, title: "Multi",
                     series: [
                       { cat_ref: "Sheet1!$A$1:$A$2", val_ref: "Sheet1!$B$1:$B$2" },
                       { cat_ref: "Sheet1!$A$1:$A$2", val_ref: "Sheet1!$C$1:$C$2" }
                     ],
                     legend: { position: "b" },
                     data_labels: { show_val: true },
                     cat_axis_title: "Category",
                     val_axis_title: "Value")

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_chart_deep_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid chart with legend entries" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Cat")
    writer.set_cell("B1", 10)
    writer.set_cell("C1", 20)
    writer.add_chart(type: :bar, title: "LegendEntries",
                     series: [
                       { cat_ref: "Sheet1!$A$1:$A$2", val_ref: "Sheet1!$B$1:$B$2" },
                       { cat_ref: "Sheet1!$A$1:$A$2", val_ref: "Sheet1!$C$1:$C$2" }
                     ],
                     legend: { position: "b", entries: [{ idx: 1, delete: true }] })

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_legend_entries_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid chart with series line formatting" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.set_cell("A2", 20)
    writer.add_chart(type: :line, title: "SeriesLine",
                     series: [{ val_ref: "Sheet1!$A$1:$A$2", line_color: "0000FF", line_width: 2 }])

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_series_line_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid chart with series marker" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.set_cell("A2", 20)
    writer.add_chart(type: :line, title: "SeriesMarker",
                     series: [{ val_ref: "Sheet1!$A$1:$A$2", marker_symbol: "diamond", marker_size: 8 }])

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_series_marker_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid marker line dash" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :line,
                     series: [{ val_ref: "Sheet1!$A$1", marker_symbol: "circle",
                                marker_line_color: "000000", marker_line_dash: "dash" }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_marker_line_dash_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid marker noFill" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :line,
                     series: [{ val_ref: "Sheet1!$A$1", marker_symbol: "circle",
                                marker_no_fill: true }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_marker_no_fill_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid marker no line" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :line,
                     series: [{ val_ref: "Sheet1!$A$1", marker_symbol: "circle",
                                marker_fill: "FF0000", marker_no_line: true }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_marker_no_line_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid shapes with preset geometry" do
    writer = Xlsxrb::Writer.new
    writer.add_shape(preset: "ellipse", text: "Hello", name: "Oval 1",
                     from_col: 1, from_row: 2, to_col: 4, to_row: 6)
    writer.add_shape(preset: "roundRect", name: "RR 1",
                     from_col: 5, from_row: 0, to_col: 8, to_row: 3)

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_shape_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid shape with adjust values" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "roundRect",
                     adjust_values: [{ name: "adj", fmla: "val 16667" }])

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_adjust_values_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid shape with text font properties" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", text: "Bold",
                     text_font: { bold: true, italic: true, size: 1400, color: "FF0000", name: "Arial" })

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_text_font_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid pivot table with col_fields and items" do
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

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_pivot_deep_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid external link" do
    writer = Xlsxrb::Writer.new
    writer.add_external_link(target: "Book2.xlsx", sheet_names: %w[Data Summary])

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_external_link_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid table with totals row and enhanced columns" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Item")
    writer.set_cell("B1", "Price")
    writer.set_cell("C1", "Tax")
    writer.add_table("A1:C5", columns: [
                       "Item",
                       { name: "Price", totals_row_function: "sum" },
                       { name: "Tax", calculated_column_formula: "[Price]*0.1" }
                     ], totals_row_count: 1, style: { name: "TableStyleLight1" })

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_table_deep_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores cell alignment correctly" do
    writer = Xlsxrb::Writer.new
    style_id = writer.add_cell_style(
      alignment: { horizontal: "center", vertical: "top", wrap_text: true,
                   text_rotation: 45, indent: 2, shrink_to_fit: true }
    )
    writer.set_cell("A1", "aligned")
    writer.set_cell_style("A1", style_id)

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_alignment_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores extended font attributes correctly" do
    writer = Xlsxrb::Writer.new
    fid = writer.add_font(
      bold: true, italic: true, strike: true, sz: 12, name: "Calibri",
      color: "FF0000FF", underline: "double", vert_align: "superscript",
      scheme: "minor", family: 2
    )
    style_id = writer.add_cell_style(font_id: fid)
    writer.set_cell("A1", "extended")
    writer.set_cell_style("A1", style_id)

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_font_extended_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores gradient fill correctly" do
    writer = Xlsxrb::Writer.new
    fill_id = writer.add_fill(
      gradient: { type: "linear", degree: 90,
                  stops: [{ position: 0, color: "FFFF0000" }, { position: 1, color: "FF0000FF" }] }
    )
    style_id = writer.add_cell_style(fill_id: fill_id)
    writer.set_cell("A1", "gradient")
    writer.set_cell_style("A1", style_id)

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_gradient_fill_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores diagonal border correctly" do
    writer = Xlsxrb::Writer.new
    brd_id = writer.add_border(
      left: { style: "thin" }, right: { style: "thin" },
      top: { style: "thin" }, bottom: { style: "thin" },
      diagonal: { style: "thin", color: "FFFF0000" },
      diagonal_up: true, diagonal_down: true
    )
    style_id = writer.add_cell_style(border_id: brd_id)
    writer.set_cell("A1", "diag")
    writer.set_cell_style("A1", style_id)

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_diagonal_border_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores cell protection correctly" do
    writer = Xlsxrb::Writer.new
    style_id = writer.add_cell_style(protection: { locked: false, hidden: true })
    writer.set_cell("A1", "protected")
    writer.set_cell_style("A1", style_id)

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_cell_protection_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores font theme and indexed colors correctly" do
    writer = Xlsxrb::Writer.new
    fid1 = writer.add_font(sz: 11, name: "Calibri", theme: 1, tint: -0.25)
    fid2 = writer.add_font(sz: 11, name: "Calibri", indexed: 10)
    s1 = writer.add_cell_style(font_id: fid1)
    s2 = writer.add_cell_style(font_id: fid2)
    writer.set_cell("A1", "theme")
    writer.set_cell_style("A1", s1)
    writer.set_cell("A2", "indexed")
    writer.set_cell_style("A2", s2)

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_theme_color_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores fill theme colors correctly" do
    writer = Xlsxrb::Writer.new
    fill_id = writer.add_fill(pattern: "solid", fg_color_theme: 4, fg_color_tint: 0.6)
    style_id = writer.add_cell_style(fill_id: fill_id)
    writer.set_cell("A1", "theme fill")
    writer.set_cell_style("A1", style_id)

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_fill_theme_color_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid expanded CF rule types" do
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

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_cf_expanded_types_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid rich text with extended font attributes" do
    writer = Xlsxrb::Writer.new
    rt = Xlsxrb::RichText.new(runs: [
                                { text: "Strike", font: { strike: true, name: "Arial", sz: 11 } },
                                { text: "Double", font: { underline: "double", name: "Arial", sz: 11 } },
                                { text: "Super", font: { vert_align: "superscript", name: "Arial", sz: 11 } },
                                { text: "Theme", font: { theme: 1, tint: 0.5, name: "Calibri", sz: 11, family: 2, scheme: "minor" } }
                              ])
    writer.set_cell("A1", rt)

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_rich_text_extended_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid CF theme/indexed colors" do
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

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_cf_theme_color_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid gradient fill with theme/indexed stop colors" do
    writer = Xlsxrb::Writer.new
    fill_id = writer.add_fill(gradient: {
                                degree: 90,
                                stops: [{ position: 0, theme: 4, tint: -0.5 }, { position: 1, indexed: 12 }]
                              })
    style_id = writer.add_cell_style(fill_id: fill_id)
    writer.set_cell("A1", "themed gradient")
    writer.set_cell_style("A1", style_id)

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_gradient_stop_theme_color_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid complete CF rule types" do
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

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_cf_complete_types_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid DXF with alignment, protection, numFmt" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.add_dxf(
      font: { bold: true, color: "FFFF0000" },
      num_fmt: { num_fmt_id: 164, format_code: "#,##0.00" },
      alignment: { horizontal: "center", wrap_text: true },
      protection: { locked: false, hidden: true }
    )

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_dxf_deep_attrs_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores error cell values correctly" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", Xlsxrb::CellError.new(code: "#N/A"))
    writer.set_cell("B1", Xlsxrb::CellError.new(code: "#DIV/0!"))
    writer.set_cell("C1", Xlsxrb::CellError.new(code: "#VALUE!"))
    writer.set_cell("D1", Xlsxrb::CellError.new(code: "#REF!"))
    writer.set_cell("E1", Xlsxrb::CellError.new(code: "#NAME?"))
    writer.set_cell("F1", Xlsxrb::CellError.new(code: "#NUM!"))
    writer.set_cell("G1", Xlsxrb::CellError.new(code: "#NULL!"))

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_error_cells_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores extended core properties correctly" do
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

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_core_properties_extended_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores split pane correctly" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_split_pane(x_split: 2400, y_split: 1800, top_left_cell: "C4")

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_split_pane_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores colorFilter and iconFilter correctly" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Header1")
    writer.set_cell("B1", "Header2")
    writer.set_auto_filter("A1:B10")
    writer.add_dxf(fill: { pattern: "solid", fg_color: "FFFF0000" })
    writer.add_filter_column(0, { type: :color_filter, dxf_id: 0 })
    writer.add_filter_column(1, { type: :icon_filter, icon_set: "3Arrows", icon_id: 1 })

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_color_icon_filter_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores Time values as fractional serial with datetime format" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", Time.utc(2024, 3, 15, 14, 30, 0))

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_datetime_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores print area and print titles correctly" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_print_area("A1:D20")
    writer.set_print_titles(rows: "1:3", cols: "A:B")

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_print_area_titles_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores hashed password sheet protection correctly" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "protected")
    hp = Xlsxrb.hash_password("mypassword", spin_count: 1000)
    writer.set_sheet_protection(**hp)

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_hash_password_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores first page header/footer with differentFirst and differentOddEven" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_header_footer(:odd_header, "&LOdd Header")
    writer.set_header_footer(:even_header, "&LEven Header")
    writer.set_header_footer(:first_header, "&CFirst Page Header")
    writer.set_header_footer(:first_footer, "&CFirst Page Footer")
    writer.set_header_footer(:different_first, true)
    writer.set_header_footer(:different_odd_even, true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_header_footer_first_page_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores extended page setup attributes correctly" do
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

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_page_setup_extended_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores data validation showDropDown and imeMode correctly" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_data_validation("A1:A10", type: "list",
                                         formula1: '"Yes,No"',
                                         show_drop_down: true,
                                         ime_mode: "hiragana")

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_data_validation_extended_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores alignment readingOrder and justifyLastLine correctly" do
    writer = Xlsxrb::Writer.new
    sid = writer.add_cell_style(alignment: { horizontal: "distributed", reading_order: 2, justify_last_line: true })
    writer.set_cell("A1", "RTL text")
    writer.set_cell_style("A1", sid)

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_alignment_extended_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores font charset attribute correctly" do
    writer = Xlsxrb::Writer.new
    fid = writer.add_font(name: "MS Gothic", sz: 11, family: 3, charset: 128)
    sid = writer.add_cell_style(font_id: fid)
    writer.set_cell("A1", "テスト")
    writer.set_cell_style("A1", sid)

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_font_charset_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores extended sheetFormatPr attributes correctly" do
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_sheet_format(:default_row_height, 15)
    writer.set_sheet_format(:outline_level_row, 3)
    writer.set_sheet_format(:outline_level_col, 2)
    writer.set_sheet_format(:zero_height, true)
    writer.set_sheet_format(:custom_height, true)

    xlsx_tempfile = Tempfile.new(["xlsxrb-writer-e2e", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer.write(xlsx_path)
    assert_openxml_sdk_scenario_passes("writer_sheet_format_extended_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores showFormulas on sheet view correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_sheet_view(:show_formulas, true)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_show_formulas_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores codeName on sheet properties correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_sheet_property(:code_name, "MySheet")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_code_name_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores workbook view visibility correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_workbook_view(:visibility, "hidden")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_workbook_view_visibility_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores phonetic properties correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "hello")
    writer.set_phonetic_properties({ font_id: 1, type: "Hiragana", alignment: "center" })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_phonetic_pr_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores custom document properties correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.add_custom_property("Project", "Alpha", type: :lpwstr)
    writer.add_custom_property("Version", 42, type: :i4)
    writer.add_custom_property("Active", true, type: :bool)
    writer.set_cell("A1", "data")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_custom_properties_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores font shadow outline condense extend correctly" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    fid = writer.add_font(name: "Arial", sz: 12, shadow: true, outline: true, condense: true, extend: true)
    sid = writer.add_cell_style(font_id: fid)
    writer.set_cell("A1", "effects")
    writer.set_cell_style("A1", sid)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_font_effects_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer output stores sheetView showZeros view showOutlineSymbols showRuler" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_sheet_view(:show_zeros, false)
    writer.set_sheet_view(:view, "pageBreakPreview")
    writer.set_sheet_view(:show_outline_symbols, false)
    writer.set_sheet_view(:show_ruler, false)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_sheet_view_extended_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid fileVersion element" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_file_version(:app_name, "xl")
    writer.set_file_version(:last_edited, "7")
    writer.set_file_version(:lowest_edited, "7")
    writer.set_file_version(:rup_build, "27425")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_file_version_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid fileSharing element" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_file_sharing(:read_only_recommended, true)
    writer.set_file_sharing(:user_name, "TestUser")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_file_sharing_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid protectedRanges element" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_protected_range(name: "EditArea", sqref: "A1:B10")
    writer.add_protected_range(name: "SecureRange", sqref: "C1:D5",
                               algorithm_name: "SHA-512",
                               hash_value: "abc123", salt_value: "salt456",
                               spin_count: 100_000)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_protected_ranges_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid indexed colors in stylesheet" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_indexed_colors(%w[FF000000 FFFFFFFF FFFF0000])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_indexed_colors_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid tableStyles element" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_table_styles_option(:default_table_style, "TableStyleMedium2")
    writer.set_table_styles_option(:default_pivot_style, "PivotStyleLight16")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_table_styles_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid cellWatches element" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 100)
    writer.set_cell("B2", 200)
    writer.add_cell_watch("A1")
    writer.add_cell_watch("B2")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_cell_watches_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid dataConsolidate element" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "data")
    writer.set_data_consolidate(
      function: "average", start_labels: true, link: true,
      data_refs: [{ ref: "A1:B10", sheet: "Sheet1" }, { ref: "C1:D10", name: "Range2" }]
    )
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_data_consolidate_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid scenarios element" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
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

    assert_openxml_sdk_scenario_passes("writer_scenarios_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid ignoredErrors element" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "123")
    writer.add_ignored_error(sqref: "A1:B2", number_stored_as_text: true, eval_error: true)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_ignored_errors_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid autoFilter extended attributes" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Name")
    writer.set_cell("B1", "Date")
    writer.set_auto_filter("A1:B10")
    writer.add_filter_column(0, { type: :filters, values: %w[Alice],
                                  calendar_type: "gregorian",
                                  date_group_items: [{ date_time_grouping: "year", year: 2024 }],
                                  hidden_button: true, show_button: false })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_auto_filter_extended_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid conformance attribute" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_workbook_property(:conformance, "transitional")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_conformance_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid pivotTableStyleInfo element" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    %w[cat1 cat2 cat3].each_with_index { |v, i| writer.set_cell("A#{i + 1}", v) }
    %w[x y z].each_with_index { |v, i| writer.set_cell("B#{i + 1}", v) }
    [10, 20, 30].each_with_index { |v, i| writer.set_cell("C#{i + 1}", v) }
    writer.add_pivot_table("Sheet1!A1:C4",
                           row_fields: [0],
                           data_fields: [{ fld: 2, name: "Sum", subtotal: "sum" }],
                           pivot_table_style: {
                             name: "PivotStyleLight16",
                             show_row_headers: true,
                             show_col_headers: true,
                             show_row_stripes: false,
                             show_col_stripes: false,
                             show_last_column: true
                           })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_pivot_table_style_info_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid expanded chart types (area, scatter, doughnut, radar)" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "a")
    %i[area scatter doughnut radar].each do |t|
      writer.add_chart(type: t, cat_ref: "Sheet1!$A$1:$A$1", val_ref: "Sheet1!$A$1:$A$1")
    end
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_chart_types_expanded_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid totalsRowFormula in table column" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Item")
    writer.set_cell("B1", "Price")
    writer.add_table("A1:B3", columns: [
                       "Item",
                       { name: "Price", totals_row_function: "custom",
                         totals_row_formula: "SUBTOTAL(109,[Price])" }
                     ], totals_row_count: 1)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_totals_row_formula_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid view3D element in chart" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :bar3d,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     view_3d: { rot_x: 30, rot_y: 20 })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_view3d_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid numFmt elements on chart axes" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     cat_axis_num_fmt: { format_code: "General", source_linked: true },
                     val_axis_num_fmt: { format_code: "0.00", source_linked: false })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_axis_numfmt_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid majorTickMark and minorTickMark on chart axes" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     cat_axis_major_tick_mark: "cross",
                     cat_axis_minor_tick_mark: "in",
                     val_axis_major_tick_mark: "out",
                     val_axis_minor_tick_mark: "none")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_axis_tickmark_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid crosses elements on chart axes" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     cat_axis_crosses: "autoZero",
                     val_axis_crosses: "max")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_axis_crosses_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid crossBetween, majorUnit, minorUnit on value axis" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     val_axis_cross_between: "between",
                     val_axis_major_unit: 10,
                     val_axis_minor_unit: 2)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_val_axis_units_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid scaling max and min on chart value axis" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     val_axis_scaling_max: 100,
                     val_axis_scaling_min: 0)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_axis_scaling_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid logBase element in chart axis scaling" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     val_axis_log_base: 10)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_axis_logbase_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid firstSliceAng and holeSize for doughnut chart" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :doughnut,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     first_slice_ang: 45,
                     hole_size: 50)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_pie_angles_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid smooth and marker elements for line chart" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :line,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     smooth: true,
                     marker: true)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_line_chart_smooth_marker_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid dropLines and hiLowLines for line chart" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :line,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     drop_lines: true, hi_low_lines: true)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_drop_lines_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid upDownBars for line chart" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :line,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     up_down_bars: { gap_width: 150 })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_up_down_bars_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid scatterStyle element for scatter chart" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "a")
    writer.add_chart(type: :scatter,
                     series: [{ cat_ref: "Sheet1!$A$1:$A$1", val_ref: "Sheet1!$A$1:$A$1" }],
                     scatter_style: "lineMarker")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_scatter_style_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid gapDepth and shape for 3D bar chart" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :bar3d,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     gap_depth: 150, bar_shape: "cylinder")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_bar3d_gap_shape_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid bubbleScale, showNegBubbles, sizeRepresents for bubble chart" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :bubble,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     bubble_3d: true, bubble_scale: 80, show_neg_bubbles: false,
                     size_represents: "area")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_bubble_chart_props_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid crossesAt element on value axis" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     val_axis_crosses_at: 5.0)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_axis_crosses_at_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid wireframe element for surface chart" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :surface,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     wireframe: true)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_surface_wireframe_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid tickLblSkip and tickMarkSkip on category axis" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     cat_axis_tick_lbl_skip: 2, cat_axis_tick_mark_skip: 3)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_axis_tick_skip_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid lblOffset and noMultiLvlLbl on category axis" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     cat_axis_lbl_offset: 50, cat_axis_no_multi_lvl_lbl: true)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_axis_lbl_offset_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid dispUnits and builtInUnit on value axis" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     val_axis_disp_units: "thousands")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_axis_disp_units_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid fileRecoveryPr in workbook" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.set_file_recovery_property(:auto_recover, false)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_file_recovery_pr_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid solidFill color on shape" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", fill_color: "FF0000")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_shape_fill_color_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid line color and width on shape" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", line_color: "0000FF", line_width: 12_700)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_shape_line_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid line color and width on image" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
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
    writer.set_cell("A1", "with image")
    writer.insert_image(png_bytes, ext: "png", name: "Pic1",
                                   line_color: "FF0000", line_width: 25_400)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_image_line_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid image with srcRect cropping" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
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
    writer.set_cell("A1", "with image")
    writer.insert_image(png_bytes, ext: "png", name: "Pic1",
                                   src_rect: { top: 10_000, bottom: 20_000, left: 5000, right: 15_000 })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_src_rect_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid image with alphaModFix transparency" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
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
    writer.set_cell("A1", "with image")
    writer.insert_image(png_bytes, ext: "png", name: "Pic1", alpha_mod_fix: 50_000)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_alpha_mod_fix_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid shape with autofit modes" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", text: "No autofit", autofit: "none")
    writer.add_shape(preset: "rect", text: "Shape autofit", autofit: "shape")
    writer.add_shape(preset: "rect", text: "Normal autofit", autofit: "normal")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_autofit_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid shape with outer shadow" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", text: "Shadow",
                     outer_shadow: { blur_rad: 50_800, dist: 38_100, dir: 2_700_000, color: "000000" })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_outer_shadow_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid shape with gradient fill" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", text: "Grad",
                     gradient_fill: { stops: [{ pos: 0, color: "FF0000" }, { pos: 100_000, color: "0000FF" }], angle: 5_400_000 })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_shape_gradient_fill_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid shape with line dash style" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", text: "Dashed", line_color: "000000", line_dash: "dash")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_line_dash_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid shape with line end arrows" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", text: "Arrow", line_color: "000000",
                     head_end: { type: "triangle", w: "med", len: "med" },
                     tail_end: { type: "stealth", w: "lg", len: "lg" })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_line_end_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid noFill and noLine on shape" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", no_fill: true, no_line: true)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_shape_no_fill_no_line_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid chart data table (dTable)" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     data_table: { show_horz_border: true, show_vert_border: false,
                                   show_outline: true, show_keys: true })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_chart_data_table_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid solidFill on chart series" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1", fill_color: "00FF00" }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_chart_series_fill_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid plot area background fill color" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     plot_area_fill: "CCCCCC")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_chart_plot_area_fill_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid plot area line dash" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     plot_area_line_color: "0000FF",
                     plot_area_line_dash: "dash")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_plot_area_line_dash_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid plot area noFill" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     plot_area_no_fill: true)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_plot_area_no_fill_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid chart space fill and border" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     chart_fill: "EEEEEE",
                     chart_line_color: "333333",
                     chart_line_width: 1.0)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_chart_space_fill_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid chart space line dash" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     chart_line_color: "333333",
                     chart_line_dash: "lgDash")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_chart_space_line_dash_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid chart space noFill" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     chart_no_fill: true)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_chart_space_no_fill_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid series data points on pie chart" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.set_cell("A2", 20)
    writer.set_cell("A3", 30)
    writer.add_chart(type: :pie,
                     series: [{ val_ref: "Sheet1!$A$1:$A$3",
                                data_points: [{ idx: 0, fill_color: "FF0000" },
                                              { idx: 1, fill_color: "00FF00" },
                                              { idx: 2, fill_color: "0000FF" }] }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_data_points_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid data point no_line" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.set_cell("A2", 20)
    writer.add_chart(type: :pie,
                     series: [{ val_ref: "Sheet1!$A$1:$A$2",
                                data_points: [{ idx: 0, fill_color: "FF0000", no_line: true }] }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_dpt_no_line_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid data point marker" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.set_cell("A2", 20)
    writer.add_chart(type: :line,
                     series: [{ val_ref: "Sheet1!$A$1:$A$2",
                                data_points: [{ idx: 0, marker_symbol: "diamond", marker_size: 8 }] }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_dpt_marker_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid data point marker spPr" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.set_cell("A2", 20)
    writer.add_chart(type: :line,
                     series: [{ val_ref: "Sheet1!$A$1:$A$2",
                                data_points: [{ idx: 0, marker_symbol: "square", marker_size: 6,
                                                marker_fill: "00FF00", marker_line_color: "0000FF",
                                                marker_line_width: 1.5 }] }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_dpt_marker_sp_pr_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid data point marker noFill" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.set_cell("A2", 20)
    writer.add_chart(type: :line,
                     series: [{ val_ref: "Sheet1!$A$1:$A$2",
                                data_points: [{ idx: 0, marker_symbol: "circle",
                                                marker_no_fill: true, marker_no_line: true }] }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_dpt_marker_no_fill_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid trendline in line chart series" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.set_cell("A2", 2)
    writer.add_chart(type: :line,
                     series: [{ val_ref: "Sheet1!$A$1:$A$2",
                                trendline: { type: "poly", order: 3,
                                             forward: 2.5, disp_r_sqr: true,
                                             disp_eq: true, name: "MyTrend" } }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_trendline_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid shape with inner shadow" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", text: "InnerShadow",
                     inner_shadow: { blur_rad: 63_500, dist: 25_400, dir: 5_400_000, color: "FF0000" })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_inner_shadow_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid shape with glow effect" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", text: "Glow",
                     glow: { rad: 101_600, color: "FF0000" })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_glow_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid shape with soft edge effect" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", text: "SoftEdge",
                     soft_edge: { rad: 63_500 })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_soft_edge_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid shape with reflection effect" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", text: "Reflect",
                     reflection: { blur_rad: 6_350, st_a: 52_000, end_a: 300, dist: 0, dir: 5_400_000,
                                   sy: -100_000, algn: "bl", rot_with_shape: false })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_reflection_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid chart with formatted title" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.set_cell("A2", 20)
    writer.add_chart(type: :bar,
                     title: { text: "My Chart", font: { bold: true, italic: true, size: 1400, color: "FF0000", name: "Arial" } },
                     series: [{ val_ref: "Sheet1!$A$1:$A$2" }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_chart_title_font_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid chart series with line cap and join" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.set_cell("A2", 20)
    writer.add_chart(type: :line,
                     series: [{ val_ref: "Sheet1!$A$1:$A$2",
                                line_color: "FF0000", line_width: 2,
                                line_cap: "rnd", line_join: "round" }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_series_line_cap_join_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid shape with strikethrough text" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", text: "Struck",
                     text_font: { strike: "sngStrike" })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_text_strike_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid shape with underlined text" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", text: "Underlined",
                     text_font: { underline: "sng" })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_text_underline_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid shape with baseline text" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", text: "Super",
                     text_font: { baseline: 30_000 })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_text_baseline_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid shape with text spacing" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", text: "Spaced",
                     text_font: { spacing: 200 })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_text_spacing_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid shape with cap text" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", text: "AllCaps",
                     text_font: { cap: "all" })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_text_cap_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid shape with text paragraph alignment" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", text: "Centered",
                     text_align: "ctr")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_text_align_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid shape with text run language" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", text: "Hello",
                     text_font: { lang: "en-US" })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_text_lang_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid shape with text vertical direction" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", text: "Vertical",
                     text_vertical: "vert")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_text_vertical_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid shape with text inset margins" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", text: "Padded",
                     text_insets: { left: 91_440, top: 45_720, right: 91_440, bottom: 45_720 })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_text_insets_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid chart with category axis label rotation" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "Cat")
    writer.set_cell("B1", 10)
    writer.add_chart(type: :bar, cat_ref: "Sheet1!A1", val_ref: "Sheet1!B1",
                     cat_axis_label_rotation: -2_700_000)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_axis_label_rotation_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid shape with East Asian font" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", text: "CJK",
                     text_font: { ea_font: "MS Gothic" })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_text_ea_font_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid shape with Complex Script font" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", text: "Complex",
                     text_font: { cs_font: "Arabic Typesetting" })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_text_cs_font_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid shape with text rotation" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", text: "Rotated",
                     text_rot: 2_700_000)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_text_rot_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid shape with paragraph indent" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", text: "Indented",
                     text_indent: { left: 457_200, right: 228_600, indent: -114_300 })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_text_indent_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid shape with text kerning" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", text: "Kerned",
                     text_font: { kern: 1200 })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_text_kern_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid shape with anchor center" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", text: "Centered",
                     text_anchor_ctr: true)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_text_anchor_ctr_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid shape with paragraph spacing" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", "test")
    writer.add_shape(preset: "rect", text: "Spaced",
                     text_spacing: { before: 600, after: 400 })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_text_spacing_para_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid axis title spPr with fill and line" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     cat_axis_title: { text: "Category", fill_color: "FFEECC", line_color: "CC6600", line_width: 0.5 },
                     val_axis_title: { text: "Value", fill_color: "EEFFEE", line_color: "006600", line_width: 1.0 },
                     series: [{ val_ref: "Sheet1!$A$1" }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_axis_title_sp_pr_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid title line dash" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     title: { text: "My Chart", line_color: "000000", line_dash: "dot" },
                     cat_axis_title: { text: "Category", line_color: "FF0000", line_dash: "dash" },
                     val_axis_title: { text: "Value", line_color: "00FF00", line_dash: "dashDot" },
                     series: [{ val_ref: "Sheet1!$A$1" }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_title_line_dash_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid axis title noFill" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     cat_axis_title: { text: "Category", no_fill: true },
                     val_axis_title: { text: "Value", no_fill: true },
                     series: [{ val_ref: "Sheet1!$A$1" }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_axis_title_no_fill_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid leader lines with spPr in pie chart dLbls" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :pie,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     data_labels: { show_val: true, show_leader_lines: true,
                                    leader_lines: { line_color: "FF0000", line_width: 0.5 } })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_leader_lines_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid legend manual layout" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     legend: { position: "b", layout: { x: 0.1, y: 0.8, w: 0.8, h: 0.15 } },
                     series: [{ val_ref: "Sheet1!$A$1" }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_legend_layout_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid axis fill" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     cat_axis_fill: "F0F0F0",
                     val_axis_fill: "E0E0E0")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_axis_fill_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid axis noFill" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     cat_axis_no_fill: true,
                     val_axis_no_fill: true)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_axis_no_fill_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid axis line dash" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     cat_axis_line_color: "FF0000", cat_axis_line_dash: "dot",
                     val_axis_line_color: "00FF00", val_axis_line_dash: "dashDot")
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_axis_line_dash_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid plot area manual layout" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     plot_area_layout: { target: "inner", x: 0.1, y: 0.2, w: 0.7, h: 0.6 })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_plot_area_layout_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid legend entry font" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     legend: { position: "r",
                               entries: [{ idx: 0, font: { size: 14, bold: true } }] })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_legend_entry_font_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid legend line_dash" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     legend: { position: "r", line_color: "000000", line_dash: "dash" })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_legend_line_dash_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid data labels line_dash" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     data_labels: { show_val: true, line_color: "333333", line_dash: "lgDash" })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_line_dash_extended_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid dLbl text" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.set_cell("A2", 20)
    writer.add_chart(type: :pie,
                     data_labels: { show_val: true,
                                    labels: [{ idx: 0, text: "Custom Label" }] },
                     series: [{ val_ref: "Sheet1!$A$1:$A$2" }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_dlbl_text_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid dLbl delete" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.set_cell("A2", 20)
    writer.add_chart(type: :pie,
                     data_labels: { show_val: true,
                                    labels: [{ idx: 1, delete: true }] },
                     series: [{ val_ref: "Sheet1!$A$1:$A$2" }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_dlbl_delete_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid dLbl numFmt" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.set_cell("A2", 20)
    writer.add_chart(type: :pie,
                     data_labels: { show_val: true,
                                    labels: [{ idx: 0,
                                               num_fmt: { format_code: "0.00%" } }] },
                     series: [{ val_ref: "Sheet1!$A$1:$A$2" }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_dlbl_numfmt_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid dLbl separator" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.set_cell("A2", 20)
    writer.add_chart(type: :pie,
                     data_labels: { show_val: true,
                                    labels: [{ idx: 0, separator: " | " }] },
                     series: [{ val_ref: "Sheet1!$A$1:$A$2" }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_dlbl_separator_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid dLbl spPr" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.set_cell("A2", 20)
    writer.add_chart(type: :pie,
                     data_labels: { show_val: true,
                                    labels: [{ idx: 0, fill_color: "AABB00",
                                               line_color: "112233", line_width: 1.5 }] },
                     series: [{ val_ref: "Sheet1!$A$1:$A$2" }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_dlbl_sppr_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid dLbl font" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.set_cell("A2", 20)
    writer.add_chart(type: :pie,
                     data_labels: { show_val: true,
                                    labels: [{ idx: 0,
                                               font: { size: 14, bold: true, color: "FF0000" } }] },
                     series: [{ val_ref: "Sheet1!$A$1:$A$2" }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_dlbl_font_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid serLines" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.set_cell("A2", 20)
    writer.add_chart(type: :bar,
                     grouping: :stacked,
                     ser_lines: true,
                     series: [{ val_ref: "Sheet1!$A$1:$A$2" }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_ser_lines_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid bandFmts" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :surface,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     band_fmts: [{ idx: 0, fill_color: "FF0000" },
                                 { idx: 1, fill_color: "00FF00" }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_band_fmts_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid trendlineLbl" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :line,
                     series: [{ val_ref: "Sheet1!$A$1",
                                trendline: { type: "linear", disp_eq: true,
                                             label: { num_fmt: "0.00%" } } }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_trendline_lbl_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid dispUnitsLbl" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     val_axis_disp_units: { built_in_unit: "thousands",
                                            label: { num_fmt: "#,##0" } })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_disp_units_lbl_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid ofPieChart" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.set_cell("B1", 20)
    writer.add_chart(type: :of_pie,
                     of_pie_type: "bar",
                     split_type: "pos",
                     split_pos: 2,
                     second_pie_size: 75,
                     series: [{ val_ref: "Sheet1!$A$1:$B$1" }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_of_pie_chart_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid custSplit on ofPieChart" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.set_cell("B1", 20)
    writer.set_cell("C1", 30)
    writer.set_cell("D1", 40)
    writer.add_chart(type: :of_pie,
                     of_pie_type: "pie",
                     split_type: "cust",
                     cust_split: [1, 3],
                     series: [{ val_ref: "Sheet1!$A$1:$D$1" }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_cust_split_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid custom error bars with plus and minus" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.set_cell("B1", 1.5)
    writer.set_cell("C1", 0.8)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1",
                                error_bars: { direction: "y", bar_type: "both",
                                              val_type: "cust",
                                              plus: "Sheet1!$B$1",
                                              minus: "Sheet1!$C$1" } }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_err_bars_cust_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid multiple trendlines" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.set_cell("A2", 4)
    writer.set_cell("A3", 9)
    writer.add_chart(type: :line,
                     series: [{ val_ref: "Sheet1!$A$1:$A$3",
                                trendlines: [
                                  { type: "linear", name: "Linear" },
                                  { type: "poly", order: 2, name: "Quadratic" }
                                ] }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_multi_trendlines_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid per-series shape on bar3D" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.set_cell("A2", 20)
    writer.add_chart(type: :bar3d,
                     series: [{ val_ref: "Sheet1!$A$1", shape: "cone" },
                              { val_ref: "Sheet1!$A$2", shape: "pyramid" }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_ser_shape_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid multiple errBars on scatter chart" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.set_cell("B1", 2)
    writer.add_chart(type: :scatter,
                     series: [{ cat_ref: "Sheet1!$A$1", val_ref: "Sheet1!$B$1",
                                error_bars_list: [
                                  { direction: "x", bar_type: "both", val_type: "fixedVal", val: 0.5 },
                                  { direction: "y", bar_type: "both", val_type: "fixedVal", val: 1.0 }
                                ] }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_multi_err_bars_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid chart title styling from flat params" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :bar,
                     title: "Styled",
                     title_font: { bold: true, size: 1400, color: "FF0000", name: "Arial" },
                     title_fill_color: "FFFF00",
                     title_line_color: "000000",
                     title_line_width: 1.0,
                     title_line_dash: "dash",
                     series: [{ val_ref: "Sheet1!$A$1" }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_title_styling_flat_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid axis title styling from flat params" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :bar,
                     cat_axis_title: "Category",
                     cat_axis_title_font: { bold: true, size: 1200, color: "0000FF" },
                     cat_axis_title_fill: "EEEEFF",
                     cat_axis_title_line_color: "0000CC",
                     cat_axis_title_line_width: 0.5,
                     cat_axis_title_line_dash: "dot",
                     val_axis_title: "Value",
                     val_axis_title_font: { italic: true, size: 1000 },
                     val_axis_title_no_fill: true,
                     series: [{ val_ref: "Sheet1!$A$1" }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_axis_title_styling_flat_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "writer generates valid surface3DChart with serAx" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-writer", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :surface3d,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     wireframe: true)
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_surface3d_chart_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "SDK validates writer trendlineLbl styling" do
    xlsx_path = File.join(Dir.tmpdir, "writer_trendline_lbl_styling_#{Process.pid}.xlsx")
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :line,
                     series: [{ val_ref: "Sheet1!$A$1",
                                trendline: { type: "linear", disp_eq: true,
                                             label: { num_fmt: "0.00%",
                                                      fill_color: "FFFF00",
                                                      line_color: "0000FF",
                                                      line_width: 1.5,
                                                      line_dash: "dash",
                                                      font: { size: 10, bold: true, color: "FF0000" } } } }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_trendline_lbl_styling_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "SDK validates writer dispUnitsLbl styling" do
    xlsx_path = File.join(Dir.tmpdir, "writer_disp_units_lbl_styling_#{Process.pid}.xlsx")
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 10)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     val_axis_disp_units: { built_in_unit: "thousands",
                                            label: { fill_color: "FFFF00",
                                                     line_color: "0000FF",
                                                     line_width: 1.5,
                                                     line_dash: "dash",
                                                     font: { size: 10, bold: true, color: "FF0000" } } })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_disp_units_lbl_styling_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "SDK validates writer chart protection" do
    xlsx_path = File.join(Dir.tmpdir, "writer_chart_protection_#{Process.pid}.xlsx")
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     protection: { chart_object: true, data: true, formatting: false,
                                   selection: true, user_interface: true })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_chart_protection_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "SDK validates writer chart printSettings" do
    xlsx_path = File.join(Dir.tmpdir, "writer_chart_print_settings_#{Process.pid}.xlsx")
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     print_settings: {
                       header_footer: { odd_header: "&CHeader", odd_footer: "&CPage &P" },
                       page_margins: { b: 0.75, l: 0.7, r: 0.7, t: 0.75, header: 0.3, footer: 0.3 },
                       page_setup: { orientation: "landscape", paper_size: 1 }
                     })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_chart_print_settings_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "SDK validates writer dLbl layout" do
    xlsx_path = File.join(Dir.tmpdir, "writer_dlbl_layout_#{Process.pid}.xlsx")
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1",
                                data_labels: { show_val: true,
                                               labels: [{ idx: 0, layout: { x: 0.05, y: -0.03 }, show_val: true }] } }])
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_dlbl_layout_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "SDK validates writer chart-level font (txPr)" do
    xlsx_path = File.join(Dir.tmpdir, "writer_chart_font_#{Process.pid}.xlsx")
    writer = Xlsxrb::Writer.new
    writer.set_cell("A1", 1)
    writer.add_chart(type: :bar,
                     series: [{ val_ref: "Sheet1!$A$1" }],
                     chart_font: { size: 12, bold: true, name: "Arial", color: "333333" })
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_chart_font_test", xlsx_path)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  private

  def assert_openxml_sdk_scenario_passes(scenario_name, xlsx_path)
    scenario_path = File.join(SCENARIO_DIR, "#{scenario_name}.cs")
    assert(File.exist?(scenario_path), "Scenario file not found: #{scenario_path}")

    result = OpenXmlSdkScenarioRunner.run_single_scenario(scenario_path, xlsx_path)
    failure_reason = extract_failure_reason(result[:stderr])

    assert(
      result[:success],
      "Open XML SDK scenario failed: #{failure_reason}\n" \
      "Scenario: #{scenario_name}\n" \
      "XLSX: #{xlsx_path}\n" \
      "STDERR:\n#{result[:stderr]}"
    )
  end

  def extract_failure_reason(stderr)
    return "unknown reason" if stderr.nil? || stderr.strip.empty?

    lines = stderr.lines.map(&:strip).reject(&:empty?)
    exception_line = lines.find { |line| line.include?("Exception:") }
    return exception_line if exception_line

    scenario_line = lines.find { |line| line.start_with?("SCENARIO_") }
    return scenario_line if scenario_line

    lines.first
  end
end
