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
    writer.write(xlsx_path)

    assert_openxml_sdk_scenario_passes("writer_row_attributes_test", xlsx_path)
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

  private

  def assert_openxml_sdk_scenario_passes(scenario_name, xlsx_path)
    scenario_path = File.join(SCENARIO_DIR, "#{scenario_name}.cs")
    assert(File.exist?(scenario_path), "Scenario file not found: #{scenario_path}")

    command = sdk_runner_command(scenario_path, xlsx_path)
    stdout, stderr, status = Open3.capture3(*command)

    failure_reason = extract_failure_reason(stderr)

    assert(
      status.success?,
      "Open XML SDK scenario failed: #{failure_reason}\n" \
      "Scenario: #{scenario_name}\n" \
      "Command: #{command.join(" ")}\n" \
      "XLSX: #{xlsx_path}\n" \
      "STDOUT:\n#{stdout}\n" \
      "STDERR:\n#{stderr}"
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

  def sdk_runner_command(scenario_path, xlsx_path)
    [
      "dotnet", File.expand_path("../../vendor/sdk_runner/bin/Release/net8.0/sdk_runner.dll", __dir__),
      scenario_path, xlsx_path
    ]
  end
end
