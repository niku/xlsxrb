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
