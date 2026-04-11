# frozen_string_literal: true

require "test_helper"
require "tempfile"

# E2E tests that validate XLSX files generated via Facade APIs
# (Streaming and In-Memory) using the OpenXML SDK.
#
# These tests are slower than CONTRACT tests (~1s each due to SDK invocation)
# but verify that Excel/other tools can actually open the generated files.
class FacadeInteroperabilityTest < Test::Unit::TestCase
  SCENARIO_DIR = File.expand_path("../fixtures/sdk_scenarios", __dir__)

  # ---- Helpers ----

  def generate_streaming_xlsx(&block)
    tmp = Tempfile.new(["facade_e2e_stream", ".xlsx"])
    Xlsxrb.generate(tmp.path, &block)
    tmp
  end

  def generate_in_memory_xlsx(&block)
    tmp = Tempfile.new(["facade_e2e_mem", ".xlsx"])
    wb = Xlsxrb.build(&block)
    Xlsxrb.write(tmp.path, wb)
    tmp
  end

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

  # =====================================================
  # Streaming API - Chart E2E
  # =====================================================

  test "streaming: bar chart passes SDK validation" do
    tmp = generate_streaming_xlsx do |w|
      w.add_sheet("Sales") do |s|
        s.add_row(["Month", "Revenue"])
        s.add_row(["Jan", 100])
        s.add_row(["Feb", 200])
        s.add_row(["Mar", 300])
        s.add_chart(type: :bar, title: "Quarterly Revenue",
                    series: [{ cat_ref: "Sales!$A$2:$A$4", val_ref: "Sales!$B$2:$B$4" }])
      end
    end

    assert_openxml_sdk_scenario_passes("facade_chart_bar_test", tmp.path)
  ensure
    tmp&.close!
  end

  test "streaming: pie chart passes SDK validation" do
    tmp = generate_streaming_xlsx do |w|
      w.add_sheet("Data") do |s|
        s.add_row(["Category", "Value"])
        s.add_row(["A", 40])
        s.add_row(["B", 60])
        s.add_chart(type: :pie, title: "Distribution",
                    series: [{ cat_ref: "Data!$A$2:$A$3", val_ref: "Data!$B$2:$B$3" }])
      end
    end

    assert_openxml_sdk_scenario_passes("facade_chart_pie_test", tmp.path)
  ensure
    tmp&.close!
  end

  test "streaming: line chart with multiple series passes SDK validation" do
    tmp = generate_streaming_xlsx do |w|
      w.add_sheet("Trends") do |s|
        s.add_row(["Month", "Series1", "Series2"])
        s.add_row(["Jan", 10, 20])
        s.add_row(["Feb", 15, 25])
        s.add_chart(type: :line, title: "Trends",
                    series: [
                      { cat_ref: "Trends!$A$2:$A$3", val_ref: "Trends!$B$2:$B$3" },
                      { cat_ref: "Trends!$A$2:$A$3", val_ref: "Trends!$C$2:$C$3" }
                    ])
      end
    end

    assert_openxml_sdk_scenario_passes("facade_chart_line_multi_series_test", tmp.path)
  ensure
    tmp&.close!
  end

  test "streaming: multiple charts on same sheet passes SDK validation" do
    tmp = generate_streaming_xlsx do |w|
      w.add_sheet("Multi") do |s|
        s.add_row(["X", "Y", "Z"])
        s.add_row([1, 10, 20])
        s.add_chart(type: :bar, title: "Chart1",
                    series: [{ cat_ref: "Multi!$A$2:$A$2", val_ref: "Multi!$B$2:$B$2" }])
        s.add_chart(type: :pie, title: "Chart2",
                    series: [{ cat_ref: "Multi!$A$2:$A$2", val_ref: "Multi!$C$2:$C$2" }])
      end
    end

    assert_openxml_sdk_scenario_passes("facade_chart_multi_test", tmp.path)
  ensure
    tmp&.close!
  end

  test "streaming: chart with legend and axis titles passes SDK validation" do
    tmp = generate_streaming_xlsx do |w|
      w.add_sheet("S1") do |s|
        s.add_row(["X", "Y"])
        s.add_row([1, 10])
        s.add_row([2, 20])
        s.add_chart(type: :bar, title: "Axes",
                    series: [{ cat_ref: "S1!$A$2:$A$3", val_ref: "S1!$B$2:$B$3" }],
                    legend: { position: "b" },
                    cat_axis_title: "Categories",
                    val_axis_title: "Values")
      end
    end

    assert_openxml_sdk_scenario_passes("facade_chart_legend_axes_test", tmp.path)
  ensure
    tmp&.close!
  end

  test "streaming: basic data passes SDK validation" do
    tmp = generate_streaming_xlsx do |w|
      w.add_sheet("Data") do |s|
        s.add_row(["Name", "Score", "Active"])
        s.add_row(["Alice", 95, true])
        s.add_row(["Bob", 87, false])
      end
    end

    assert_openxml_sdk_scenario_passes("facade_basic_data_test", tmp.path)
  ensure
    tmp&.close!
  end

  # =====================================================
  # In-Memory API - Chart E2E
  # =====================================================

  test "in_memory: bar chart passes SDK validation" do
    tmp = generate_in_memory_xlsx do |w|
      w.add_sheet("Sales") do |s|
        s.add_row(["Month", "Revenue"])
        s.add_row(["Jan", 100])
        s.add_row(["Feb", 200])
        s.add_row(["Mar", 300])
        s.add_chart(type: :bar, title: "Quarterly Revenue",
                    series: [{ cat_ref: "Sales!$A$2:$A$4", val_ref: "Sales!$B$2:$B$4" }])
      end
    end

    assert_openxml_sdk_scenario_passes("facade_chart_bar_test", tmp.path)
  ensure
    tmp&.close!
  end

  test "in_memory: pie chart passes SDK validation" do
    tmp = generate_in_memory_xlsx do |w|
      w.add_sheet("Data") do |s|
        s.add_row(["Category", "Value"])
        s.add_row(["A", 40])
        s.add_row(["B", 60])
        s.add_chart(type: :pie, title: "Distribution",
                    series: [{ cat_ref: "Data!$A$2:$A$3", val_ref: "Data!$B$2:$B$3" }])
      end
    end

    assert_openxml_sdk_scenario_passes("facade_chart_pie_test", tmp.path)
  ensure
    tmp&.close!
  end

  test "in_memory: line chart with multiple series passes SDK validation" do
    tmp = generate_in_memory_xlsx do |w|
      w.add_sheet("Trends") do |s|
        s.add_row(["Month", "Series1", "Series2"])
        s.add_row(["Jan", 10, 20])
        s.add_row(["Feb", 15, 25])
        s.add_chart(type: :line, title: "Trends",
                    series: [
                      { cat_ref: "Trends!$A$2:$A$3", val_ref: "Trends!$B$2:$B$3" },
                      { cat_ref: "Trends!$A$2:$A$3", val_ref: "Trends!$C$2:$C$3" }
                    ])
      end
    end

    assert_openxml_sdk_scenario_passes("facade_chart_line_multi_series_test", tmp.path)
  ensure
    tmp&.close!
  end

  test "in_memory: multiple charts on same sheet passes SDK validation" do
    tmp = generate_in_memory_xlsx do |w|
      w.add_sheet("Multi") do |s|
        s.add_row(["X", "Y", "Z"])
        s.add_row([1, 10, 20])
        s.add_chart(type: :bar, title: "Chart1",
                    series: [{ cat_ref: "Multi!$A$2:$A$2", val_ref: "Multi!$B$2:$B$2" }])
        s.add_chart(type: :pie, title: "Chart2",
                    series: [{ cat_ref: "Multi!$A$2:$A$2", val_ref: "Multi!$C$2:$C$2" }])
      end
    end

    assert_openxml_sdk_scenario_passes("facade_chart_multi_test", tmp.path)
  ensure
    tmp&.close!
  end

  test "in_memory: chart with legend and axis titles passes SDK validation" do
    tmp = generate_in_memory_xlsx do |w|
      w.add_sheet("S1") do |s|
        s.add_row(["X", "Y"])
        s.add_row([1, 10])
        s.add_row([2, 20])
        s.add_chart(type: :bar, title: "Axes",
                    series: [{ cat_ref: "S1!$A$2:$A$3", val_ref: "S1!$B$2:$B$3" }],
                    legend: { position: "b" },
                    cat_axis_title: "Categories",
                    val_axis_title: "Values")
      end
    end

    assert_openxml_sdk_scenario_passes("facade_chart_legend_axes_test", tmp.path)
  ensure
    tmp&.close!
  end

  test "in_memory: basic data passes SDK validation" do
    tmp = generate_in_memory_xlsx do |w|
      w.add_sheet("Data") do |s|
        s.add_row(["Name", "Score", "Active"])
        s.add_row(["Alice", 95, true])
        s.add_row(["Bob", 87, false])
      end
    end

    assert_openxml_sdk_scenario_passes("facade_basic_data_test", tmp.path)
  ensure
    tmp&.close!
  end
end
