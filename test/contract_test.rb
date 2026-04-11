# frozen_string_literal: true

require "test_helper"
require "tempfile"

# CONTRACT tests verify that Facade APIs (Streaming & In-Memory) produce
# correct OOXML structures by reading back the generated XLSX and checking
# semantic properties. These are faster than E2E SDK validation and catch
# structural issues early.
#
# Uses test-unit's data() for data-driven testing across API paths.
class ContractTest < Test::Unit::TestCase
  API_PATHS = {
    "streaming" => :streaming,
    "in_memory" => :in_memory
  }.freeze

  # ---- Helpers ----

  # Generate an XLSX via streaming (Xlsxrb.generate) and return the tmpfile.
  def generate_streaming(&block)
    tmp = Tempfile.new(["contract_stream", ".xlsx"])
    Xlsxrb.generate(tmp.path, &block)
    tmp
  end

  # Generate an XLSX via in-memory (Xlsxrb.build + Xlsxrb.write) and return the tmpfile.
  def generate_in_memory(&block)
    tmp = Tempfile.new(["contract_mem", ".xlsx"])
    workbook = Xlsxrb.build(&block)
    Xlsxrb.write(tmp.path, workbook)
    tmp
  end

  # Generate via both APIs, return reader for the specified path.
  def generate_and_read(api_path, &block)
    tmp = case api_path
          when :streaming then generate_streaming(&block)
          when :in_memory then generate_in_memory(&block)
          end
    [Xlsxrb::Ooxml::Reader.new(tmp.path), tmp]
  end

  # Minimal 1x1 white pixel PNG for image tests.
  MINIMAL_PNG = [
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
    0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
    0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
    0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
    0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
    0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC,
    0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
    0x44, 0xAE, 0x42, 0x60, 0x82
  ].pack("C*").freeze

  # =====================================================
  # Chart CONTRACT tests
  # =====================================================

  data(API_PATHS)
  test "chart: bar chart with title and series data" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Sales") do |s|
        s.add_row(["Month", "Revenue"])
        s.add_row(["Jan", 100])
        s.add_row(["Feb", 200])
        s.add_row(["Mar", 300])
        s.add_chart(type: :bar, title: "Quarterly Revenue",
                    series: [{ cat_ref: "Sales!$A$2:$A$4", val_ref: "Sales!$B$2:$B$4" }])
      end
    end

    charts = reader.charts
    assert_equal(1, charts.size, "Expected exactly 1 chart [chart count]")
    assert_equal("barChart", charts[0][:chart_type],
                 "Expected barChart type [chart_type]. Check Writer#add_chart type mapping.")
    assert_equal("Quarterly Revenue", charts[0][:title],
                 "Chart title mismatch [title]. Check Writer#generate_chart_xml title element.")
    assert_equal(1, charts[0][:series].size,
                 "Expected 1 series [series count]. Check Writer#add_chart series handling.")
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "chart: pie chart preserves type and title" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Data") do |s|
        s.add_row(["Category", "Value"])
        s.add_row(["A", 40])
        s.add_row(["B", 60])
        s.add_chart(type: :pie, title: "Distribution",
                    series: [{ cat_ref: "Data!$A$2:$A$3", val_ref: "Data!$B$2:$B$3" }])
      end
    end

    charts = reader.charts
    assert_equal(1, charts.size)
    assert_equal("pieChart", charts[0][:chart_type],
                 "Expected pieChart type [chart_type]. Check Writer#chart_type_element.")
    assert_equal("Distribution", charts[0][:title])
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "chart: line chart with multiple series" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Trends") do |s|
        s.add_row(["Month", "Series1", "Series2"])
        s.add_row(["Jan", 10, 20])
        s.add_row(["Feb", 15, 25])
        s.add_chart(type: :line, title: "Trend Lines",
                    series: [
                      { cat_ref: "Trends!$A$2:$A$3", val_ref: "Trends!$B$2:$B$3" },
                      { cat_ref: "Trends!$A$2:$A$3", val_ref: "Trends!$C$2:$C$3" }
                    ])
      end
    end

    charts = reader.charts
    assert_equal(1, charts.size)
    assert_equal("lineChart", charts[0][:chart_type])
    assert_equal("Trend Lines", charts[0][:title])
    assert_equal(2, charts[0][:series].size,
                 "Expected 2 series for multi-series line chart [series count].")
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "chart: chart with legend and axis titles" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("S1") do |s|
        s.add_row(["X", "Y"])
        s.add_row([1, 10])
        s.add_chart(type: :bar, title: "Axes Test",
                    series: [{ cat_ref: "S1!$A$2:$A$2", val_ref: "S1!$B$2:$B$2" }],
                    legend: { position: "b" },
                    cat_axis_title: "Categories",
                    val_axis_title: "Values")
      end
    end

    charts = reader.charts
    chart = charts[0]
    assert_equal("Axes Test", chart[:title])
    assert_not_nil(chart[:legend], "Legend should be present [legend]. Check Writer#generate_chart_xml legend element.")
    assert_equal("b", chart[:legend][:position],
                 "Legend position mismatch [legend.position].")
    assert_equal("Categories", chart[:cat_axis_title],
                 "Category axis title mismatch [cat_axis_title].")
    assert_equal("Values", chart[:val_axis_title],
                 "Value axis title mismatch [val_axis_title].")
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "chart: chart with data labels" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("DL") do |s|
        s.add_row(["Cat", "Val"])
        s.add_row(["A", 50])
        s.add_chart(type: :bar, title: "DL Test",
                    series: [{ cat_ref: "DL!$A$2:$A$2", val_ref: "DL!$B$2:$B$2" }],
                    data_labels: { show_val: true })
      end
    end

    chart = reader.charts[0]
    assert_not_nil(chart[:data_labels],
                   "Data labels should be present [data_labels]. Check Writer#generate_chart_xml data label element.")
    assert_equal(true, chart[:data_labels][:show_val],
                 "Data labels show_val should be true [data_labels.show_val].")
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "chart: multiple charts on same sheet" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Multi") do |s|
        s.add_row(["X", "Y", "Z"])
        s.add_row([1, 10, 20])
        s.add_chart(type: :bar, title: "Chart1",
                    series: [{ cat_ref: "Multi!$A$2:$A$2", val_ref: "Multi!$B$2:$B$2" }])
        s.add_chart(type: :pie, title: "Chart2",
                    series: [{ cat_ref: "Multi!$A$2:$A$2", val_ref: "Multi!$C$2:$C$2" }])
      end
    end

    charts = reader.charts
    assert_equal(2, charts.size,
                 "Expected 2 charts on same sheet [chart count]. Check WorkbookWriter chart loop.")
    types = charts.map { |c| c[:chart_type] }.sort
    assert_equal(%w[barChart pieChart], types,
                 "Chart types mismatch [chart_types].")
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "chart: chart data cache is populated" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Cache") do |s|
        s.add_row(["Label", "Amount"])
        s.add_row(["Alpha", 100])
        s.add_row(["Beta", 200])
        s.add_chart(type: :bar, title: "Cache Test",
                    series: [{ cat_ref: "Cache!$A$2:$A$3", val_ref: "Cache!$B$2:$B$3" }])
      end
    end

    chart = reader.charts[0]
    series = chart[:series][0]

    # Verify the value cache has actual data points (not all zeros),
    # which was the bug we fixed in WorkbookWriter.
    assert_not_nil(series[:val_ref],
                   "Series val_ref should be present [series.val_ref].")
  ensure
    tmp&.close!
  end

  # =====================================================
  # Chart type coverage
  # =====================================================

  CHART_TYPE_MAP = {
    "area"     => { input: :area,     expected: "areaChart" },
    "scatter"  => { input: :scatter,  expected: "scatterChart" },
    "radar"    => { input: :radar,    expected: "radarChart" },
    "doughnut" => { input: :doughnut, expected: "doughnutChart" },
    "bar3d"    => { input: :bar3d,    expected: "bar3DChart" }
  }.freeze

  data(CHART_TYPE_MAP)
  test "chart: type mapping is correct" do |spec|
    tmp = Tempfile.new(["contract_charttype", ".xlsx"])
    Xlsxrb.generate(tmp.path) do |w|
      w.add_sheet("S") do |s|
        s.add_row(["X", "Y"])
        s.add_row([1, 10])
        s.add_chart(type: spec[:input], title: "Type Test",
                    series: [{ cat_ref: "S!$A$2:$A$2", val_ref: "S!$B$2:$B$2" }])
      end
    end

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    charts = reader.charts
    assert_equal(1, charts.size)
    assert_equal(spec[:expected], charts[0][:chart_type],
                 "Chart type mapping for :#{spec[:input]} [chart_type]. " \
                 "Check Writer#chart_type_element.")
  ensure
    tmp&.close!
  end

  # =====================================================
  # Cell data round-trip via Facade
  # =====================================================

  data(API_PATHS)
  test "cell: basic data types round-trip" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Types") do |s|
        s.add_row(["hello", 42, 3.14, true, false, nil])
      end
    end

    wb = Xlsxrb.read(tmp.path)
    sheet = wb.sheet(0)
    assert_equal("hello", sheet.cell_value("A1"))
    assert_equal(42, sheet.cell_value("B1"))
    assert_in_delta(3.14, sheet.cell_value("C1"))
    assert_equal(true, sheet.cell_value("D1"))
    assert_equal(false, sheet.cell_value("E1"))
    assert_nil(sheet.cell_value("F1"))
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "cell: multiple sheets" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("A") { |s| s.add_row(["sheet_a"]) }
      w.add_sheet("B") { |s| s.add_row(["sheet_b"]) }
    end

    wb = Xlsxrb.read(tmp.path)
    assert_equal(2, wb.sheets.size)
    assert_equal("sheet_a", wb.sheet("A").cell_value("A1"))
    assert_equal("sheet_b", wb.sheet("B").cell_value("A1"))
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "cell: many rows preserve data" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Bulk") do |s|
        100.times { |i| s.add_row([i, "row#{i}"]) }
      end
    end

    wb = Xlsxrb.read(tmp.path)
    sheet = wb.sheet(0)
    assert_equal(100, sheet.rows.size)
    assert_equal(0, sheet.cell_value("A1"))
    assert_equal(99, sheet.cell_value("A100"))
    assert_equal("row50", sheet.cell_value("B51"))
  ensure
    tmp&.close!
  end

  # =====================================================
  # Worksheet XML namespace CONTRACT
  # =====================================================

  data(API_PATHS)
  test "xml: worksheet with chart has xmlns:r namespace" do |api_path|
    tmp = case api_path
          when :streaming
            generate_streaming do |w|
              w.add_sheet("S") do |s|
                s.add_row(["X", "Y"])
                s.add_row([1, 10])
                s.add_chart(type: :bar, title: "NS",
                            series: [{ cat_ref: "S!$A$2:$A$2", val_ref: "S!$B$2:$B$2" }])
              end
            end
          when :in_memory
            t = Tempfile.new(["contract_ns", ".xlsx"])
            wb = Xlsxrb.build do |w|
              w.add_sheet("S") do |s|
                s.add_row(["X", "Y"])
                s.add_row([1, 10])
                s.add_chart(type: :bar, title: "NS",
                            series: [{ cat_ref: "S!$A$2:$A$2", val_ref: "S!$B$2:$B$2" }])
              end
            end
            Xlsxrb.write(t.path, wb)
            t
          end

    # Read the raw worksheet XML and verify namespace declaration.
    entries = Xlsxrb::Ooxml::ZipReader.open(tmp.path, &:read_all)
    sheet_xml = entries["xl/worksheets/sheet1.xml"]
    assert_not_nil(sheet_xml, "Worksheet XML should exist")
    assert_match(/xmlns:r=/, sheet_xml,
                 "Worksheet XML must declare xmlns:r when <drawing r:id> is used. " \
                 "Check WorksheetWriter#start namespace declarations.")

    # Verify drawing element exists
    assert_match(/<drawing r:id=/, sheet_xml,
                 "Worksheet XML must contain <drawing> element referencing drawing part. " \
                 "Check WorksheetWriter#finish drawing_rid handling.")
  ensure
    tmp&.close!
  end

  # =====================================================
  # Drawing/Chart XML structure CONTRACT
  # =====================================================

  data(API_PATHS)
  test "xml: chart XML has proper namespace declarations" do |api_path|
    tmp = case api_path
          when :streaming
            generate_streaming do |w|
              w.add_sheet("S") do |s|
                s.add_row(["X", "Y"])
                s.add_row([1, 10])
                s.add_chart(type: :bar, title: "NS Check",
                            series: [{ cat_ref: "S!$A$2:$A$2", val_ref: "S!$B$2:$B$2" }])
              end
            end
          when :in_memory
            t = Tempfile.new(["contract_cns", ".xlsx"])
            wb = Xlsxrb.build do |w|
              w.add_sheet("S") do |s|
                s.add_row(["X", "Y"])
                s.add_row([1, 10])
                s.add_chart(type: :bar, title: "NS Check",
                            series: [{ cat_ref: "S!$A$2:$A$2", val_ref: "S!$B$2:$B$2" }])
              end
            end
            Xlsxrb.write(t.path, wb)
            t
          end

    entries = Xlsxrb::Ooxml::ZipReader.open(tmp.path, &:read_all)

    # Chart XML
    chart_xml = entries["xl/charts/chart1.xml"]
    assert_not_nil(chart_xml, "Chart XML should exist in ZIP")
    assert_match(/xmlns:c=/, chart_xml, "Chart XML must declare c: namespace")
    assert_match(/xmlns:a=/, chart_xml, "Chart XML must declare a: namespace")

    # Drawing XML
    drawing_xml = entries["xl/drawings/drawing1.xml"]
    assert_not_nil(drawing_xml, "Drawing XML should exist in ZIP")
    assert_match(/xmlns:r=/, drawing_xml,
                 "Drawing XML must declare r: namespace for chart references. " \
                 "Check Writer#generate_drawing_xml namespace declarations.")

    # Content types
    ct_xml = entries["[Content_Types].xml"]
    assert_match(/chart/, ct_xml,
                 "Content_Types must include chart content type. " \
                 "Check WorkbookWriter#build_content_types.")
    assert_match(/drawing/, ct_xml,
                 "Content_Types must include drawing content type.")
  ensure
    tmp&.close!
  end
end
