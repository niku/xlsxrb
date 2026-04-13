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
  def generate_streaming(&)
    tmp = Tempfile.new(["contract_stream", ".xlsx"])
    Xlsxrb.generate(tmp.path, &)
    tmp
  end

  # Generate an XLSX via in-memory (Xlsxrb.build + Xlsxrb.write) and return the tmpfile.
  def generate_in_memory(&)
    tmp = Tempfile.new(["contract_mem", ".xlsx"])
    workbook = Xlsxrb.build(&)
    Xlsxrb.write(tmp.path, workbook)
    tmp
  end

  # Generate via both APIs, return reader for the specified path.
  def generate_and_read(api_path, &)
    tmp = case api_path
          when :streaming then generate_streaming(&)
          when :in_memory then generate_in_memory(&)
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
        s.add_row(%w[Month Revenue])
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
        s.add_row(%w[Category Value])
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
        s.add_row(%w[Month Series1 Series2])
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
        s.add_row(%w[X Y])
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
        s.add_row(%w[Cat Val])
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
        s.add_row(%w[X Y Z])
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
        s.add_row(%w[Label Amount])
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
    "area" => { input: :area, expected: "areaChart" },
    "scatter" => { input: :scatter, expected: "scatterChart" },
    "radar" => { input: :radar, expected: "radarChart" },
    "doughnut" => { input: :doughnut, expected: "doughnutChart" },
    "bar3d" => { input: :bar3d, expected: "bar3DChart" }
  }.freeze

  data(CHART_TYPE_MAP)
  test "chart: type mapping is correct" do |spec|
    tmp = Tempfile.new(["contract_charttype", ".xlsx"])
    Xlsxrb.generate(tmp.path) do |w|
      w.add_sheet("S") do |s|
        s.add_row(%w[X Y])
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
    _, tmp = generate_and_read(api_path) do |w|
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
    _, tmp = generate_and_read(api_path) do |w|
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
    _, tmp = generate_and_read(api_path) do |w|
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
                s.add_row(%w[X Y])
                s.add_row([1, 10])
                s.add_chart(type: :bar, title: "NS",
                            series: [{ cat_ref: "S!$A$2:$A$2", val_ref: "S!$B$2:$B$2" }])
              end
            end
          when :in_memory
            t = Tempfile.new(["contract_ns", ".xlsx"])
            wb = Xlsxrb.build do |w|
              w.add_sheet("S") do |s|
                s.add_row(%w[X Y])
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
                s.add_row(%w[X Y])
                s.add_row([1, 10])
                s.add_chart(type: :bar, title: "NS Check",
                            series: [{ cat_ref: "S!$A$2:$A$2", val_ref: "S!$B$2:$B$2" }])
              end
            end
          when :in_memory
            t = Tempfile.new(["contract_cns", ".xlsx"])
            wb = Xlsxrb.build do |w|
              w.add_sheet("S") do |s|
                s.add_row(%w[X Y])
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

  data(API_PATHS)
  test "xml: shape colors are normalized for DrawingML" do |api_path|
    tmp = case api_path
          when :streaming
            generate_streaming do |w|
              w.add_sheet("S") do |s|
                s.add_row(["See the shape below"])
                s.add_shape(preset: "rect", text: "Important!",
                            from_col: 0, from_row: 2, to_col: 3, to_row: 6,
                            fill_color: "#FFFFC0", line_color: "#FF0000")
              end
            end
          when :in_memory
            t = Tempfile.new(["contract_shape_color", ".xlsx"])
            wb = Xlsxrb.build do |w|
              w.add_sheet("S") do |s|
                s.add_row(["See the shape below"])
                s.add_shape(preset: "rect", text: "Important!",
                            from_col: 0, from_row: 2, to_col: 3, to_row: 6,
                            fill_color: "#FFFFC0", line_color: "#FF0000")
              end
            end
            Xlsxrb.write(t.path, wb)
            t
          end

    entries = Xlsxrb::Ooxml::ZipReader.open(tmp.path, &:read_all)
    drawing_xml = entries["xl/drawings/drawing1.xml"]
    assert_not_nil(drawing_xml, "Drawing XML should exist in ZIP")
    assert_match(%r{<a:srgbClr val="FFFFC0"/>}, drawing_xml,
                 "Shape fill color should be emitted as hex without '#'")
    assert_match(%r{<a:srgbClr val="FF0000"/>}, drawing_xml,
                 "Shape line color should be emitted as hex without '#'")
    assert_not_match(/<a:srgbClr val="#/, drawing_xml,
                     "DrawingML srgbClr must not include leading '#'")
  ensure
    tmp&.close!
  end

  # =====================================================
  # Hyperlink CONTRACT tests
  # =====================================================

  data(API_PATHS)
  test "hyperlink: external URL is preserved" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Links") do |s|
        s.add_row(["Click me"])
        s.add_hyperlink("A1", "https://example.com", display: "Example")
      end
    end

    links = reader.hyperlinks
    assert(links.key?("A1"), "Hyperlink on A1 should exist [hyperlinks]")
  ensure
    tmp&.close!
  end

  # =====================================================
  # Auto Filter CONTRACT tests
  # =====================================================

  data(API_PATHS)
  test "autofilter: range is preserved" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Data") do |s|
        s.add_row(%w[Name Score])
        s.add_row(["Alice", 95])
        s.add_row(["Bob", 87])
        s.set_auto_filter("A1:B3")
      end
    end

    af = reader.auto_filter
    assert_equal("A1:B3", af, "Auto filter range mismatch [auto_filter]")
  ensure
    tmp&.close!
  end

  # =====================================================
  # Data Validation CONTRACT tests
  # =====================================================

  data(API_PATHS)
  test "data_validation: rule is preserved" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("DV") do |s|
        s.add_row(["Value"])
        s.add_data_validation("A2:A100", type: :whole, formula1: "1", formula2: "100")
      end
    end

    dvs = reader.data_validations
    assert_equal(1, dvs.size, "Expected 1 data validation [data_validation count]")
    assert_equal("A2:A100", dvs[0][:sqref], "Data validation sqref mismatch")
  ensure
    tmp&.close!
  end

  # =====================================================
  # Conditional Formatting CONTRACT tests
  # =====================================================

  data(API_PATHS)
  test "conditional_format: rule is preserved" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("CF") do |s|
        s.add_row([10, 20, 30])
        s.add_conditional_format("A1:C1", type: :cell_is, operator: :greaterThan, formula: "15", priority: 1)
      end
    end

    cfs = reader.conditional_formats
    assert_equal(1, cfs.size, "Expected 1 conditional format [conditional_format count]")
    assert_equal("A1:C1", cfs[0][:sqref])
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "conditional_format: visual style emits dxf linkage" do |api_path|
    _reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("CF") do |s|
        s.add_row([90, 45, 72, 88])
        s.add_conditional_format("A1:D1",
                                 type: :cell_is,
                                 operator: :greaterThan,
                                 formula: "80",
                                 priority: 1,
                                 fill_color: "FFFFC7CE")
      end
    end

    entries = Xlsxrb::Ooxml::ZipReader.open(tmp.path, &:read_all)
    sheet_xml = entries["xl/worksheets/sheet1.xml"]
    styles_xml = entries["xl/styles.xml"]

    assert_not_nil(sheet_xml, "sheet1.xml should exist")
    assert_not_nil(styles_xml, "styles.xml should exist")
    assert_match(/cfRule[^>]*dxfId="0"/, sheet_xml,
                 "Expected cfRule to reference dxfId [conditional_format dxf linkage]")
    assert_match(/<dxfs count="1">/, styles_xml,
                 "Expected styles.xml to include one dxf [styles dxfs]")
  ensure
    tmp&.close!
  end

  # =====================================================
  # Table CONTRACT tests
  # =====================================================

  data(API_PATHS)
  test "table: definition is preserved" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Tables") do |s|
        s.add_row(%w[Name Score])
        s.add_row(["Alice", 95])
        s.add_table("A1:B2", columns: %w[Name Score], name: "TestTable")
      end
    end

    tables = reader.tables
    assert_equal(1, tables.size, "Expected 1 table [table count]")
    assert_equal("TestTable", tables[0][:name], "Table name mismatch")
    assert_equal("A1:B2", tables[0][:ref], "Table ref mismatch")
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "table: style string emits table part and valid style info" do |api_path|
    _reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Tables") do |s|
        s.add_row(%w[Name Score])
        s.add_row(["Alice", 95])
        s.add_table("A1:B2",
                    columns: %w[Name Score],
                    name: "StyledTable",
                    style: "TableStyleMedium9",
                    show_first_column: true,
                    show_row_stripes: false)
      end
    end

    entries = Xlsxrb::Ooxml::ZipReader.open(tmp.path, &:read_all)
    assert_not_nil(entries["xl/worksheets/sheet1.xml"], "Worksheet part should exist")
    tbl_xml = entries["xl/tables/table1.xml"]
    assert_not_nil(tbl_xml, "Table part should exist")
    assert_match(/tableStyleInfo[^>]*name="TableStyleMedium9"/, tbl_xml)
    assert_match(/showFirstColumn="1"/, tbl_xml)
    assert_match(/showRowStripes="0"/, tbl_xml)
  ensure
    tmp&.close!
  end

  # =====================================================
  # Merge Cells CONTRACT tests
  # =====================================================

  data(API_PATHS)
  test "merge_cells: range is preserved" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Merge") do |s|
        s.add_row(["Merged", nil, nil])
        s.merge_cells("A1:C1")
      end
    end

    merged = reader.merged_cells
    assert_equal(["A1:C1"], merged, "Merged cells mismatch [merged_cells]")
  ensure
    tmp&.close!
  end

  # =====================================================
  # Freeze Pane CONTRACT tests
  # =====================================================

  data(API_PATHS)
  test "freeze_pane: settings are preserved" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Frozen") do |s|
        s.add_row(["Header"])
        s.add_row([1])
        s.set_freeze_pane(row: 1, col: 0)
      end
    end

    pane = reader.freeze_pane
    assert_not_nil(pane, "Freeze pane should exist [freeze_pane]")
    assert_equal(:frozen, pane[:state], "Freeze pane state should be frozen")
    assert_equal(1, pane[:row], "Freeze pane row mismatch")
  ensure
    tmp&.close!
  end

  # =====================================================
  # Page Margins CONTRACT tests
  # =====================================================

  data(API_PATHS)
  test "page_margins: values are preserved" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Margins") do |s|
        s.add_row(["Data"])
        s.set_page_margins(left: 1.0, right: 1.0, top: 1.5, bottom: 1.5)
      end
    end

    margins = reader.page_margins
    assert_not_nil(margins, "Page margins should exist [page_margins]")
    assert_in_delta(1.0, margins[:left], 0.01, "Left margin mismatch")
    assert_in_delta(1.5, margins[:top], 0.01, "Top margin mismatch")
  ensure
    tmp&.close!
  end

  # =====================================================
  # Page Setup CONTRACT tests
  # =====================================================

  data(API_PATHS)
  test "page_setup: orientation is preserved" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Setup") do |s|
        s.add_row(["Data"])
        s.set_page_setup(orientation: :landscape)
      end
    end

    ps = reader.page_setup
    assert_not_nil(ps, "Page setup should exist [page_setup]")
    assert_equal("landscape", ps[:orientation], "Orientation mismatch")
  ensure
    tmp&.close!
  end

  # =====================================================
  # Header/Footer CONTRACT tests
  # =====================================================

  data(API_PATHS)
  test "header_footer: text is preserved" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("HF") do |s|
        s.add_row(["Data"])
        s.set_header_footer(odd_header: "&CReport", odd_footer: "&CPage &P")
      end
    end

    hf = reader.header_footer
    assert_not_nil(hf, "Header/footer should exist [header_footer]")
    assert_equal("&CReport", hf[:odd_header], "Header mismatch")
    assert_equal("&CPage &P", hf[:odd_footer], "Footer mismatch")
  ensure
    tmp&.close!
  end

  # =====================================================
  # Sheet Protection CONTRACT tests
  # =====================================================

  data(API_PATHS)
  test "sheet_protection: settings are preserved" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Prot") do |s|
        s.add_row(["Data"])
        s.set_sheet_protection(sheet: true, objects: true)
      end
    end

    prot = reader.sheet_protection
    assert_not_nil(prot, "Sheet protection should exist [sheet_protection]")
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "sheet_protection: plain password is hashed" do |api_path|
    _reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Prot") do |s|
        s.add_row(["Data"])
        s.set_sheet_protection(sheet: true, password: "secret")
      end
    end

    entries = Xlsxrb::Ooxml::ZipReader.open(tmp.path, &:read_all)
    sheet_xml = entries["xl/worksheets/sheet1.xml"]
    assert_not_nil(sheet_xml, "sheet1.xml should exist")
    assert_match(/algorithmName="SHA-512"/, sheet_xml,
                 "Expected SHA-512 protection hash metadata [sheet_protection hash] ")
    assert_match(%r{hashValue="[A-Za-z0-9+/]+=*"}, sheet_xml,
                 "Expected hashValue in sheet protection")
    assert_match(%r{saltValue="[A-Za-z0-9+/]+=*"}, sheet_xml,
                 "Expected saltValue in sheet protection")
    assert_match(/spinCount="100000"/, sheet_xml,
                 "Expected default spinCount in sheet protection")
    assert_not_match(/password="secret"/, sheet_xml,
                     "Plain-text password must not be emitted")
  ensure
    tmp&.close!
  end

  # =====================================================
  # Comments CONTRACT tests
  # =====================================================

  data(API_PATHS)
  test "comment: text is preserved" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Comments") do |s|
        s.add_row(["Value"])
        s.add_comment("A1", "Test comment", author: "Author")
      end
    end

    comments = reader.comments
    assert_equal(1, comments.size, "Expected 1 comment [comment count]")
    assert_equal("A1", comments[0][:ref], "Comment ref mismatch")
  ensure
    tmp&.close!
  end

  # =====================================================
  # Combined features CONTRACT test
  # =====================================================

  data(API_PATHS)
  test "combined: multiple features on same sheet" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("All") do |s|
        s.add_row(%w[Name Score])
        s.add_row(["Alice", 95])
        s.set_auto_filter("A1:B2")
        s.merge_cells("A1:A1")
        s.set_freeze_pane(row: 1)
        s.set_page_margins(left: 1.0, right: 1.0)
        s.add_data_validation("B2", type: :whole, formula1: "0", formula2: "100")
      end
    end

    assert_equal("A1:B2", reader.auto_filter)
    assert_equal(["A1:A1"], reader.merged_cells)
    assert_not_nil(reader.freeze_pane)
    assert_not_nil(reader.page_margins)
    assert_equal(1, reader.data_validations.size)
  ensure
    tmp&.close!
  end

  # =====================================================
  # Missing Facade CONTRACT tests
  # =====================================================

  data(API_PATHS)
  test "worksheet: column and row formatting" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Sheet1") do |s|
        s.set_column(0, width: 20.0, hidden: true, outline_level: 1)
        s.add_row(["Test"], height: 30.0, hidden: true, outline_level: 2)
      end
    end

    cols = reader.column_attributes
    assert_equal(1, cols.size, "Expected 1 column attribute")
    assert_in_delta(20.0, reader.columns["A"], 0.01)
    assert_equal(true, cols["A"][:hidden])
    assert_equal(1, cols["A"][:outline_level])

    rows = reader.row_attributes
    assert_equal(1, rows.size, "Expected 1 row attribute")
    assert_in_delta(30.0, rows[1][:height], 0.01)
    assert_equal(true, rows[1][:hidden])
    assert_equal(2, rows[1][:outline_level])
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "worksheet: split pane and selection" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Sheet1") do |s|
        s.add_row(["Data"])
        s.set_split_pane(x_split: 1, y_split: 2, top_left_cell: "B3")
        s.set_selection("B3", sqref: "B3:C4", pane: :bottomRight)
      end
    end

    pane = reader.freeze_pane
    assert_not_nil(pane)
    assert_equal(:split, pane[:state])
    assert_equal(1, pane[:col])
    assert_equal(2, pane[:row])
    assert_equal("B3", pane[:top_left_cell])

    sel = reader.selection
    assert_not_nil(sel)
    # Selection comes as hash
    assert_equal("B3", sel[:active_cell])
    assert_equal("B3:C4", sel[:sqref])
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "worksheet: sheet properties and views" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Sheet1") do |s|
        s.add_row(["Data"])
        s.set_sheet_property(:tab_color, "FF0000")
        s.set_sheet_view(:show_grid_lines, false)
        s.set_sheet_view(:zoom_scale, 150)
      end
    end

    props = reader.sheet_properties
    assert_equal("FF0000", props[:tab_color])

    views = reader.sheet_view
    assert_equal(false, views[:show_grid_lines])
    assert_equal(150, views[:zoom_scale])
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "workbook: defined names" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Sheet1") do |s|
        s.add_row(["Data"])
      end
      w.add_defined_name("MyConstant", "100")
      w.add_defined_name("LocalName", "Sheet1!$A$1", sheet: "Sheet1")
    end

    names = reader.defined_names
    assert_equal(2, names.size)
    assert_not_nil(names.find { |n| n[:name] == "MyConstant" && n[:value] == "100" })
    assert_not_nil(names.find { |n| n[:name] == "LocalName" && n[:local_sheet_id].zero? })
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "worksheet: print area and titles" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Sheet1") do |s|
        s.add_row(["Data"])
      end
      w.set_print_area("Sheet1!$A$1:$C$10", sheet: "Sheet1")
      w.set_print_titles(rows: "1:2", cols: "A:B", sheet: "Sheet1")
    end

    assert_equal("Sheet1!$A$1:$C$10", reader.print_area(sheet: "Sheet1"))
    # In Writer, ranges are emitted in a consistent order but potentially with quotes
    # The actual output might be "'Sheet1'!$A:$B,'Sheet1'!$1:$2"
    assert_match(/Sheet1[^,]+,.*Sheet1/, reader.print_titles(sheet: "Sheet1"))
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "worksheet: print options" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Sheet1") do |s|
        s.add_row(["Data"])
        s.set_print_option(:grid_lines, true)
        s.set_print_option(:horizontal_centered, true)
      end
    end

    opts = reader.print_options
    assert_equal(true, opts[:grid_lines])
    assert_equal(true, opts[:horizontal_centered])
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "worksheet: row and col breaks" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Sheet1") do |s|
        s.add_row(["Data"])
        s.add_row_break(1)
        s.add_col_break(2)
      end
    end

    r_breaks = reader.row_breaks
    assert_equal([1], r_breaks.map { |b| b[:id] })
    c_breaks = reader.col_breaks
    assert_equal([2], c_breaks.map { |b| b[:id] })
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "worksheet: sort state" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Sheet1") do |s|
        s.add_row(["Data"])
        s.set_auto_filter("A1:A10")
        s.set_sort_state("A1:A10", [{ ref: "A1:A10", descending: true }])
      end
    end

    sort = reader.sort_state
    assert_not_nil(sort, "Sort state should be written with auto_filter")
    assert_equal("A1:A10", sort[:ref])
    assert_equal(1, sort[:sort_conditions].size)
    assert_equal("A1:A10", sort[:sort_conditions][0][:ref])
    assert_equal(true, sort[:sort_conditions][0][:descending])
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "worksheet: filter columns" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Sheet1") do |s|
        s.add_row(["Data"])
        s.set_auto_filter("A1:A10")
        s.add_filter_column(0, type: :filters, values: ["Data"])
      end
    end

    cols = reader.filter_columns
    assert_equal(1, cols.size)
    assert_equal(:filters, cols[0][:type])
    assert_equal(["Data"], cols[0][:values])
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "workbook: protection" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Sheet1") do |s|
        s.add_row(["Data"])
      end
      w.set_workbook_protection(lock_structure: true)
    end

    prot = reader.workbook_protection
    assert_equal(true, prot[:lock_structure])
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "workbook: core and app properties" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Sheet1") do |s|
        s.add_row(["Data"])
      end
      w.set_core_property(:creator, "Author")
      w.set_app_property(:application, "Xlsxrb")
      w.add_custom_property("MyProp", "MyValue")
    end

    assert_equal("Author", reader.core_properties[:creator])
    assert_equal("Xlsxrb", reader.app_properties[:application])
    prop = reader.custom_properties.find { |p| p[:name] == "MyProp" }
    assert_not_nil(prop)
    assert_equal("MyValue", prop[:value])
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "worksheet: add image" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Sheet1") do |s|
        s.add_row(["Data"])
        s.add_image(MINIMAL_PNG, ext: "png", from_col: 1, from_row: 1)
      end
    end

    images = reader.images
    assert_equal(1, images.size)
    assert_equal(1, images[0][:from_col])
    assert_equal(1, images[0][:from_row])
  ensure
    tmp&.close!
  end

  # =====================================================
  # Cell Style CONTRACT tests
  # =====================================================

  data(API_PATHS)
  test "style: bold font and fill color are reflected in cell_styles" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Styled") do |s|
        s.add_style("bold_green") do |style|
          style.bold.fill_color("00FF00")
        end
        s.add_row(["Header"], styles: ["bold_green"])
        s.add_row(["Normal"])
      end
    end

    cell_styles = reader.cell_styles
    a1 = cell_styles["A1"]
    assert_not_nil(a1, "Cell A1 should have style info [cell_styles]. " \
                       "Check StyleBuilder#bold and WorkbookWriter#build_styles_xml font output.")
    assert_not_nil(a1[:font], "A1 should have font entry [cell_styles.font]")
    assert_equal(true, a1[:font][:bold],
                 "A1 font should be bold [cell_styles.font.bold]. " \
                 "Check Writer#add_font bold handling and StylesParser bold element.")
    assert_not_nil(a1[:fill], "A1 should have fill entry [cell_styles.fill]")
    fg = a1[:fill][:fg_color]
    assert_not_nil(fg, "A1 fill fg_color should be present [cell_styles.fill.fg_color]")
    assert_equal("00FF00", fg,
                 "A1 fill fg_color mismatch [cell_styles.fill.fg_color]. " \
                 "Check StyleBuilder#fill_color and WorkbookWriter fgColor rgb attribute.")
    assert_nil(cell_styles["A2"],
               "Unstyled cell A2 should not appear in cell_styles")
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "style: number format is reflected in cell_formats" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Formats") do |s|
        # numFmtId 4 = "#,##0.00" (built-in)
        s.add_style("currency") { |style| style.number_format(4) }
        s.add_row([1234.56], styles: ["currency"])
        s.add_row([7890])
      end
    end

    fmts = reader.cell_formats
    assert_not_nil(fmts["A1"],
                   "Cell A1 should have a format code [cell_formats]. " \
                   "Check StyleBuilder#number_format and Reader#cell_formats num_fmt_id resolution.")
    assert_match(/#,##0/, fmts["A1"],
                 "Expected '#,##0.00' style format for numFmtId 4 [cell_formats.A1].")
    assert_nil(fmts["A2"],
               "Unstyled cell A2 should not appear in cell_formats")
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "style: border is written to styles.xml" do |api_path|
    _reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Bordered") do |s|
        s.add_style("boxed") { |style| style.border_all(style: "thin") }
        s.add_row(["Item"], styles: ["boxed"])
      end
    end

    entries = Xlsxrb::Ooxml::ZipReader.open(tmp.path, &:read_all)
    styles_xml = entries["xl/styles.xml"]
    assert_not_nil(styles_xml, "styles.xml should exist")
    assert_match(/style="thin"/, styles_xml,
                 "styles.xml should contain thin border style [border]. " \
                 "Check StyleBuilder#border_all and WorkbookWriter border emission.")
  ensure
    tmp&.close!
  end

  # =====================================================
  # Date cell value CONTRACT tests
  # =====================================================

  data(API_PATHS)
  test "cell: Date value with date numFmt style is resolved by Reader#cells" do |api_path|
    today = Date.today
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Dates") do |s|
        # numFmtId 14 = "m/d/yy" (built-in date format); triggers date conversion in Reader
        s.add_style("date_fmt") { |style| style.number_format(14) }
        s.add_row([today], styles: ["date_fmt"])
      end
    end

    # reader.cells resolves dates using numFmt from styles.xml
    cells = reader.cells
    assert_instance_of(Date, cells["A1"],
                       "Expected Date value for date-formatted cell via reader.cells [cell type]. " \
                       "Check WorksheetWriter#write_row_values Date serialization " \
                       "(Utils.date_to_serial) and Reader#resolve_date_cells numFmt detection.")
    assert_equal(today, cells["A1"], "Date value mismatch [cell value]")
  ensure
    tmp&.close!
  end

  # =====================================================
  # Shape round-trip CONTRACT tests
  # =====================================================

  data(API_PATHS)
  test "shape: reader.shapes returns correct preset, text, and anchor positions" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Shapes") do |s|
        s.add_row(["Data"])
        s.add_shape(preset: "roundRect", text: "Call Out",
                    from_col: 1, from_row: 2, to_col: 4, to_row: 6)
      end
    end

    shapes = reader.shapes
    assert_equal(1, shapes.size,
                 "Expected 1 shape [shape count]. Check WorkbookWriter#generate_drawing_xml shape element.")
    shape = shapes[0]
    assert_equal("roundRect", shape[:preset],
                 "Shape preset mismatch [preset]. Check generate_drawing_xml prstGeom prst attribute.")
    assert_equal("Call Out", shape[:text],
                 "Shape text mismatch [text]. Check generate_drawing_xml txBody paragraph text.")
    assert_equal(1, shape[:from_col], "Shape from_col mismatch [from_col]")
    assert_equal(2, shape[:from_row], "Shape from_row mismatch [from_row]")
    assert_equal(4, shape[:to_col], "Shape to_col mismatch [to_col]")
    assert_equal(6, shape[:to_row], "Shape to_row mismatch [to_row]")
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "shape: multiple shapes on same sheet are all returned" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Multi") do |s|
        s.add_row(["Data"])
        s.add_shape(preset: "rect", text: "Box", from_col: 0, from_row: 1, to_col: 2, to_row: 3)
        s.add_shape(preset: "ellipse", text: "Oval", from_col: 3, from_row: 1, to_col: 5, to_row: 3)
      end
    end

    shapes = reader.shapes
    assert_equal(2, shapes.size,
                 "Expected 2 shapes [shape count]. Check WorkbookWriter drawing loop.")
    presets = shapes.map { |s| s[:preset] }.sort
    assert_equal(%w[ellipse rect], presets,
                 "Shape preset types mismatch [shape presets].")
  ensure
    tmp&.close!
  end

  # =====================================================
  # Hyperlink metadata CONTRACT tests
  # =====================================================

  data(API_PATHS)
  test "hyperlink: display text is preserved" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Links") do |s|
        s.add_row(["Click here"])
        s.add_hyperlink("A1", "https://example.com", display: "Visit Example")
      end
    end

    links = reader.hyperlinks
    assert(links.key?("A1"), "Hyperlink on A1 should exist")
    link = links["A1"]
    assert_equal("Visit Example", link[:display],
                 "Hyperlink display text mismatch [display]. " \
                 "Check WorkbookWriter hyperlink display attribute and Reader HyperlinksListener.")
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "hyperlink: internal location link is preserved" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Nav") do |s|
        s.add_row(["Jump"])
        s.add_hyperlink("A1", location: "Nav!B2", display: "Go to B2")
      end
    end

    links = reader.hyperlinks
    assert(links.key?("A1"),
           "Internal location hyperlink on A1 should exist [hyperlinks]. " \
           "Check WorksheetWriter write_hyperlinks location attribute.")
    link = links["A1"]
    # Internal links have no external URL rId; they appear via location attribute
    assert(link[:location] || link[:url].nil?,
           "Internal hyperlink should carry location or have no external URL [hyperlink.location].")
  ensure
    tmp&.close!
  end

  # =====================================================
  # Data Validation extended CONTRACT tests
  # =====================================================

  data(API_PATHS)
  test "data_validation: list type is preserved" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("DV") do |s|
        s.add_row(%w[Name Status])
        s.add_data_validation("B2:B100", type: :list, formula1: '"Active,Inactive,Pending"')
      end
    end

    dvs = reader.data_validations
    assert_equal(1, dvs.size, "Expected 1 data validation [count]")
    dv = dvs[0]
    assert_equal("B2:B100", dv[:sqref], "Data validation sqref mismatch")
    assert_equal("list", dv[:type].to_s,
                 "Expected list type [type]. Check WorksheetWriter#write_data_validations type attribute.")
    assert_match(/Active/, dv[:formula1].to_s,
                 "Expected formula1 to contain list values [formula1].")
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "data_validation: multiple rules on same sheet are all preserved" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("DV") do |s|
        s.add_row(%w[Name Score Grade])
        s.add_data_validation("B2:B100", type: :whole, formula1: "0", formula2: "100")
        s.add_data_validation("C2:C100", type: :list, formula1: '"A,B,C,D,F"')
      end
    end

    dvs = reader.data_validations
    assert_equal(2, dvs.size,
                 "Expected 2 data validations [count]. " \
                 "Check WorksheetWriter#write_data_validations multiple rule handling.")
    sqrefs = dvs.map { |d| d[:sqref] }.sort
    assert_equal(%w[B2:B100 C2:C100], sqrefs,
                 "Data validation sqrefs mismatch")
    types = dvs.map { |d| d[:type].to_s }.sort
    assert_equal(%w[list whole], types,
                 "Data validation types mismatch")
  ensure
    tmp&.close!
  end

  # =====================================================
  # Merge Cells extended CONTRACT tests
  # =====================================================

  data(API_PATHS)
  test "merge_cells: multiple ranges are all preserved" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Merge") do |s|
        s.add_row(["Title", nil, nil, "Sub", nil])
        s.add_row(["Content"])
        s.merge_cells("A1:C1")
        s.merge_cells("D1:E1")
        s.merge_cells("A2:E2")
      end
    end

    merged = reader.merged_cells
    assert_equal(3, merged.size,
                 "Expected 3 merged ranges [merge count]. " \
                 "Check WorksheetWriter#write_merge_cells multiple mergeCell elements.")
    assert_include(merged, "A1:C1", "A1:C1 merge should exist")
    assert_include(merged, "D1:E1", "D1:E1 merge should exist")
    assert_include(merged, "A2:E2", "A2:E2 merge should exist")
  ensure
    tmp&.close!
  end

  # =====================================================
  # Comment extended CONTRACT tests
  # =====================================================

  data(API_PATHS)
  test "comment: text and author are preserved" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Notes") do |s|
        s.add_row(["Value"])
        s.add_comment("A1", "This is important!", author: "Alice")
      end
    end

    comments = reader.comments
    assert_equal(1, comments.size, "Expected 1 comment [comment count]")
    c = comments[0]
    assert_equal("A1", c[:ref], "Comment ref mismatch")
    assert_equal("This is important!", c[:text],
                 "Comment text mismatch [text]. Check WorkbookWriter#generate_comments_xml t element.")
    assert_equal("Alice", c[:author],
                 "Comment author mismatch [author]. Check generate_comments_xml authors element.")
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "comment: multiple comments on same sheet are all preserved" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Notes") do |s|
        s.add_row(%w[A B])
        s.add_comment("A1", "Note on A1", author: "Alice")
        s.add_comment("B1", "Note on B1", author: "Bob")
      end
    end

    comments = reader.comments
    assert_equal(2, comments.size,
                 "Expected 2 comments [comment count]. Check WorkbookWriter#generate_comments_xml.")
    refs = comments.map { |c| c[:ref] }.sort
    assert_equal(%w[A1 B1], refs, "Comment refs mismatch")
    authors = comments.map { |c| c[:author] }.sort
    assert_equal(%w[Alice Bob], authors, "Comment authors mismatch")
  ensure
    tmp&.close!
  end

  # =====================================================
  # Table columns CONTRACT tests
  # =====================================================

  data(API_PATHS)
  test "table: columns array is preserved with correct names" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("People") do |s|
        s.add_row(%w[Name Age City])
        s.add_row(["Alice", 30, "NYC"])
        s.add_table("A1:C2", columns: %w[Name Age City], name: "PeopleTable")
      end
    end

    tables = reader.tables
    assert_equal(1, tables.size, "Expected 1 table")
    tbl = tables[0]
    assert_equal("PeopleTable", tbl[:name])
    assert_equal("A1:C2", tbl[:ref])
    assert_not_nil(tbl[:columns],
                   "Table columns should be present [columns]. " \
                   "Check WorkbookWriter#build_table_xml tableColumn elements.")
    col_names = tbl[:columns].map { |c| c[:name] }.sort
    assert_equal(%w[Age City Name], col_names,
                 "Table column names mismatch [columns.name].")
  ensure
    tmp&.close!
  end

  # =====================================================
  # IO read CONTRACT tests
  # =====================================================

  test "io: Xlsxrb.read accepts an IO object (StringIO)" do
    tmp = Tempfile.new(["contract_io", ".xlsx"])
    Xlsxrb.generate(tmp.path) do |w|
      w.add_sheet("IOTest") do |s|
        s.add_row(["io_string", 999])
      end
    end

    raw = File.binread(tmp.path)
    sio = StringIO.new(raw)
    wb = Xlsxrb.read(sio)
    assert_instance_of(Xlsxrb::Elements::Workbook, wb,
                       "Xlsxrb.read(StringIO) should return a Workbook")
    sheet = wb.sheet(0)
    assert_not_nil(sheet, "Sheet 0 should exist when reading from IO")
    assert_equal("io_string", sheet.cell_value("A1"),
                 "String cell value mismatch when reading from StringIO [A1]")
    assert_equal(999, sheet.cell_value("B1"),
                 "Numeric cell value mismatch when reading from StringIO [B1]")
  ensure
    tmp&.close!
  end

  # =====================================================
  # foreach CONTRACT tests
  # =====================================================

  test "foreach: streaming read yields all rows in order" do
    tmp = Tempfile.new(["contract_foreach", ".xlsx"])
    Xlsxrb.generate(tmp.path) do |w|
      w.add_sheet("Seq") do |s|
        s.add_row([1, "one"])
        s.add_row([2, "two"])
        s.add_row([3, "three"])
      end
    end

    collected = []
    Xlsxrb.foreach(tmp.path) do |row|
      assert_instance_of(Xlsxrb::Elements::Row, row,
                         "foreach should yield Elements::Row instances")
      collected << row.cells.map(&:value)
    end

    assert_equal(3, collected.size,
                 "foreach should yield exactly 3 rows [foreach count]. " \
                 "Check Xlsxrb.foreach streaming mechanics.")
    assert_equal([1, "one"], collected[0], "Row 1 values mismatch [foreach row 0]")
    assert_equal([2, "two"], collected[1], "Row 2 values mismatch [foreach row 1]")
    assert_equal([3, "three"], collected[2], "Row 3 values mismatch [foreach row 2]")
  ensure
    tmp&.close!
  end

  test "foreach: without block returns Enumerator" do
    tmp = Tempfile.new(["contract_foreach_enum", ".xlsx"])
    Xlsxrb.generate(tmp.path) do |w|
      w.add_sheet("E") do |s|
        s.add_row(["alpha"])
        s.add_row(["beta"])
      end
    end

    enum = Xlsxrb.foreach(tmp.path)
    assert_respond_to(enum, :each,
                      "foreach without block should return an Enumerator-like object")
    values = enum.map { |row| row.cells[0]&.value }
    assert_equal(%w[alpha beta], values,
                 "Enumerator should yield correct values [foreach enumerator]")
  ensure
    tmp&.close!
  end

  test "foreach: sheet keyword argument selects the named sheet" do
    tmp = Tempfile.new(["contract_foreach_sheet", ".xlsx"])
    Xlsxrb.generate(tmp.path) do |w|
      w.add_sheet("First") { |s| s.add_row(["from_first"]) }
      w.add_sheet("Second") { |s| s.add_row(["from_second"]) }
    end

    collected = []
    Xlsxrb.foreach(tmp.path, sheet: "Second") { |row| collected << row.cells[0]&.value }
    assert_equal(["from_second"], collected,
                 "foreach with sheet: keyword should read the named sheet [foreach sheet]. " \
                 "Check Xlsxrb.foreach sheet resolution logic.")
  ensure
    tmp&.close!
  end

  # =====================================================
  # Additional cell data CONTRACT tests
  # =====================================================

  data(API_PATHS)
  test "cell: empty string is preserved" do |api_path|
    _, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Empty") do |s|
        s.add_row(["", "non-empty"])
      end
    end

    wb = Xlsxrb.read(tmp.path)
    assert_equal("", wb.sheet(0).cell_value("A1"),
                 "Empty string cell should round-trip as empty string [A1]")
    assert_equal("non-empty", wb.sheet(0).cell_value("B1"),
                 "Adjacent non-empty cell should be unaffected [B1]")
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "cell: sparse row with nil gaps preserves non-nil values" do |api_path|
    _, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Sparse") do |s|
        s.add_row(["first", nil, nil, "fourth"])
      end
    end

    wb = Xlsxrb.read(tmp.path)
    sheet = wb.sheet(0)
    assert_equal("first", sheet.cell_value("A1"), "A1 should be 'first'")
    assert_nil(sheet.cell_value("B1"), "B1 should be nil (sparse)")
    assert_nil(sheet.cell_value("C1"), "C1 should be nil (sparse)")
    assert_equal("fourth", sheet.cell_value("D1"), "D1 should be 'fourth'")
  ensure
    tmp&.close!
  end

  # =====================================================
  # Workbook-level Zip structure CONTRACT tests
  # =====================================================

  data(API_PATHS)
  test "zip: required parts exist in generated file" do |api_path|
    _, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Sheet1") { |s| s.add_row(["data"]) }
    end

    entries = Xlsxrb::Ooxml::ZipReader.open(tmp.path, &:read_all)
    %w[
      [Content_Types].xml
      _rels/.rels
      xl/workbook.xml
      xl/_rels/workbook.xml.rels
      xl/styles.xml
      xl/worksheets/sheet1.xml
    ].each do |required_part|
      assert(entries.key?(required_part),
             "Required ZIP part '#{required_part}' should exist. " \
             "Check WorkbookWriter#write_to entry generation.")
    end
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "zip: shared strings part exists when strings are present" do |api_path|
    _, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("S") { |s| s.add_row(["hello"]) }
    end

    entries = Xlsxrb::Ooxml::ZipReader.open(tmp.path, &:read_all)
    assert(entries.key?("xl/sharedStrings.xml"),
           "xl/sharedStrings.xml should exist when string cells are present. " \
           "Check WorkbookWriter#write_to sharedStrings generation.")
    assert_match(/hello/, entries["xl/sharedStrings.xml"],
                 "sharedStrings.xml should contain the cell string value.")
  ensure
    tmp&.close!
  end

  data(API_PATHS)
  test "content_types: sheet1 content type is declared" do |api_path|
    _, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Sheet1") { |s| s.add_row([1]) }
    end

    entries = Xlsxrb::Ooxml::ZipReader.open(tmp.path, &:read_all)
    ct = entries["[Content_Types].xml"]
    assert_match(/worksheet/, ct,
                 "[Content_Types].xml should declare worksheet content type. " \
                 "Check WorkbookWriter#build_content_types.")
  ensure
    tmp&.close!
  end

  # =====================================================
  # Page setup extended CONTRACT tests
  # =====================================================

  data(API_PATHS)
  test "page_setup: paper_size is preserved" do |api_path|
    reader, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Setup") do |s|
        s.add_row(["Data"])
        s.set_page_setup(orientation: :landscape, paper_size: 9)
      end
    end

    ps = reader.page_setup
    assert_not_nil(ps, "Page setup should exist")
    assert_equal("landscape", ps[:orientation], "Orientation mismatch")
    assert_equal(9, ps[:paper_size],
                 "Paper size mismatch [paper_size]. Check WorksheetWriter#write_page_setup paperSize attribute.")
  ensure
    tmp&.close!
  end

  # =====================================================
  # Sheet name CONTRACT tests
  # =====================================================

  data(API_PATHS)
  test "workbook: reader.sheet_names returns all sheet names in order" do |api_path|
    _, tmp = generate_and_read(api_path) do |w|
      w.add_sheet("Alpha") { |s| s.add_row([1]) }
      w.add_sheet("Beta") { |s| s.add_row([2]) }
      w.add_sheet("Gamma") { |s| s.add_row([3]) }
    end

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    names = reader.sheet_names
    assert_equal(%w[Alpha Beta Gamma], names,
                 "Sheet names should be returned in order [sheet_names]. " \
                 "Check WorkbookWriter workbook.xml sheet ordering.")
  ensure
    tmp&.close!
  end
end
