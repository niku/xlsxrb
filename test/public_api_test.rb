# frozen_string_literal: true

require "test_helper"
require "tempfile"

class PublicApiTest < Test::Unit::TestCase
  # --- Read / Write ---

  test "Xlsxrb.read returns a Workbook from a written file" do
    tmp = Tempfile.new(["facade_test", ".xlsx"])
    begin
      cell_a1 = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "Hello")
      cell_b1 = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 1, value: 42)
      cell_a2 = Xlsxrb::Elements::Cell.new(row_index: 1, column_index: 0, value: true)
      cell_b2 = Xlsxrb::Elements::Cell.new(row_index: 1, column_index: 1, value: 3.14)

      row1 = Xlsxrb::Elements::Row.new(index: 0, cells: [cell_a1, cell_b1])
      row2 = Xlsxrb::Elements::Row.new(index: 1, cells: [cell_a2, cell_b2])
      ws = Xlsxrb::Elements::Worksheet.new(name: "TestSheet", rows: [row1, row2])
      wb = Xlsxrb::Elements::Workbook.new(sheets: [ws])

      Xlsxrb.write(tmp.path, wb)

      result = Xlsxrb.load(tmp.path)

      assert_instance_of(Xlsxrb::Elements::Workbook, result)
      assert_equal(1, result.sheets.size)
      assert_equal("TestSheet", result.sheets[0].name)
      assert_equal(2, result.sheets[0].rows.size)

      sheet = result.sheet(0)
      assert_equal("Hello", sheet["A1"]&.value)
      assert_equal(42, sheet["B1"]&.value)
      assert_equal(true, sheet["A2"]&.value)
      assert_in_delta(3.14, sheet["B2"]&.value)
    ensure
      tmp.close!
    end
  end

  test "Xlsxrb.write returns binary string when target is omitted" do
    ws = Xlsxrb::Elements::Worksheet.new(name: "S", rows: [
                                           Xlsxrb::Elements::Row.new(index: 0, cells: [
                                                                       Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "In-Memory")
                                                                     ])
                                         ])
    wb = Xlsxrb::Elements::Workbook.new(sheets: [ws])

    binary = Xlsxrb.write(wb)
    assert_kind_of(String, binary)
    assert(binary.start_with?("PK\x03\x04"))
    assert_equal(Encoding::ASCII_8BIT, binary.encoding)

    # Validate that it can be parsed back
    parsed = Xlsxrb.load(binary)
    assert_equal(1, parsed.sheets.size)
    assert_equal("In-Memory", parsed.sheet(0)["A1"]&.value)
  end

  test "Xlsxrb.read automatically handles raw binary string" do
    wb = Xlsxrb.build do |b|
      b.sheet("Users") do |s|
        s.row(["Alice", 30])
      end
    end
    binary = Xlsxrb.write(wb)

    parsed = Xlsxrb.load(binary)
    assert_equal(1, parsed.sheets.size)
    assert_equal("Alice", parsed.sheet("Users")["A1"]&.value)
  end

  test "Xlsxrb.write raises on nil target when two arguments are given" do
    ws = Xlsxrb::Elements::Worksheet.new(name: "S")
    wb = Xlsxrb::Elements::Workbook.new(sheets: [ws])
    assert_raise(Xlsxrb::Error) { Xlsxrb.write(nil, wb) }
  end

  test "Xlsxrb.write raises on non-workbook" do
    assert_raise(Xlsxrb::Error) { Xlsxrb.write("/tmp/test.xlsx", "not a workbook") }
    assert_raise(Xlsxrb::Error) { Xlsxrb.write("not a workbook") }
  end

  test "round-trip preserves numeric types" do
    tmp = Tempfile.new(["numeric_rt", ".xlsx"])
    begin
      cells = [
        Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: 0),
        Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 1, value: -99),
        Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 2, value: 1_000_000),
        Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 3, value: 1.5)
      ]
      row = Xlsxrb::Elements::Row.new(index: 0, cells: cells)
      ws = Xlsxrb::Elements::Worksheet.new(name: "Numbers", rows: [row])
      wb = Xlsxrb::Elements::Workbook.new(sheets: [ws])

      Xlsxrb.write(tmp.path, wb)
      result = Xlsxrb.load(tmp.path)

      sheet = result.sheet(0)
      assert_equal(0, sheet["A1"]&.value)
      assert_equal(-99, sheet["B1"]&.value)
      assert_equal(1_000_000, sheet["C1"]&.value)
      assert_in_delta(1.5, sheet["D1"]&.value)
    ensure
      tmp.close!
    end
  end

  test "round-trip preserves multiple sheets" do
    tmp = Tempfile.new(["multi_sheet", ".xlsx"])
    begin
      ws1 = Xlsxrb::Elements::Worksheet.new(
        name: "First",
        rows: [Xlsxrb::Elements::Row.new(
          index: 0,
          cells: [Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "sheet1")]
        )]
      )
      ws2 = Xlsxrb::Elements::Worksheet.new(
        name: "Second",
        rows: [Xlsxrb::Elements::Row.new(
          index: 0,
          cells: [Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "sheet2")]
        )]
      )
      wb = Xlsxrb::Elements::Workbook.new(sheets: [ws1, ws2])

      Xlsxrb.write(tmp.path, wb)
      result = Xlsxrb.load(tmp.path)

      assert_equal(2, result.sheets.size)
      assert_equal(%w[First Second], result.sheet_names)
      assert_equal("sheet1", result.sheet("First")["A1"]&.value)
      assert_equal("sheet2", result.sheet("Second")["A1"]&.value)
    ensure
      tmp.close!
    end
  end

  test "round-trip preserves boolean values" do
    tmp = Tempfile.new(["bool_rt", ".xlsx"])
    begin
      cells = [
        Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: true),
        Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 1, value: false)
      ]
      row = Xlsxrb::Elements::Row.new(index: 0, cells: cells)
      ws = Xlsxrb::Elements::Worksheet.new(name: "Bool", rows: [row])
      wb = Xlsxrb::Elements::Workbook.new(sheets: [ws])

      Xlsxrb.write(tmp.path, wb)
      result = Xlsxrb.load(tmp.path)

      assert_equal(true, result.sheet(0)["A1"]&.value)
      assert_equal(false, result.sheet(0)["B1"]&.value)
    ensure
      tmp.close!
    end
  end

  test "round-trip preserves empty string" do
    tmp = Tempfile.new(["empty_str", ".xlsx"])
    begin
      cells = [Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "")]
      row = Xlsxrb::Elements::Row.new(index: 0, cells: cells)
      ws = Xlsxrb::Elements::Worksheet.new(name: "S", rows: [row])
      wb = Xlsxrb::Elements::Workbook.new(sheets: [ws])

      Xlsxrb.write(tmp.path, wb)
      result = Xlsxrb.load(tmp.path)

      assert_equal("", result.sheet(0)["A1"]&.value)
    ensure
      tmp.close!
    end
  end

  test "round-trip with many rows" do
    tmp = Tempfile.new(["many_rows", ".xlsx"])
    begin
      rows = (0...100).map do |i|
        Xlsxrb::Elements::Row.new(
          index: i,
          cells: [
            Xlsxrb::Elements::Cell.new(row_index: i, column_index: 0, value: i),
            Xlsxrb::Elements::Cell.new(row_index: i, column_index: 1, value: "row#{i}")
          ]
        )
      end
      ws = Xlsxrb::Elements::Worksheet.new(name: "Bulk", rows: rows)
      wb = Xlsxrb::Elements::Workbook.new(sheets: [ws])

      Xlsxrb.write(tmp.path, wb)
      result = Xlsxrb.load(tmp.path)

      assert_equal(100, result.sheet(0).rows.size)
      assert_equal(0, result.sheet(0)["A1"]&.value)
      assert_equal(99, result.sheet(0)["A100"]&.value)
      assert_equal("row50", result.sheet(0)["B51"]&.value)
    ensure
      tmp.close!
    end
  end

  test "Xlsxrb.write serializes first-class state, defined_names, comments, and hyperlinks from DOM models" do
    ws = Xlsxrb::Elements::Worksheet.new(
      name: "HiddenSheet",
      state: :hidden,
      comments: [{ ref: "A1", text: "Note on A1", author: "Tester" }],
      hyperlinks: { "B1" => { ref: "B1", url: "https://example.com" } }
    )
    wb = Xlsxrb::Elements::Workbook.new(
      sheets: [ws],
      defined_names: [{ name: "MyConstant", value: "'HiddenSheet'!A1", local_sheet_name: "HiddenSheet" }]
    )
    binary = Xlsxrb.write(wb)
    read_wb = Xlsxrb.load(binary)
    sheet = read_wb.sheets.first

    assert_equal(:hidden, sheet.state)
    assert(sheet.hidden?)
    assert_equal(1, read_wb.defined_names.size)
    assert_equal("MyConstant", read_wb.defined_names.first.name)
    assert_equal("'HiddenSheet'!A1", read_wb.defined_names.first.value)
    assert_not_nil(read_wb.defined_name("MyConstant", sheet: "HiddenSheet"))
    assert_equal("Note on A1", sheet.comment("A1").text)
    assert_equal("https://example.com", sheet.hyperlink("B1").url)
  end

  # --- Streaming read via Xlsxrb.read(&block) ---

  test "Xlsxrb.read with block yields rows one at a time" do
    tmp = Tempfile.new(["read_stream_test", ".xlsx"])
    begin
      rows = (0...5).map do |i|
        Xlsxrb::Elements::Row.new(
          index: i, cells: [Xlsxrb::Elements::Cell.new(row_index: i, column_index: 0, value: i * 10)]
        )
      end
      ws = Xlsxrb::Elements::Worksheet.new(name: "Data", rows: rows)
      wb = Xlsxrb::Elements::Workbook.new(sheets: [ws])
      Xlsxrb.write(tmp.path, wb)

      collected = []
      Xlsxrb.read(tmp.path) do |sheet|
        sheet.each do |row|
          assert(row.is_a?(Xlsxrb::StreamRow) || row.is_a?(Xlsxrb::Elements::Row))
          collected << row.cells[0].value
        end
      end

      assert_equal([0, 10, 20, 30, 40], collected)
    ensure
      tmp.close!
    end
  end

  test "StreamRow supports lazy cell streaming via each_cell" do
    wb = Xlsxrb.build do |b|
      b.sheet("Wide") do |s|
        s.row(%w[A B C D E F])
      end
    end
    binary = Xlsxrb.write(wb)

    cells_streamed = []
    Xlsxrb.read(binary) do |sheet|
      sheet.each_row do |row|
        assert_kind_of(Xlsxrb::StreamRow, row)
        row.each_cell do |cell|
          cells_streamed << cell.value
        end
      end
    end

    assert_equal(%w[A B C D E F], cells_streamed)
  end

  test "StreamSheet#each_cell streams all cells across rows continuously" do
    wb = Xlsxrb.build do |b|
      b.sheet("Grid") do |s|
        s.row([1, 2])
        s.row([3, 4])
      end
    end
    binary = Xlsxrb.write(wb)

    cells = []
    Xlsxrb.read(binary) do |sheet|
      sheet.each_cell do |cell|
        cells << cell.value
      end
    end

    assert_equal([1, 2, 3, 4], cells)
  end

  test "Xlsxrb.read with trim_empty_rows: true omits trailing empty rows" do
    wb = Xlsxrb.build do |b|
      b.sheet("Trimming") do |s|
        s.row(["Header", 100])
        s.row([])
        s.row(["Data", 200])
        s.row([])
        s.row([])
      end
    end
    binary = Xlsxrb.write(wb)

    all_rows = []
    Xlsxrb.read(binary) do |sheet|
      sheet.each_row_values { |v| all_rows << v }
    end
    assert_equal([["Header", 100], [], ["Data", 200], [], []], all_rows)

    trimmed_rows = []
    Xlsxrb.read(binary, trim_empty_rows: true) do |sheet|
      assert_true(sheet.trim_empty_rows?)
      sheet.each_row_values { |v| trimmed_rows << v }
    end
    assert_equal([["Header", 100], [], ["Data", 200]], trimmed_rows)
  end

  test "Xlsxrb.read with multiple sheets streaming" do
    tmp = Tempfile.new(["read_multi_sheet", ".xlsx"])
    begin
      ws1 = Xlsxrb::Elements::Worksheet.new(
        name: "First",
        rows: [Xlsxrb::Elements::Row.new(
          index: 0,
          cells: [Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "A")]
        )]
      )
      ws2 = Xlsxrb::Elements::Worksheet.new(
        name: "Second",
        rows: [Xlsxrb::Elements::Row.new(
          index: 0,
          cells: [Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "B")]
        )]
      )
      wb = Xlsxrb::Elements::Workbook.new(sheets: [ws1, ws2])
      Xlsxrb.write(tmp.path, wb)

      sheets_found = []
      Xlsxrb.read(tmp.path) do |sheet|
        sheet.each_row { |row| sheets_found << [sheet.name, row.cells[0].value] }
      end
      assert_equal([%w[First A], %w[Second B]], sheets_found)
    ensure
      tmp.close!
    end
  end

  test "Xlsxrb.read without block returns lazy StreamSheets requiring load for random access" do
    wb_data = Xlsxrb.build do |b|
      b.sheet("LazySheet") do |s|
        s.row(%w[Header1 Header2])
        s.row([100, 200])
      end
    end
    binary = Xlsxrb.write(wb_data)

    sheets = Xlsxrb.read(binary).to_a
    sheet = sheets.first

    # 1. Sheet is a StreamSheet (O(1) memory)
    assert_kind_of(Xlsxrb::StreamSheet, sheet)

    # 2. Streaming iteration works on StreamSheet without parsing all rows up-front
    streamed_values = []
    sheet.each_row do |row|
      assert_kind_of(Xlsxrb::StreamRow, row)
      streamed_values << row.values
    end
    assert_equal([%w[Header1 Header2], [100, 200]], streamed_values)

    # 3. StreamSheet does not have coordinate random access ([])
    assert_not_respond_to(sheet, :[])

    # 4. Explicit loading into Elements::Worksheet enables random access
    doc_sheet = sheet.load
    assert_kind_of(Xlsxrb::Elements::Worksheet, doc_sheet)
    assert_equal("Header1", doc_sheet["A1"].value)
    assert_equal(200, doc_sheet["B2"].value)
    assert_equal(2, doc_sheet.rows.size)

    # 5. Xlsxrb.load loads all sheets into memory
    loaded_wb = Xlsxrb.load(binary)
    assert_kind_of(Xlsxrb::Elements::Worksheet, loaded_wb.sheet(0))
    assert_equal("Header1", loaded_wb["LazySheet"]["A1"].value)
  end

  # --- generate ---

  test "Xlsxrb.write creates a valid XLSX" do
    tmp = Tempfile.new(["generate_test", ".xlsx"])
    begin
      Xlsxrb.write(tmp.path) do |w|
        w.sheet("Output")
        w.row(%w[Name Score])
        w.row(["Alice", 95])
        w.row(["Bob", 87])
      end

      result = Xlsxrb.load(tmp.path)
      assert_equal(1, result.sheets.size)
      assert_equal("Output", result.sheet(0).name)
      assert_equal(3, result.sheet(0).rows.size)
      assert_equal("Name", result.sheet(0)["A1"]&.value)
      assert_equal(95, result.sheet(0)["B2"]&.value)
    ensure
      tmp.close!
    end
  end

  test "Xlsxrb.write with multiple sheets" do
    tmp = Tempfile.new(["gen_multi", ".xlsx"])
    begin
      Xlsxrb.write(tmp.path) do |w|
        w.sheet("Sheet1")
        w.row([1, 2, 3])
        w.sheet("Sheet2")
        w.row([4, 5, 6])
      end

      result = Xlsxrb.load(tmp.path)
      assert_equal(2, result.sheets.size)
      assert_equal(1, result.sheet("Sheet1")["A1"]&.value)
      assert_equal(4, result.sheet("Sheet2")["A1"]&.value)
    ensure
      tmp.close!
    end
  end

  test "Xlsxrb.write without explicit sheet" do
    tmp = Tempfile.new(["gen_implicit", ".xlsx"])
    begin
      Xlsxrb.write(tmp.path) do |w|
        w.row(["auto"])
      end

      result = Xlsxrb.load(tmp.path)
      assert_equal("Sheet1", result.sheet(0).name)
      assert_equal("auto", result.sheet(0)["A1"]&.value)
    ensure
      tmp.close!
    end
  end

  test "Xlsxrb.write raises on nil target" do
    assert_raise(Xlsxrb::Error) { Xlsxrb.write(nil) { |_w| } } # rubocop:disable Lint/EmptyBlock
  end

  test "Xlsxrb.write raises without block" do
    assert_raise(Xlsxrb::Error) { Xlsxrb.write("/tmp/test.xlsx") }
  end

  test "Xlsxrb.write with booleans and nil" do
    tmp = Tempfile.new(["gen_types", ".xlsx"])
    begin
      Xlsxrb.write(tmp.path) do |w|
        w.row([true, false, nil, "text", 42])
      end

      result = Xlsxrb.load(tmp.path)
      row = result.sheet(0).rows[0]
      assert_equal(true, row.cell_at(0).value)
      assert_equal(false, row.cell_at(1).value)
      assert_nil(row.cell_at(2)&.value)
      assert_equal("text", row.cell_at(3).value)
      assert_equal(42, row.cell_at(4).value)
    ensure
      tmp.close!
    end
  end

  # --- Streaming benchmarks ---

  test "read with block processes large file without excessive memory" do
    tmp = Tempfile.new(["large_foreach", ".xlsx"])
    begin
      Xlsxrb.write(tmp.path) do |w|
        w.sheet("Big")
        10_000.times do |i|
          w.row([i, "row#{i}", i * 0.5])
        end
      end

      count = 0
      sum = 0
      Xlsxrb.read(tmp.path) do |sheet|
        sheet.each do |row|
          count += 1
          sum += row.cells[0].value.to_i
        end
      end

      assert_equal(10_000, count)
      assert_equal((0...10_000).sum, sum)
    ensure
      tmp.close!
    end
  end

  test "generate can write many rows" do
    tmp = Tempfile.new(["large_gen", ".xlsx"])
    begin
      Xlsxrb.write(tmp.path) do |w|
        w.sheet("Large")
        10_000.times do |i|
          w.row([i, "data#{i}"])
        end
      end

      result = Xlsxrb.load(tmp.path)
      assert_equal(10_000, result.sheet(0).rows.size)
      assert_equal(0, result.sheet(0)["A1"]&.value)
      assert_equal(9999, result.sheet(0)["A10000"]&.value)
    ensure
      tmp.close!
    end
  end

  # --- Round-trip ---

  test "write then read then write again produces consistent result" do
    tmp1 = Tempfile.new(["rt1", ".xlsx"])
    tmp2 = Tempfile.new(["rt2", ".xlsx"])
    begin
      ws = Xlsxrb::Elements::Worksheet.new(
        name: "RT",
        rows: [
          Xlsxrb::Elements::Row.new(
            index: 0,
            cells: [
              Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "A"),
              Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 1, value: 1)
            ]
          ),
          Xlsxrb::Elements::Row.new(
            index: 1,
            cells: [
              Xlsxrb::Elements::Cell.new(row_index: 1, column_index: 0, value: "B"),
              Xlsxrb::Elements::Cell.new(row_index: 1, column_index: 1, value: 2)
            ]
          )
        ]
      )
      wb = Xlsxrb::Elements::Workbook.new(sheets: [ws])
      Xlsxrb.write(tmp1.path, wb)

      wb2 = Xlsxrb.load(tmp1.path)
      Xlsxrb.write(tmp2.path, wb2)

      wb3 = Xlsxrb.load(tmp2.path)

      assert_equal("A", wb3.sheet(0)["A1"]&.value)
      assert_equal(1, wb3.sheet(0)["B1"]&.value)
      assert_equal("B", wb3.sheet(0)["A2"]&.value)
      assert_equal(2, wb3.sheet(0)["B2"]&.value)
    ensure
      tmp1.close!
      tmp2.close!
    end
  end

  test "generate then foreach round-trip" do
    tmp = Tempfile.new(["gen_foreach_rt", ".xlsx"])
    begin
      Xlsxrb.write(tmp.path) do |w|
        w.row(["x", 1])
        w.row(["y", 2])
        w.row(["z", 3])
      end

      values = []
      Xlsxrb.read(tmp.path) do |sheet|
        sheet.each do |row|
          values << row.values
        end
      end

      assert_equal([["x", 1], ["y", 2], ["z", 3]], values)
    ensure
      tmp.close!
    end
  end

  test "chart works in Streaming generate API" do
    tmp = Tempfile.new(["facade_chart_stream", ".xlsx"])
    Xlsxrb.write(tmp.path) do |w|
      w.sheet("Sales") do |s|
        s.row(%w[Month Value])
        s.row(["Jan", 100])
        s.row(["Feb", 200])
        w.chart(type: :bar, title: "Sales Data", series: [{ cat_ref: "Sales!$A$2:$A$3", val_ref: "Sales!$B$2:$B$3" }])
      end
    end

    # Use reader to verify the chart was generated
    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    charts = reader.charts
    assert_equal(1, charts.size)
    assert_equal("barChart", charts[0][:chart_type])
    assert_equal("Sales Data", charts[0][:title])
  ensure
    tmp&.close
    tmp&.unlink
  end

  test "chart works in In-Memory build API" do
    workbook = Xlsxrb.build do |w|
      w.sheet("Sales") do |s|
        s.row(%w[Month Value])
        s.row(["Jan", 100])
        s.row(["Feb", 200])
        s.chart(type: :pie, title: "Sales Pie", series: [{ cat_ref: "Sales!$A$2:$A$3", val_ref: "Sales!$B$2:$B$3" }])
      end
    end

    tmp = Tempfile.new(["facade_chart_mem", ".xlsx"])
    Xlsxrb.write(tmp.path, workbook)

    # Use reader to verify the chart was generated
    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    charts = reader.charts
    assert_equal(1, charts.size)
    assert_equal("pieChart", charts[0][:chart_type])
    assert_equal("Sales Pie", charts[0][:title])
  ensure
    tmp&.close
    tmp&.unlink
  end

  test "style supports options form in build API" do
    workbook = Xlsxrb.build do |w|
      w.sheet("Styled") do |s|
        s.style("header", bold: true, size: 12, font_color: "FF0000FF")
        s.row(%w[Name Score], cell_styles: %w[header header])
      end
    end

    tmp = Tempfile.new(["facade_style_opts", ".xlsx"])
    Xlsxrb.write(tmp.path, workbook)

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    styles = reader.cell_styles
    assert(styles.key?("A1"))
    assert_equal(true, styles["A1"].dig(:font, :bold))
  ensure
    tmp&.close
    tmp&.unlink
  end

  test "chart supports block form in generate API" do
    tmp = Tempfile.new(["facade_chart_block_stream", ".xlsx"])
    Xlsxrb.write(tmp.path) do |w|
      w.sheet("Sales") do
        w.row(%w[Month Value])
        w.row(["Jan", 100])
        w.row(["Feb", 200])
        w.chart do |c|
          c.type = :bar
          c.title = "Sales Data"
          c.series(cat_ref: "Sales!$A$2:$A$3", val_ref: "Sales!$B$2:$B$3")
        end
      end
    end

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    charts = reader.charts
    assert_equal(1, charts.size)
    assert_equal("barChart", charts[0][:chart_type])
    assert_equal("Sales Data", charts[0][:title])
  ensure
    tmp&.close
    tmp&.unlink
  end

  # 3. Workbook#[]
  test "Workbook#[] fetches sheet by index or name" do
    wb = Xlsxrb::Elements::Workbook.new(
      sheets: [
        Xlsxrb::Elements::Worksheet.new(name: "First"),
        Xlsxrb::Elements::Worksheet.new(name: "Second")
      ]
    )

    assert_equal "First", wb[0].name
    assert_equal "Second", wb["Second"].name
    assert_nil wb["Nonexistent"]
  end

  # 4. Hash and Range styling in sheet.row and sheet.column
  test "Hash and Range styling in sheet.row and sheet.column" do
    temp_file = Tempfile.new(["test_styles", ".xlsx"])
    temp_file.close

    Xlsxrb.write(temp_file.path) do |wb|
      wb.style("bold", &:bold)
      wb.style("italic", &:italic)
      wb.style("red") { |s| s.font_color(:red) }

      wb.sheet("Test")
      wb.row(
        { "A" => 1, "B" => 2, "C" => 3, "D" => 4 },
        cell_styles: { "A" => "bold", "B".."C" => "italic", "D" => "red" }
      )

      wb.column("A".."B", width: 20)
      wb.column(%w[C D], width: 10)
    end

    parsed = Xlsxrb.load(temp_file.path)
    sheet = parsed.sheet("Test")

    assert_equal 1, sheet["A1"]&.value
    assert_equal 2, sheet["B1"]&.value
    assert_equal 3, sheet["C1"]&.value
    assert_equal 4, sheet["D1"]&.value
  ensure
    temp_file&.unlink
  end

  # 5. StyleBuilder properties
  test "StyleBuilder apply_options! configures correctly" do
    sb = Xlsxrb::StyleBuilder.new("test")
    sb.apply_options!(
      font: { bold: true, color: :red, name: "Arial", size: 14 },
      fill: { color: :blue },
      border: { all: { style: "thick", color: :black } },
      alignment: { horizontal: "center", vertical: "top", wrap_text: true },
      number_format: "0.00"
    )

    assert_equal true, sb.font_props[:bold]
    assert_equal "FFFF0000", sb.font_props[:color]
    assert_equal "Arial", sb.font_props[:name]
    assert_equal 14, sb.font_props[:size]

    assert_equal "solid", sb.fill_props[:pattern]
    assert_equal "FF0000FF", sb.fill_props[:fg_color]

    assert_equal "thick", sb.border_props[:top][:style]

    assert_equal "center", sb.alignment[:horizontal]
    assert_equal "0.00", sb.number_format
  end

  test "StyleBuilder DSL configures correctly" do
    sb = Xlsxrb::StyleBuilder.new("test")
    sb.font(bold: true, color: :red)
    sb.fill_color(:blue)
    sb.border_all(style: "thick")
    sb.align_horizontal("center")
    sb.number_format("0.00")

    assert_equal true, sb.font_props[:bold]
    assert_equal "FFFF0000", sb.font_props[:color]
    assert_equal "FF0000FF", sb.fill_props[:fg_color]
  end

  test "Xlsxrb.formula returns a formula element" do
    f = Xlsxrb.formula("SUM(A1:A10)")
    assert_instance_of(Xlsxrb::Elements::Formula, f)
    assert_equal("SUM(A1:A10)", f.expression)
    assert_nil(f.cached_value)
    assert_equal(true, f.calculate_always)
  end

  test "Xlsxrb.array_formula returns an array formula element and serializes in XML" do
    f = Xlsxrb.array_formula("SUM(B1:B10*C1:C10)", ref: "A1:C1", cached_value: 0)
    assert_instance_of(Xlsxrb::Elements::Formula, f)
    assert_equal("SUM(B1:B10*C1:C10)", f.expression)
    assert_equal(:array, f.type)
    assert_equal("A1:C1", f.ref)
    assert_equal(0, f.cached_value)

    # Test serialization in streaming mode
    data = Xlsxrb.write do |w|
      w.sheet("Sheet1") do |s|
        s.row([f])
      end
    end
    xml_content = Xlsxrb::Ooxml::ZipReader.open(StringIO.new(data)) { |zr| zr.read_entry("xl/worksheets/sheet1.xml") }
    assert_includes(xml_content, '<f t="array" ref="A1:C1">SUM(B1:B10*C1:C10)</f><v>0</v>')

    # Test serialization in DOM mode
    wb = Xlsxrb.build do |b|
      b.sheet("Sheet1") do |s|
        s.row([f])
      end
    end
    dom_data = Xlsxrb.write(wb)
    dom_xml = Xlsxrb::Ooxml::ZipReader.open(StringIO.new(dom_data)) { |zr| zr.read_entry("xl/worksheets/sheet1.xml") }
    assert_includes(dom_xml, '<f t="array" ref="A1:C1">SUM(B1:B10*C1:C10)</f><v>0</v>')
  end

  test "Xlsxrb.rich_text returns a rich text element" do
    rt = Xlsxrb.rich_text(text: "Hello", bold: true)
    assert_instance_of(Xlsxrb::Elements::RichText, rt)
    assert_equal(1, rt.runs.size)
    assert_equal("Hello", rt.runs[0].text)
    assert_equal({ bold: true }, rt.runs[0].font)
  end

  test "Xlsxrb.rich_text supports multiple runs" do
    rt = Xlsxrb.rich_text({ text: "A", font: { bold: true } }, { text: "B", font: { italic: true } })
    assert_instance_of(Xlsxrb::Elements::RichText, rt)
    assert_equal(2, rt.runs.size)
    assert_equal("A", rt.runs[0].text)
    assert_equal({ bold: true }, rt.runs[0].font)
    assert_equal("B", rt.runs[1].text)
    assert_equal({ italic: true }, rt.runs[1].font)
  end

  test "Xlsxrb.rich_text supports text and font pair arrays" do
    rt = Xlsxrb.rich_text("Normal, ", ["bold text, ", { bold: true }], ["red text", { color: "FF0000" }])
    assert_equal("Normal, bold text, red text", rt.to_s)
    assert_equal(3, rt.runs.size)
    assert_equal("Normal, ", rt.runs[0].text)
    assert_nil(rt.runs[0].font)
    assert_equal("bold text, ", rt.runs[1].text)
    assert_equal({ bold: true }, rt.runs[1].font)
    assert_equal("red text", rt.runs[2].text)
    assert_equal({ color: "FF0000" }, rt.runs[2].font)
  end

  test "Xlsxrb.rich_text supports alternating format hashes and format objects with to_font_hash" do
    custom_format = Object.new
    def custom_format.to_font_hash
      { italic: true, color: "0000FF" }
    end

    rt = Xlsxrb.rich_text("Normal, ", { bold: true }, "bold text, ", custom_format, "italic blue", ", normal end")
    assert_equal("Normal, bold text, italic blue, normal end", rt.to_s)
    assert_equal(4, rt.runs.size)
    assert_nil(rt.runs[0].font)
    assert_equal({ bold: true }, rt.runs[1].font)
    assert_equal({ italic: true, color: "0000FF" }, rt.runs[2].font)
    assert_nil(rt.runs[3].font)
  end

  test "WorkbookBuilder#workbook_property configures correctly" do
    wb = Xlsxrb.build do |w|
      w.workbook_property(:test_prop, "test_val")
      w.sheet("S")
    end
    assert_equal("test_val", wb.unmapped_data.dig(:facade, :workbook_properties, :test_prop))
  end

  test "WorksheetBuilder#auto_filter configures correctly" do
    wb = Xlsxrb.build do |w|
      w.sheet("S") do |s|
        s.auto_filter("A1:B10")
      end
    end
    assert_equal("A1:B10", wb.sheet(0).unmapped_data.dig(:facade, :auto_filter))
  end

  test "WorksheetBuilder#freeze_pane configures correctly" do
    wb = Xlsxrb.build do |w|
      w.sheet("S") do |s|
        s.freeze_pane(row: 1, column: "B")
      end
    end
    assert_equal({ row: 1, column: 1, col: 1 }, wb.sheet(0).unmapped_data.dig(:facade, :freeze_pane))
  end

  test "WorksheetBuilder#header_footer configures correctly" do
    wb = Xlsxrb.build do |w|
      w.sheet("S") do |s|
        s.header_footer(odd_header: "test")
      end
    end
    assert_equal({ odd_header: "test" }, wb.sheet(0).unmapped_data.dig(:facade, :header_footer))
  end

  test "WorkbookBuilder#core_property configures correctly" do
    wb = Xlsxrb.build do |w|
      w.core_property(:creator, "test")
      w.sheet("S")
    end
    assert_equal({ creator: "test" }, wb.unmapped_data.dig(:facade, :core_properties))
  end

  test "WorkbookBuilder#app_property configures correctly" do
    wb = Xlsxrb.build do |w|
      w.app_property(:company, "test")
      w.sheet("S")
    end
    assert_equal({ company: "test" }, wb.unmapped_data.dig(:facade, :app_properties))
  end

  test "WorkbookBuilder#custom_property configures correctly" do
    wb = Xlsxrb.build do |w|
      w.custom_property("test", "test")
      w.sheet("S")
    end
    assert_equal([{ name: "test", value: "test", type: :string }], wb.unmapped_data.dig(:facade, :custom_properties))
  end

  test "WorkbookBuilder#workbook_protection configures correctly" do
    wb = Xlsxrb.build do |w|
      w.workbook_protection(password: "123")
      w.sheet("S")
    end
    assert_equal({ password: "123" }, wb.unmapped_data.dig(:facade, :workbook_protection))
  end

  test "Elements::Workbook provides document properties and protection accessors" do
    wb = Xlsxrb.build do |w|
      w.core_property(:title, "Quarterly Report")
      w.core_property(:creator, "Finance Team")
      w.app_property(:company, "Acme Corp")
      w.custom_property("Confidential", true, type: :boolean)
      w.workbook_protection(lock_structure: true)
      w.sheet("S") do |s|
        s.row(["data"])
      end
    end

    assert_equal("Quarterly Report", wb.core_properties.title)
    assert_equal("Finance Team", wb.core_properties.creator)
    assert_equal("Acme Corp", wb.app_properties.company)
    assert_equal(true, wb.custom_properties["Confidential"].value)
    assert_equal(true, wb.workbook_protection.lock_structure)

    # Test round-trip via write and read
    binary = Xlsxrb.write(wb)
    read_wb = Xlsxrb.load(binary)

    assert_equal("Quarterly Report", read_wb.core_properties.title)
    assert_equal("Finance Team", read_wb.core_properties.creator)
    assert_equal("Acme Corp", read_wb.app_properties.company)
    assert_equal(1, read_wb.custom_properties.size)
    assert_equal("Confidential", read_wb.custom_properties["Confidential"].name)
    assert_equal(true, read_wb.custom_properties["Confidential"].value)
    assert_equal(true, read_wb.workbook_protection.lock_structure)
  end

  test "WorksheetBuilder#split_pane configures correctly" do
    wb = Xlsxrb.build do |w|
      w.sheet("S") do |s|
        s.split_pane(x_split: 1, y_split: 1)
      end
    end
    val = wb.sheet(0).unmapped_data.dig(:facade, :split_pane)
    assert_not_nil(val)
    assert_not_equal([], val)
    assert_not_equal({}, val)
  end

  test "WorksheetBuilder#page_margins configures correctly" do
    wb = Xlsxrb.build do |w|
      w.sheet("S") do |s|
        s.page_margins(left: 0.5)
      end
    end
    val = wb.sheet(0).unmapped_data.dig(:facade, :page_margins)
    assert_not_nil(val)
    assert_not_equal([], val)
    assert_not_equal({}, val)
  end

  test "WorksheetBuilder#page_setup configures correctly" do
    wb = Xlsxrb.build do |w|
      w.sheet("S") do |s|
        s.page_setup(orientation: "landscape")
      end
    end
    val = wb.sheet(0).unmapped_data.dig(:facade, :page_setup)
    assert_not_nil(val)
    assert_not_equal([], val)
    assert_not_equal({}, val)
  end

  test "WorksheetBuilder#sheet_protection configures correctly" do
    wb = Xlsxrb.build do |w|
      w.sheet("S") do |s|
        s.sheet_protection(password: "123")
      end
    end
    val = wb.sheet(0).unmapped_data.dig(:facade, :sheet_protection)
    assert_not_nil(val)
    assert_not_equal([], val)
    assert_not_equal({}, val)
  end

  test "WorksheetBuilder#filter_column configures correctly" do
    wb = Xlsxrb.build do |w|
      w.sheet("S") do |s|
        s.filter_column(0, { type: :value })
      end
    end
    val = wb.sheet(0).unmapped_data.dig(:facade, :filter_columns)
    assert_not_nil(val)
    assert_not_equal([], val)
    assert_not_equal({}, val)
  end

  test "WorksheetBuilder#sort_state configures correctly" do
    wb = Xlsxrb.build do |w|
      w.sheet("S") do |s|
        s.sort_state("A1:A10", [])
      end
    end
    val = wb.sheet(0).unmapped_data.dig(:facade, :sort_state)
    assert_not_nil(val)
    assert_not_equal([], val)
    assert_not_equal({}, val)
  end

  test "WorksheetBuilder#data_validation configures correctly" do
    wb = Xlsxrb.build do |w|
      w.sheet("S") do |s|
        s.data_validation("A1", type: :whole)
      end
    end
    val = wb.sheet(0).unmapped_data.dig(:facade, :data_validations)
    assert_not_nil(val)
    assert_not_equal([], val)
    assert_not_equal({}, val)
  end

  test "WorksheetBuilder#conditional_format configures correctly" do
    wb = Xlsxrb.build do |w|
      w.sheet("S") do |s|
        s.conditional_format("A1", type: :cellIs)
      end
    end
    val = wb.sheet(0).unmapped_data.dig(:facade, :conditional_formats)
    assert_not_nil(val)
    assert_not_equal([], val)
    assert_not_equal({}, val)
  end

  test "WorksheetBuilder#table configures correctly" do
    wb = Xlsxrb.build do |w|
      w.sheet("S") do |s|
        s.table("A1:B2", columns: [{ name: "A" }])
      end
    end
    val = wb.sheet(0).unmapped_data.dig(:facade, :tables)
    assert_not_nil(val)
    assert_not_equal([], val)
    assert_not_equal({}, val)
  end

  test "WorksheetBuilder#comment configures correctly" do
    wb = Xlsxrb.build do |w|
      w.sheet("S") do |s|
        s.comment("A1", "test")
      end
    end
    val = wb.sheet(0).unmapped_data.dig(:facade, :comments)
    assert_not_nil(val)
    assert_not_equal([], val)
    assert_not_equal({}, val)
  end

  test "WorksheetBuilder#image configures correctly" do
    wb = Xlsxrb.build do |w|
      w.sheet("S") do |s|
        s.image("test.png", at: "A1")
      end
    end
    val = wb.sheet(0).unmapped_data.dig(:facade, :images)
    assert_not_nil(val)
    assert_not_equal([], val)
    assert_not_equal({}, val)
  end

  test "WorksheetBuilder#shape configures correctly" do
    wb = Xlsxrb.build do |w|
      w.sheet("S") do |s|
        s.shape(type: :rect, at: "A1")
      end
    end
    val = wb.sheet(0).unmapped_data.dig(:facade, :shapes)
    assert_not_nil(val)
    assert_not_equal([], val)
    assert_not_equal({}, val)
  end

  test "WorkbookBuilder#defined_name configures correctly" do
    wb = Xlsxrb.build do |w|
      w.defined_name("A", "B")
      w.sheet("S")
    end
    val = wb.unmapped_data.dig(:facade, :defined_names)
    assert_not_nil(val)
    assert_not_equal([], val)
    assert_not_equal({}, val)
  end

  test "WorkbookBuilder#print_area configures correctly" do
    wb = Xlsxrb.build do |w|
      w.print_area("A1:A2")
      w.sheet("S")
    end
    val = wb.unmapped_data.dig(:facade, :defined_names)
    assert_not_nil(val)
    assert_not_equal([], val)
    assert_not_equal({}, val)
  end

  test "WorkbookBuilder#print_titles configures correctly" do
    wb = Xlsxrb.build do |w|
      w.print_titles(rows: "1:2")
      w.sheet("S")
    end
    val = wb.unmapped_data.dig(:facade, :defined_names)
    assert_not_nil(val)
    assert_not_equal([], val)
    assert_not_equal({}, val)
  end

  test "WorksheetBuilder#print_options configures correctly" do
    wb = Xlsxrb.build do |w|
      w.sheet("S") do |s|
        s.print_options(:headings, true)
      end
    end
    val = wb.sheet(0).unmapped_data.dig(:facade, :print_options)
    assert_not_nil(val)
    assert_not_equal([], val)
    assert_not_equal({}, val)
  end

  test "WorksheetBuilder#sheet_properties configures correctly" do
    wb = Xlsxrb.build do |w|
      w.sheet("S") do |s|
        s.sheet_properties(:tab_color, "FF0000")
      end
    end
    val = wb.sheet(0).unmapped_data.dig(:facade, :sheet_properties)
    assert_not_nil(val)
    assert_not_equal([], val)
    assert_not_equal({}, val)
  end

  test "WorksheetBuilder#sheet_view configures correctly" do
    wb = Xlsxrb.build do |w|
      w.sheet("S") do |s|
        s.sheet_view(:zoom_scale, 150)
      end
    end
    val = wb.sheet(0).unmapped_data.dig(:facade, :sheet_view)
    assert_not_nil(val)
    assert_not_equal([], val)
    assert_not_equal({}, val)
  end

  test "WorkbookBuilder#properties configures correctly" do
    wb = Xlsxrb.build do |w|
      w.properties(core: { creator: "test" })
      w.sheet("S")
    end
    val = wb.unmapped_data.dig(:facade, :core_properties)
    assert_not_nil(val)
    assert_not_equal([], val)
    assert_not_equal({}, val)
  end

  test "WorksheetBuilder#hyperlink and StreamWriter#hyperlink accept Elements::Hyperlink" do
    hl = Xlsxrb.hyperlink("https://example.com", display: "Example Site", tooltip: "Go to Example")

    # StreamWriter
    stream_data = Xlsxrb.write do |w|
      w.sheet("Links") do |s|
        s.row(["Test"])
        s.hyperlink("A1", hl)
      end
    end
    xml = Xlsxrb::Ooxml::ZipReader.open(StringIO.new(stream_data)) { |zr| zr.read_entry("xl/worksheets/sheet1.xml") }
    assert_includes(xml, 'tooltip="Go to Example"')
    assert_includes(xml, 'display="Example Site"')

    # WorksheetBuilder
    wb = Xlsxrb.build do |w|
      w.sheet("Links") do |s|
        s.row(["Test"])
        s.hyperlink("A1", hl)
      end
    end
    dom_data = Xlsxrb.write(wb)
    dom_xml = Xlsxrb::Ooxml::ZipReader.open(StringIO.new(dom_data)) { |zr| zr.read_entry("xl/worksheets/sheet1.xml") }
    assert_includes(dom_xml, 'tooltip="Go to Example"')
    assert_includes(dom_xml, 'display="Example Site"')
  end

  test "WorksheetBuilder#sparkline_group configures correctly" do
    wb = Xlsxrb.build do |w|
      w.sheet("S") do |s|
        s.sparkline_group(sparklines: ["A1"])
      end
    end
    val = wb.sheet(0).unmapped_data.dig(:facade, :sparkline_groups)
    assert_not_nil(val)
    assert_not_equal([], val)
    assert_not_equal({}, val)
  end

  # Tests migrated from test/facade_detailed_mutation_test.rb
  test "ChartBuilder#series adds value directly when no block given" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.chart(target: "A1", type: :bar) do |c|
          c.series(val: "Sheet1!$B$1:$B$5")
        end
      end
    end
    c = wb.sheet(0).charts.first
    assert_equal [{ val: "Sheet1!$B$1:$B$5" }], c.series
  end

  test "WorksheetBuilder#hyperlink captures options properly" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.hyperlink("A1", "https://example.com", tooltip: "Example")
      end
    end
    links = wb.sheet(0).unmapped_data.dig(:facade, :hyperlinks)
    assert_equal "Example", links.first[:tooltip]
  end

  test "WorksheetBuilder#pivot_table captures options properly" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.pivot_table("Sheet1!A1:B10", row_fields: ["A"], data_fields: ["B"], name: "Pivot1")
      end
    end
    pivots = wb.sheet(0).unmapped_data.dig(:facade, :pivot_tables)
    assert_equal "Pivot1", pivots.first[:name]
  end

  test "WorksheetBuilder#merge supports different forms" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.merge("A1:B2")
      end
    end
    merges = wb.sheet(0).unmapped_data.dig(:facade, :merge_cells)
    assert_equal "A1:B2", merges.first
  end

  test "WorksheetBuilder#selection captures active cell" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.selection("C3")
      end
    end
    sel = wb.sheet(0).unmapped_data.dig(:facade, :selection)
    assert_equal "C3", sel[:active_cell]
  end

  test "WorksheetBuilder#row_break adds row breaks" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.row_break(10)
        s.row_break(20)
      end
    end
    breaks = wb.sheet(0).unmapped_data.dig(:facade, :row_breaks)
    assert_equal 10, breaks[0][:id]
    assert_equal 20, breaks[1][:id]
  end

  test "WorksheetBuilder#column_break adds col breaks" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.column_break("C")
        s.column_break(5)
      end
    end
    breaks = wb.sheet(0).unmapped_data.dig(:facade, :col_breaks)
    assert_equal 2, breaks[0][:id]
    assert_equal 5, breaks[1][:id]
  end

  test "WorksheetBuilder#auto_filter captures range" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.auto_filter("A1:B10")
      end
    end
    filter = wb.sheet(0).unmapped_data.dig(:facade, :auto_filter)
    assert_equal "A1:B10", filter
  end

  test "WorksheetBuilder#filter_column captures options" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.filter_column(0, { val: "A" })
      end
    end
    cols = wb.sheet(0).unmapped_data.dig(:facade, :filter_columns)
    assert_equal "A", cols[0][:val]
  end

  test "WorksheetBuilder#sort_state captures ref" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.sort_state("A2:A10", [])
      end
    end
    state = wb.sheet(0).unmapped_data.dig(:facade, :sort_state)
    assert_equal "A2:A10", state[:ref]
  end

  test "WorksheetBuilder#data_validation captures validation" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.data_validation("A1", type: :list)
      end
    end
    vals = wb.sheet(0).unmapped_data.dig(:facade, :data_validations)
    assert_equal :list, vals.first[:type]
  end

  test "WorksheetBuilder#comment captures text" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.comment("A1", "Comment text", author: "Author")
      end
    end
    vals = wb.sheet(0).unmapped_data.dig(:facade, :comments)
    assert_equal "Comment text", vals.first[:text]
    assert_equal "Author", vals.first[:author]
  end

  test "WorksheetBuilder#sparkline_group captures options" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.sparkline_group(sparklines: ["A1"])
      end
    end
    vals = wb.sheet(0).unmapped_data.dig(:facade, :sparkline_groups)
    assert_equal ["A1"], vals.first[:sparklines]
  end

  test "WorksheetBuilder#freeze_pane captures cell" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.freeze_pane(row: 1, column: "B")
      end
    end
    pane = wb.sheet(0).unmapped_data.dig(:facade, :freeze_pane)
    assert_equal 1, pane[:row]
    assert_equal 1, pane[:col]
  end

  test "WorksheetBuilder#split_pane captures values" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.split_pane(x_split: 10, y_split: 20)
      end
    end
    pane = wb.sheet(0).unmapped_data.dig(:facade, :split_pane)
    assert_equal 10, pane[:x_split]
    assert_equal 20, pane[:y_split]
  end

  test "WorksheetBuilder#page_margins captures values" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.page_margins(left: 1.0)
      end
    end
    marg = wb.sheet(0).unmapped_data.dig(:facade, :page_margins)
    assert_equal 1.0, marg[:left]
  end

  # Tests migrated from test/facade_detailed_mutation_test_3.rb
  test "WorksheetBuilder#shape captures values" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.shape(preset: "rect", text: "Hi")
      end
    end
    shapes = wb.sheet(0).shapes
    assert_equal "rect", shapes.first.preset
    assert_equal "Hi", shapes.first.text
    assert_equal 1, shapes.size
  end

  test "WorksheetBuilder#table captures values" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.table("A1:B10", columns: [], name: "Table1")
      end
    end
    tables = wb.sheet(0).tables
    assert_equal "Table1", tables.first.name
    assert_equal tables.map(&:to_h), wb.sheet(0).unmapped_data.dig(:facade, :tables)
  end

  test "WorkbookBuilder#custom_property captures values" do
    wb = Xlsxrb.build do |w|
      w.sheet("S")
      w.custom_property("Prop1", "Val1")
    end
    cp = wb.unmapped_data.dig(:facade, :custom_properties)
    assert_equal "Val1", cp.first[:value]
  end

  test "WorkbookBuilder#print_area captures values" do
    wb = Xlsxrb.build do |w|
      w.sheet("Sheet1")
      w.print_area("A1:B10", sheet: "Sheet1")
    end
    dn = wb.unmapped_data.dig(:facade, :defined_names)
    assert(dn.any? { |d| d[:name] == "_xlnm.Print_Area" })
  end

  test "WorkbookBuilder#properties captures block" do
    wb = Xlsxrb.build do |w|
      w.sheet("S")
      w.properties(core: { creator: "Dev" })
    end
    cp = wb.unmapped_data.dig(:facade, :core_properties)
    assert_equal "Dev", cp[:creator]
  end

  # Tests migrated from test/facade_detailed_mutation_test_4.rb
  test "WorksheetBuilder#row captures height" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.row([1, 2], height: 30)
      end
    end
    row = wb.sheet(0).rows.first
    assert_equal 30, row.height
    assert_equal 1, row.cells[0].value
  end

  test "WorksheetBuilder#column captures options" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.column(0, width: 25, hidden: true)
      end
    end
    col = wb.sheet(0).columns.first
    assert_equal 25, col.width
    assert_equal true, col.hidden?
  end

  test "WorkbookBuilder#workbook_property captures values" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1")
      w.workbook_property(:date1904, true)
    end
    props = wb.unmapped_data.dig(:facade, :workbook_properties)
    assert_equal true, props[:date1904]
  end

  test "WorksheetBuilder#page_setup captures options" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.page_setup(paper_size: 9, orientation: "landscape")
      end
    end
    setup = wb.sheet(0).unmapped_data.dig(:facade, :page_setup)
    assert_equal 9, setup[:paper_size]
    assert_equal "landscape", setup[:orientation]
  end

  test "WorksheetBuilder#conditional_format captures values" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.conditional_format("A1:A10", type: :cellIs, operator: :greaterThan, formula1: "10")
      end
    end
    cf = wb.sheet(0).unmapped_data.dig(:facade, :conditional_formats)
    assert_equal "A1:A10", cf.first[:sqref]
    assert_equal :cellIs, cf.first[:type]
    assert_equal :greaterThan, cf.first[:operator]
    assert_equal "10", cf.first[:formula1]
  end

  test "ChartBuilder#title works via method_missing" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.chart(target: "A1", type: :pie) do |c|
          c.title text: "My Chart"
        end
      end
    end
    chart = wb.sheet(0).charts.first
    assert_equal({ text: "My Chart" }, chart.title)
  end

  test "ChartBuilder#legend works via method_missing" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.chart(target: "A1", type: :pie) do |c|
          c.legend position: "b"
        end
      end
    end
    chart = wb.sheet(0).charts.first
    assert_equal({ position: "b" }, chart.legend)
  end

  test "ChartBuilder#category_axis works via method_missing" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.chart(target: "A1", type: :pie) do |c|
          c.category_axis title: "X Axis"
        end
      end
    end
    chart = wb.sheet(0).charts.first
    assert_equal({ title: "X Axis" }, chart.category_axis)
  end

  test "ChartBuilder#value_axis works via method_missing" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.chart(target: "A1", type: :pie) do |c|
          c.value_axis title: "Y Axis"
        end
      end
    end
    chart = wb.sheet(0).charts.first
    assert_equal({ title: "Y Axis" }, chart.value_axis)
  end

  test "WorksheetBuilder#chart without block" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.chart(target: "A1", type: :bar, title: { text: "No Block" })
      end
    end
    chart = wb.sheet(0).charts.first
    assert_equal({ text: "No Block" }, chart.title)
  end

  test "Xlsxrb.write writes string directly" do
    wb = Xlsxrb.build do |w|
      w.sheet("S") do |s|
        s.row [1]
      end
    end
    out = StringIO.new
    Xlsxrb.write(out, wb)
    assert out.string.start_with?("PK")
  end

  test "Xlsxrb.write streaming output yields chunks" do
    out = StringIO.new
    Xlsxrb.write(out) do |w|
      w.sheet("S") do |s|
        s.row [1]
      end
    end
    assert out.string.start_with?("PK")
  end

  test "WorksheetBuilder#row with formula" do
    wb = Xlsxrb.build do |w|
      w.sheet("S") do |s|
        s.row [Xlsxrb.formula("SUM(A1:A2)")]
      end
    end
    cell = wb.sheet(0).rows.first.cells.first
    assert_equal "SUM(A1:A2)", cell.formula.expression
  end

  # Tests migrated from test/facade_detailed_mutation_test_5.rb
  test "Xlsxrb.rich_text creates RichText object" do
    rt = Xlsxrb.rich_text("bold part", bold: true)
    assert_equal "Xlsxrb::Elements::RichText", rt.class.name
    assert_equal 1, rt.runs.size
    assert_equal true, rt.runs[0][:font][:bold] if rt.runs[0].is_a?(Hash)
  end

  test "WorkbookBuilder#sheet captures multiple sheets" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1")
      w.sheet("S2")
      w.sheet("S3")
    end
    assert_equal 3, wb.sheets.size
    assert_equal %w[S1 S2 S3], wb.sheets.map(&:name)
  end

  test "WorksheetBuilder#column custom_width and outline_level" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.column(0..2, width: 25, outline_level: 2, custom_width: true)
      end
    end
    col = wb.sheet(0).columns.first
    assert_equal 25, col.width
    assert_equal 2, col.outline_level
    assert_equal true, col.custom_width?
  end

  test "Xlsxrb.read parses workbook" do
    temp = Tempfile.new(["test_read", ".xlsx"])
    Xlsxrb.write(temp.path) do |w|
      w.sheet("S1") { |s| s.row ["A"] }
    end

    wb = Xlsxrb.load(temp.path)
    assert_equal "S1", wb.sheets.first.name
    assert_equal "A", wb.sheets.first.rows.first.cells.first.value
  ensure
    temp&.close
    temp&.unlink
  end

  test "WorksheetBuilder#merge with options" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.merge("A1:B2")
      end
    end
    merges = wb.sheet(0).unmapped_data.dig(:facade, :merge_cells)
    assert_equal "A1:B2", merges.first
  end

  test "WorksheetBuilder#row_break creates break" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.row_break(5)
      end
    end
    breaks = wb.sheet(0).unmapped_data.dig(:facade, :row_breaks)
    assert breaks
  end

  test "WorksheetBuilder#column_break creates break" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.column_break("C")
      end
    end
    breaks = wb.sheet(0).unmapped_data.dig(:facade, :col_breaks)
    assert breaks
  end

  test "WorksheetBuilder#selection captures active cell coordinate" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.selection("D4")
      end
    end
    selection = wb.sheet(0).unmapped_data.dig(:facade, :selection)
    assert_equal "D4", selection[:active_cell]
  end

  test "WorksheetBuilder#data_validation captures options" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.data_validation("A1:A10", type: :list)
      end
    end
    val = wb.sheet(0).unmapped_data.dig(:facade, :data_validations).first
    assert_equal :list, val[:type]
  end

  test "Elements::Cell#to_i converts string" do
    cell = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "123")
    assert_equal 123, cell.to_i
  end

  test "Elements::Cell#to_f converts string" do
    cell = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "123.45")
    assert_equal 123.45, cell.to_f
  end

  test "Elements::Cell#to_s returns string value" do
    cell = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: 123)
    assert_equal "123", cell.to_s
  end

  test "Elements::Cell#to_date converts datetime" do
    cell = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: 46_204.5)
    date = cell.to_date
    assert_not_nil date
    assert_equal 2026, date.year
  end

  test "Elements::Cell#to_time converts datetime" do
    cell = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: 46_204.5)
    time = cell.to_time
    assert_not_nil time
    assert_equal 2026, time.year
  end

  test "Elements::Row#cell_at works" do
    row = Xlsxrb::Elements::Row.new(index: 0, cells: [
                                      Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "A"),
                                      Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 1, value: "B")
                                    ])
    assert_equal "A", row.cell_at(0).value
    assert_equal "B", row.cell_at(1).value
  end

  test "Elements::Row#each_cell iterates" do
    row = Xlsxrb::Elements::Row.new(index: 0, cells: [
                                      Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "A"),
                                      Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 1, value: "B")
                                    ])
    cells = []
    row.each_cell { |c| cells << c.value }
    assert_equal %w[A B], cells
  end

  test "Elements::Row#values returns array" do
    row = Xlsxrb::Elements::Row.new(index: 0, cells: [
                                      Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "A"),
                                      Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 1, value: "B")
                                    ])
    assert_equal %w[A B], row.values
  end

  test "Elements::Row#to_a returns cells via Enumerable#to_a" do
    c1 = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "A")
    c2 = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 1, value: "B")
    row = Xlsxrb::Elements::Row.new(index: 0, cells: [c1, c2])
    assert_equal [c1, c2], row.to_a
    assert_equal %w[A B], row.values
  end

  test "Elements::Worksheet#cell_value works" do
    ws = Xlsxrb::Elements::Worksheet.new(name: "S1", rows: [
                                           Xlsxrb::Elements::Row.new(index: 0, cells: [
                                                                       Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "A")
                                                                     ])
                                         ])
    assert_equal "A", ws["A1"]&.value
  end

  test "Elements::Worksheet#cells returns flattened array" do
    ws = Xlsxrb::Elements::Worksheet.new(name: "S1", rows: [
                                           Xlsxrb::Elements::Row.new(index: 0, cells: [
                                                                       Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "A")
                                                                     ]),
                                           Xlsxrb::Elements::Row.new(index: 1, cells: [
                                                                       Xlsxrb::Elements::Cell.new(row_index: 1, column_index: 0, value: "B")
                                                                     ])
                                         ])
    assert_equal %w[A B], ws.cells.map(&:value)
  end

  test "Elements::Worksheet#first_row works" do
    ws = Xlsxrb::Elements::Worksheet.new(name: "S1", rows: [
                                           Xlsxrb::Elements::Row.new(index: 0, cells: [
                                                                       Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "A")
                                                                     ]),
                                           Xlsxrb::Elements::Row.new(index: 1, cells: [
                                                                       Xlsxrb::Elements::Cell.new(row_index: 1, column_index: 0, value: "B")
                                                                     ])
                                         ])
    assert_equal "A", ws.first_row.cells[0].value
  end

  test "Elements::Worksheet#last_row works" do
    ws = Xlsxrb::Elements::Worksheet.new(name: "S1", rows: [
                                           Xlsxrb::Elements::Row.new(index: 0, cells: [
                                                                       Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "A")
                                                                     ]),
                                           Xlsxrb::Elements::Row.new(index: 1, cells: [
                                                                       Xlsxrb::Elements::Cell.new(row_index: 1, column_index: 0, value: "B")
                                                                     ])
                                         ])
    assert_equal "B", ws.last_row.cells[0].value
  end

  test "Elements::Worksheet#row_at works" do
    ws = Xlsxrb::Elements::Worksheet.new(name: "S1", rows: [
                                           Xlsxrb::Elements::Row.new(index: 0, cells: [
                                                                       Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "A")
                                                                     ])
                                         ])
    assert_equal "A", ws.row_at(0).cells[0].value
  end

  test "Elements::Workbook#sheet_names works" do
    wb = Xlsxrb::Elements::Workbook.new(sheets: [
                                          Xlsxrb::Elements::Worksheet.new(name: "S1", rows: []),
                                          Xlsxrb::Elements::Worksheet.new(name: "S2", rows: [])
                                        ])
    assert_equal %w[S1 S2], wb.sheet_names
  end

  test "Elements::Cell.validate validates correctly" do
    assert_equal [], Xlsxrb::Elements::Cell.validate(0, 0, "A")
    assert_not_equal [], Xlsxrb::Elements::Cell.validate(-1, 0, "A")
  end

  test "Elements::Column.validate validates correctly" do
    assert_equal [], Xlsxrb::Elements::Column.validate(0)
    assert_not_equal [], Xlsxrb::Elements::Column.validate(-1)
  end

  test "Elements::Row.validate validates correctly" do
    assert_equal [], Xlsxrb::Elements::Row.validate(0, [])
    assert_not_equal [], Xlsxrb::Elements::Row.validate(-1, [])
  end

  test "Elements::Workbook.validate validates correctly" do
    assert_equal [], Xlsxrb::Elements::Workbook.validate([Xlsxrb::Elements::Worksheet.new(name: "S1", rows: [])])
  end

  test "Elements::Worksheet.validate validates correctly" do
    assert_equal [], Xlsxrb::Elements::Worksheet.validate("Sheet1", [])
    assert_not_equal [], Xlsxrb::Elements::Worksheet.validate("", [])
  end

  test "Elements::Workbook#sheet finds sheet" do
    s1 = Xlsxrb::Elements::Worksheet.new(name: "S1", rows: [])
    wb = Xlsxrb::Elements::Workbook.new(sheets: [s1])
    assert_equal s1, wb.sheet("S1")
    assert_equal s1, wb.sheet(0)
    assert_nil wb.sheet("S2")
  end

  test "Elements::Workbook#each yields sheets" do
    s1 = Xlsxrb::Elements::Worksheet.new(name: "S1", rows: [])
    wb = Xlsxrb::Elements::Workbook.new(sheets: [s1])
    yielded = wb.map { |s| s }
    assert_equal [s1], yielded
  end

  test "StyleBuilder#font method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.font(name: "Arial", size: 12)
    assert_equal "Arial", builder.font_props[:name]
    assert_equal 12, builder.font_props[:size]
  end

  test "StyleBuilder#bold method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.bold(true)
    assert_equal true, builder.font_props[:bold]
  end

  test "StyleBuilder#italic method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.italic(true)
    assert_equal true, builder.font_props[:italic]
  end

  test "StyleBuilder#font_size method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.font_size(16)
    assert_equal 16, builder.font_props[:size]
  end

  test "StyleBuilder#font_name method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.font_name("Times")
    assert_equal "Times", builder.font_props[:name]
  end

  test "StyleBuilder#font_color method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.font_color("FF00FF")
    assert_equal "FF00FF", builder.font_props[:color]
  end

  test "StyleBuilder#underline method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.underline("double")
    assert_equal "double", builder.font_props[:underline]
  end

  test "StyleBuilder#strike method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.strike(true)
    assert_equal true, builder.font_props[:strike]
  end

  test "StyleBuilder#vert_align method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.vert_align("superscript")
    assert_equal "superscript", builder.font_props[:vert_align]
  end

  test "StyleBuilder#fill_pattern method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.fill_pattern("solid", fg_color: "112233")
    assert_equal "solid", builder.fill_props[:pattern]
    assert_equal "112233", builder.fill_props[:fg_color]
  end

  test "StyleBuilder#fill_color method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.fill_color("00FF00")
    assert_equal "00FF00", builder.fill_props[:fg_color]
    assert_equal "solid", builder.fill_props[:pattern]
  end

  test "StyleBuilder#fill method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.fill(pattern: "darkGray", bg_color: "FFFFFF")
    assert_equal "darkGray", builder.fill_props[:pattern]
    assert_equal "FFFFFF", builder.fill_props[:bg_color]
  end

  test "StyleBuilder#fill_gradient method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.fill_gradient(type: "linear", degree: 45)
    assert_equal "linear", builder.fill_props[:gradient][:type]
    assert_equal 45, builder.fill_props[:gradient][:degree]
  end

  test "StyleBuilder#border method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.border(left: { style: "thick" })
    assert_equal "thick", builder.border_props[:left][:style]
  end

  test "StyleBuilder#border_all method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.border_all(style: "dashed")
    assert_equal "dashed", builder.border_props[:left][:style]
    assert_equal "dashed", builder.border_props[:right][:style]
    assert_equal "dashed", builder.border_props[:top][:style]
    assert_equal "dashed", builder.border_props[:bottom][:style]
  end

  test "StyleBuilder#border_left method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.border_left(style: "dotted")
    assert_equal "dotted", builder.border_props[:left][:style]
  end

  test "StyleBuilder#border_right method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.border_right(style: "dotted")
    assert_equal "dotted", builder.border_props[:right][:style]
  end

  test "StyleBuilder#border_top method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.border_top(style: "dotted")
    assert_equal "dotted", builder.border_props[:top][:style]
  end

  test "StyleBuilder#border_bottom method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.border_bottom(style: "dotted")
    assert_equal "dotted", builder.border_props[:bottom][:style]
  end

  test "StyleBuilder#border_diagonal method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.border_diagonal(style: "thin", up: true)
    assert_equal "thin", builder.border_props[:diagonal][:style]
    assert_equal true, builder.border_props[:diagonal_up]
  end

  test "StyleBuilder#align_horizontal method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.align_horizontal("center")
    assert_equal "center", builder.alignment[:horizontal]
  end

  test "StyleBuilder#align_vertical method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.align_vertical("center")
    assert_equal "center", builder.alignment[:vertical]
  end

  test "StyleBuilder#wrap_text method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.wrap_text(true)
    assert_equal true, builder.alignment[:wrap_text]
  end

  test "StyleBuilder#shrink_to_fit method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.shrink_to_fit(true)
    assert_equal true, builder.alignment[:shrink_to_fit]
  end

  test "StyleBuilder#text_rotation method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.text_rotation(45)
    assert_equal 45, builder.alignment[:text_rotation]
  end

  test "StyleBuilder#indent method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.indent(2)
    assert_equal 2, builder.alignment[:indent]
  end

  test "StyleBuilder#number_format method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.number_format(14)
    assert_equal 14, builder.number_format
  end
  test "WorksheetBuilder#row handles Hash values" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.row({ "B" => 1, "D" => 2 })
      end
    end
    cells = wb.sheet(0).rows[0].cells
    assert_equal 1, cells[0].value
    assert_equal 2, cells[1].value
  end
  test "WorksheetBuilder#row handles Hash styles with ranges" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.style("bold", font: { bold: true })
        s.style("italic", font: { italic: true })
        s.row([1, 2, 3, 4, 5], cell_styles: { "A" => "bold", "C".."D" => "italic" })
      end
    end
    cells = wb.sheet(0).rows[0].cells
    # Styles might be converted to indices, let's just assert they are non-nil integers
    assert_kind_of Integer, cells[0].style_index
    assert_nil cells[1].style_index
    assert_kind_of Integer, cells[2].style_index
    assert_kind_of Integer, cells[3].style_index
    assert_nil cells[4].style_index
  end
  test "WorksheetBuilder#row handles Time values and auto-formats" do
    t = Time.new(2026, 1, 1, 12, 0, 0)
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.row([t])
      end
    end
    cells = wb.sheet(0).rows[0].cells
    assert_kind_of Integer, cells[0].style_index
  end
  test "WorksheetBuilder#row handles Array of Hashes for RichText" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.row([[{ text: "Hello", bold: true }, { text: " World" }]])
      end
    end
    val = wb.sheet(0).rows[0].cells[0].value
    assert_kind_of Xlsxrb::Elements::RichText, val
    assert_equal "Hello", val.runs[0][:text]
    assert_equal true, val.runs[0][:font][:bold]
    assert_equal " World", val.runs[1][:text]
  end
  test "WorksheetBuilder#row handles inline Hash styles" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.row([1], cell_styles: [{ font: { bold: true } }])
      end
    end
    cells = wb.sheet(0).rows[0].cells
    assert_kind_of Integer, cells[0].style_index
  end
  test "WorksheetBuilder#row handles Hash with formula" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.row([{ formula: "SUM(A1:A2)", value: 3 }])
      end
    end
    cell = wb.sheet(0).rows[0].cells[0]
    assert_equal "SUM(A1:A2)", cell.formula.expression
    assert_equal 3, cell.value
  end

  test "Writer handles all sheet configuration options" do
    io = StringIO.new
    Xlsxrb.write(io) do |w|
      w.workbook_property(:update_links, "always")
      w.style("bold", font: { bold: true })

      w.sheet("Test") do |s|
        assert s.respond_to?(:row)

        s.row([1, 2, 3])

        s.merge("A1:C1")
        s.freeze_pane(row: 1, column: 1)
        s.split_pane(x_split: 1000, y_split: 1000, top_left_cell: "B2")
        s.selection("A1", sqref: "A1:A2", pane: "topRight")

        s.page_margins(left: 0.5, right: 0.5, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2)
        s.print_options(:grid_lines, true)
        s.page_setup(paper_size: 9, orientation: "landscape")
        s.header_footer(odd_header: "&L&T")

        s.sheet_view(:show_grid_lines, false)
        s.sheet_protection(password: "secret", sheet: true)
      end
    end
    assert io.size.positive?
  end

  test "Writer raises error when writing to inactive sheet" do
    io = StringIO.new
    Xlsxrb.write(io) do |w|
      s1 = nil
      w.sheet("S1") { |s| s1 = s }
      w.sheet("S2") { |s| s.row([1]) }
      assert_raise(Xlsxrb::Error) { s1.row([1]) }
    end
  end

  test "WorksheetBuilder#row raises error on invalid value type" do
    Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        assert_raise(ArgumentError) { s.row([Object.new]) }
      end
    end
  end

  test "WorksheetBuilder raises error for column width limit in strict mode" do
    Xlsxrb.build(strict_excel_mode: true) do |w|
      w.sheet("S1") do |s|
        assert_raise(ArgumentError) { s.column(0, width: 300) }
      end
    end
  end

  test "Additional Coverage for WorkbookBuilder and Date auto-format" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.row([Date.today])
        s.merge("A1:C1")
      end
    end
    assert wb
  end
  def test_workbook_builder_public_methods
    wb = Xlsxrb::WorkbookBuilder.new(strict_excel_mode: false)
    wb.workbook_property(:update_links, "always")

    # Check defined_name
    wb.defined_name("MyName", "Sheet1!$A$1", hidden: true)

    # Check print_area
    wb.print_area("A1:D10", sheet: "Sheet1")

    # Check print_titles
    wb.print_titles(rows: "1:2", columns: "A:B", sheet: "Sheet1")

    # Check properties
    wb.workbook_protection(password: "secret", lock_structure: true)
    wb.core_property(:title, "Test Title")
    wb.app_property(:company, "MyCompany")
    wb.properties(core: { creator: "Me" }, app: { manager: "Boss" })
    wb.custom_property("MyCustomProp", "Value", type: :string)

    wb.sheet("Sheet1") do |sheet|
      sheet.row(["Data"])
    end

    workbook = wb.build
    assert_instance_of Xlsxrb::Elements::Workbook, workbook
    assert_equal "always", workbook.unmapped_data[:facade][:workbook_properties][:update_links]
    assert_equal "Test Title", workbook.unmapped_data[:facade][:core_properties][:title]
    assert_equal "MyCompany", workbook.unmapped_data[:facade][:app_properties][:company]
    assert_equal "Me", workbook.unmapped_data[:facade][:core_properties][:creator]
    assert_equal "Boss", workbook.unmapped_data[:facade][:app_properties][:manager]
    assert_equal "MyCustomProp", workbook.unmapped_data[:facade][:custom_properties].first[:name]
    assert_equal "secret", workbook.unmapped_data[:facade][:workbook_protection][:password]
  end

  def test_worksheet_builder_public_methods
    wb = Xlsxrb::WorkbookBuilder.new(strict_excel_mode: false)
    wb.sheet("Sheet1") do |sheet|
      # sheet properties
      sheet.sheet_properties(:tab_color, "FF0000")
      sheet.page_setup(orientation: "landscape")
      sheet.page_margins(left: 0.5)

      # rows
      sheet.row(%w[A1 B1 C1], height: 20)
      sheet.row([]) # empty row
      sheet.row([Xlsxrb.formula("SUM(1,2)"), Xlsxrb.rich_text(text: "Hi", bold: true)])

      # merges
      sheet.merge("A1:B1")

      # freeze panes
      sheet.freeze_pane(row: 1, column: 1)

      # column
      sheet.column(1, width: 15, hidden: false)

      # protection
      sheet.sheet_protection(password: "123")

      # data validation
      sheet.data_validation("A1", type: "list", formula1: '"A,B,C"')

      # chart
      sheet.chart(type: "bar", anchor: "D1", width: 300, height: 200) do |c|
        c.title = "My Chart"
        c.series do |s|
          s.values = [1, 2, 3]
        end
      end

      # image
      sheet.image("binary_data", anchor: "E1", width: 100, height: 50)

      # styles
      sheet.style("mystyle", bold: true, font_size: 14)
      sheet.row(["Styled"], cell_styles: ["mystyle"])
    end

    assert_instance_of Xlsxrb::Elements::Workbook, wb.build
  end

  def test_stream_writer_public_methods
    require "stringio"
    io = StringIO.new

    # Test block format
    Xlsxrb.write(io, strict_excel_mode: false) do |writer|
      writer.sheet("Sheet1") do |sheet|
        sheet.row(%w[A1 B1])
        sheet.row([])
        sheet.row([1, 2, 3])
      end

      writer.sheet("Sheet2")
    end

    assert io.size.positive?

    # Test manual open/close
    io2 = StringIO.new
    sw = Xlsxrb::StreamWriter.new(io2, strict_excel_mode: false)
    sw.sheet("Sheet1") do |s|
      s.row([1, 2])
    end
    sw.close
    sw.cleanup!
    assert io2.size.positive?
  end

  def test_chart_builder
    cb = Xlsxrb::ChartBuilder.new
    cb.type = "line"
    cb.title = "Sales"
    cb.legend position: "bottom"
    cb.plot_area(layout: "default")
    cb.chart_space(rounded_corners: true)
    cb.style = 2
    cb.data_labels(show_val: true)
    cb.plot_visible_only = true
    cb.display_blanks_as = "gap"
    cb.view3d(rot_x: 30)
    cb.category_axis(title: "Months")
    cb.value_axis(title: "Revenue")
    cb.show_legend_key = true

    cb.series do |s|
      s.name = "Q1"
      s.categories = %w[Jan Feb]
      s.values = [10, 20]
      s.marker symbol: "circle"
      s.fill type: "solid"
      s.line color: "FF0000"
      s.trendline type: "linear"
      s.data_labels show_val: false
      s.smooth = true
      s.shape = "cone"
      s.type = "bar"
    end

    assert_equal "line", cb.options[:type]
    assert_equal "Sales", cb.options[:title]
    assert_equal 1, cb.options[:series].size
    assert_equal "Q1", cb.options[:series].first[:name]
  end

  def test_ooxml_writer
    require "tempfile"
    Tempfile.create(["test", ".xlsx"]) do |f|
      Xlsxrb::Ooxml::WorkbookWriter.write(
        f.path,
        sheets: [{ name: "Sheet1", rows: [], columns: [] }],
        shared_strings: [],
        shared_strings_index: {},
        styles: {}
      )
      assert File.exist?(f.path)
    end
  end

  test "first-class conditional_formatting and data_validations in Worksheet roundtrip" do
    wb = Xlsxrb.build do |w|
      w.sheet("Sheet1") do |s|
        s.row([50, 5])
        s.conditional_format("A1:A10", type: :cellIs, operator: :greaterThan, formula: "100")
        s.data_validation("B1:B10", type: :whole, operator: :between, formula1: "1", formula2: "10")
      end
    end

    ws = wb.sheet(0)
    assert_equal 1, ws.conditional_formats.size
    assert_equal 1, ws.data_validations.size

    Tempfile.create(["cf_dv_test", ".xlsx"]) do |f|
      Xlsxrb.write(f.path, wb)
      reloaded_wb = Xlsxrb.load(f.path)
      reloaded_ws = reloaded_wb.sheet(0)
      assert_equal 1, reloaded_ws.conditional_formats.size
      assert_equal 1, reloaded_ws.data_validations.size
      assert_equal "A1:A10", reloaded_ws.conditional_formats.first.sqref
      assert_equal "B1:B10", reloaded_ws.data_validations.first.sqref
    end
  end

  test "Xlsxrb.read loads merged_cells and auto_filter onto Elements::Worksheet" do
    Tempfile.create(["merge_filter_test", ".xlsx"]) do |f|
      Xlsxrb.write(f.path) do |wb|
        wb.sheet("Sales") do |s|
          s.row(%w[Header1 Header2 Header3])
          s.row([1, 2, 3])
          s.merge("A1:B1")
          s.auto_filter("A1:C2")
        end
      end

      # StreamSheet via block
      Xlsxrb.read(f.path) do |stream_sheet|
        assert_equal(["A1:B1"], stream_sheet.merged_cells)
        assert_equal("A1:C2", stream_sheet.auto_filter)

        # Elements::Worksheet via load
        loaded_ws = stream_sheet.load
        assert_kind_of(Xlsxrb::Elements::Worksheet, loaded_ws)
        assert_equal(["A1:B1"], loaded_ws.merged_cells)
        assert_equal("A1:C2", loaded_ws.auto_filter)
      end

      # Elements::Workbook via load
      loaded_wb = Xlsxrb.load(f.path)
      assert_equal(["A1:B1"], loaded_wb.sheet(0).merged_cells)
      assert_equal("A1:C2", loaded_wb.sheet(0).auto_filter)
    end
  end

  test "Xlsxrb top-level date and time serial conversion utilities" do
    d = Date.new(2026, 1, 1)
    s = Xlsxrb.date_to_serial(d)
    assert_equal(46_023, s)
    assert_equal(d, Xlsxrb.serial_to_date(s))

    t = Time.utc(2026, 1, 1, 12, 0, 0)
    st = Xlsxrb.time_to_serial(t)
    assert_in_delta(46_023.5, st, 1e-6)
    assert_equal(t, Xlsxrb.serial_to_time(st))

    # 1904 date system
    s04 = Xlsxrb.date_to_serial(d, date1904: true)
    assert_equal(44_561, s04)
    assert_equal(d, Xlsxrb.serial_to_date(s04, date1904: true))

    st04 = Xlsxrb.time_to_serial(t, date1904: true)
    assert_in_delta(44_561.5, st04, 1e-6)
    assert_equal(t, Xlsxrb.serial_to_time(st04, date1904: true))

    # valid_excel_date? validation
    assert_true(Xlsxrb.valid_excel_date?(s))
    assert_false(Xlsxrb.valid_excel_date?(-1))
    assert_true(Xlsxrb.valid_excel_date?(s04, date1904: true))
    assert_false(Xlsxrb.valid_excel_date?(3_000_000, date1904: true))
  end

  test "Xlsxrb.write returns in-memory binary string when target is omitted" do
    # Streaming block with omitted target
    data1 = Xlsxrb.write do |w|
      w.sheet("Items") do |s|
        s.row(%w[Header1 Header2])
        s << ["Val1", 10] << ["Val2", 20]
      end
    end

    assert_kind_of(String, data1)
    assert(data1.start_with?("PK\x03\x04"))
    wb1 = Xlsxrb.load(data1)
    sheet1 = wb1.sheet(0)
    assert_equal("Items", sheet1.name)
    assert_equal("Header1", sheet1["A1"]&.value)
    assert_equal("Val1", sheet1["A2"]&.value)
    assert_equal(10, sheet1["B2"]&.value)
    assert_equal("Val2", sheet1["A3"]&.value)
    assert_equal(20, sheet1["B3"]&.value)

    # In-memory workbook with omitted target
    wb_in = Xlsxrb.build do |w|
      w.sheet("DOM") do |s|
        s.row(["DOM_A", 100])
        s << ["DOM_B", 200] << ["DOM_C", 300]
      end
    end
    data2 = Xlsxrb.write(wb_in)
    assert_kind_of(String, data2)
    assert(data2.start_with?("PK\x03\x04"))
    wb2 = Xlsxrb.load(data2)
    sheet2 = wb2.sheet(0)
    assert_equal("DOM", sheet2.name)
    assert_equal("DOM_A", sheet2["A1"]&.value)
    assert_equal(100, sheet2["B1"]&.value)
    assert_equal("DOM_B", sheet2["A2"]&.value)
    assert_equal(200, sheet2["B2"]&.value)
    assert_equal("DOM_C", sheet2["A3"]&.value)
    assert_equal(300, sheet2["B3"]&.value)
  end

  test "sheet_format configured via build and streaming write round-trips correctly" do
    # 1. Via Xlsxrb.build
    wb_built = Xlsxrb.build do |w|
      w.sheet("Formatted") do |s|
        s.sheet_format(default_row_height: 24.5, default_column_width: 18.0)
        s.sheet_format(:outline_level_row, 2)
        s.row(["Header 1", "Header 2"])
      end
    end

    assert_equal(24.5, wb_built.sheet(0).sheet_format.default_row_height)
    assert_equal(18.0, wb_built.sheet(0).sheet_format.default_column_width)
    assert_equal(2, wb_built.sheet(0).sheet_format.outline_level_row)

    # Write and read back
    data = Xlsxrb.write(wb_built)
    stream_sheet = Xlsxrb.read(data).first
    assert_equal(24.5, stream_sheet.sheet_format.default_row_height)
    assert_equal(18.0, stream_sheet.sheet_format.default_column_width)
    assert_equal(2, stream_sheet.sheet_format.outline_level_row)

    loaded_sheet = stream_sheet.load
    assert_equal(24.5, loaded_sheet.sheet_format.default_row_height)
    assert_equal(18.0, loaded_sheet.sheet_format.default_column_width)
    assert_equal(2, loaded_sheet.sheet_format.outline_level_row)

    # 2. Via streaming Xlsxrb.write
    stream_data = Xlsxrb.write do |w|
      w.sheet("StreamFormat") do |s|
        s.sheet_format(default_row_height: 30.0, default_column_width: 12.5)
        s.row(%w[Val1 Val2])
      end
    end

    stream_sheet2 = Xlsxrb.read(stream_data).first
    assert_equal(30.0, stream_sheet2.sheet_format.default_row_height)
    assert_equal(12.5, stream_sheet2.sheet_format.default_column_width)
    assert_equal(30.0, stream_sheet2.load.sheet_format.default_row_height)
    assert_equal(12.5, stream_sheet2.load.sheet_format.default_column_width)
  end

  test "protected_range and cell_watch configured via build and streaming write round-trips correctly" do
    # 1. Via Xlsxrb.build
    wb_built = Xlsxrb.build do |w|
      w.sheet("ProtectedSheet") do |s|
        s.protected_range("Range1", "A1:B10", password: "CCEE")
        s.protected_range("Range2", "C1:D5", password: "modern_password")
        s.cell_watch("A1")
        s.cell_watch("B10")
        s.row(%w[A B C D])
      end
    end

    built_sheet = wb_built.sheet(0)
    assert_equal(2, built_sheet.protected_ranges.size)
    assert_equal("Range1", built_sheet.protected_ranges[0].name)
    assert_equal("A1:B10", built_sheet.protected_ranges[0].sqref)
    assert_equal("CCEE", built_sheet.protected_ranges[0].password)
    assert_equal("Range2", built_sheet.protected_ranges[1].name)
    assert_equal("C1:D5", built_sheet.protected_ranges[1].sqref)
    assert_equal("SHA-512", built_sheet.protected_ranges[1].algorithm_name)
    assert_not_nil(built_sheet.protected_ranges[1].hash_value)

    assert_equal(%w[A1 B10], built_sheet.cell_watches)

    # Write and read back
    data = Xlsxrb.write(wb_built)
    stream_sheet = Xlsxrb.read(data).first

    assert_equal(2, stream_sheet.protected_ranges.size)
    assert_equal("Range1", stream_sheet.protected_ranges[0].name)
    assert_equal("A1:B10", stream_sheet.protected_ranges[0].sqref)
    assert_equal("Range2", stream_sheet.protected_ranges[1].name)
    assert_equal("C1:D5", stream_sheet.protected_ranges[1].sqref)
    assert_equal("SHA-512", stream_sheet.protected_ranges[1].algorithm_name)
    assert_equal(%w[A1 B10], stream_sheet.cell_watches)

    loaded_sheet = stream_sheet.load
    assert_equal(2, loaded_sheet.protected_ranges.size)
    assert_equal("Range1", loaded_sheet.protected_ranges[0].name)
    assert_equal("A1:B10", loaded_sheet.protected_ranges[0].sqref)
    assert_equal(%w[A1 B10], loaded_sheet.cell_watches)

    # 2. Via streaming Xlsxrb.write
    stream_data = Xlsxrb.write do |w|
      w.sheet("StreamProtected") do |s|
        s.protected_range("StreamRange", "B2:F20")
        s.cell_watch("C3")
        s.row(%w[Stream Row])
      end
    end

    ss = Xlsxrb.read(stream_data).first
    assert_equal(1, ss.protected_ranges.size)
    assert_equal("StreamRange", ss.protected_ranges[0].name)
    assert_equal("B2:F20", ss.protected_ranges[0].sqref)
    assert_equal(["C3"], ss.cell_watches)

    ls = ss.load
    assert_equal(1, ls.protected_ranges.size)
    assert_equal("StreamRange", ls.protected_ranges[0].name)
    assert_equal(["C3"], ls.cell_watches)
  end

  test "table and data_validation block forms configure options and round-trip correctly" do
    # 1. Via Xlsxrb.build
    wb_built = Xlsxrb.build do |w|
      w.sheet("BlockConfig") do |s|
        s.table("A1:C5") do |t|
          t.column "ID"
          t.column "Name"
          t.column "Score"
          t.name = "ScoreTable"
          t.style = "TableStyleMedium2"
          t.show_row_stripes = true
        end

        s.data_validation("B2:B5") do |dv|
          dv.type = :list
          dv.formula1 = '"Alice,Bob,Charlie"'
          dv.allow_blank = true
          dv.error_title = "Invalid Name"
          dv.error_message = "Select a valid student name"
        end

        s.row(%w[ID Name Score])
        s.row([1, "Alice", 95])
        s.row([2, "Bob", 88])
      end
    end

    built_sheet = wb_built.sheet(0)
    assert_equal(1, built_sheet.tables.size)
    table_cfg = built_sheet.tables.first
    assert_equal("A1:C5", table_cfg.ref)
    assert_equal(%w[ID Name Score], table_cfg.columns.map(&:name))
    assert_equal("ScoreTable", table_cfg.name)
    assert_equal("TableStyleMedium2", table_cfg.style_name)
    assert_equal(true, table_cfg.show_row_stripes?)

    assert_equal(1, built_sheet.data_validations.size)
    dv_cfg = built_sheet.data_validations.first
    assert_equal("B2:B5", dv_cfg.sqref)
    assert_equal(:list, dv_cfg.type)
    assert_equal('"Alice,Bob,Charlie"', dv_cfg.formula1)
    assert_equal(true, dv_cfg.allow_blank)
    assert_equal("Invalid Name", dv_cfg.error_title)
    assert_equal("Select a valid student name", dv_cfg.error)

    # Write and read back
    data = Xlsxrb.write(wb_built)
    stream_sheet = Xlsxrb.read(data).first

    assert_equal(1, stream_sheet.tables.size)
    assert_equal("ScoreTable", stream_sheet.tables.first.name)
    assert_equal("A1:C5", stream_sheet.tables.first.ref)

    assert_equal(1, stream_sheet.data_validations.size)
    assert_equal("B2:B5", stream_sheet.data_validations.first.sqref)
    assert_equal("list", stream_sheet.data_validations.first.type)

    # 2. Via streaming Xlsxrb.write
    stream_data = Xlsxrb.write do |w|
      w.sheet("StreamBlock") do |s|
        s.table("A1:B3") do |t|
          t.columns = %w[Key Value]
          t.style = "TableStyleLight1"
        end
        s.data_validation("B2:B3") do |dv|
          dv.type = :whole
          dv.operator = :greaterThan
          dv.formula1 = "0"
        end
        s.row(%w[Key Value])
        s.row(["K1", 10])
      end
    end

    wb_stream_sheet = Xlsxrb.read(stream_data).first
    assert_equal(1, wb_stream_sheet.tables.size)
    assert_equal("A1:B3", wb_stream_sheet.tables.first.ref)
    assert_equal(1, wb_stream_sheet.data_validations.size)
    assert_equal("B2:B3", wb_stream_sheet.data_validations.first.sqref)

    # 3. Validation errors
    assert_raise(ArgumentError) do
      Xlsxrb.build { |w| w.sheet("Err") { |s| s.table("A1:B2") { |_t| nil } } }
    end
    assert_raise(ArgumentError) do
      Xlsxrb.build { |w| w.sheet("Err") { |s| s.data_validation { |_dv| nil } } }
    end
  end

  test "Xlsxrb.load and Ooxml::Reader exclude chartsheets from worksheet enumeration" do
    buf = StringIO.new
    buf.binmode
    Xlsxrb::Ooxml::ZipWriter.open(buf) do |zw|
      zw.add_entry("[Content_Types].xml", <<~XML)
        <?xml version="1.0" encoding="UTF-8"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/chartsheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/>
          <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
        </Types>
      XML
      zw.add_entry("_rels/.rels", <<~XML)
        <?xml version="1.0" encoding="UTF-8"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
      XML
      zw.add_entry("xl/workbook.xml", <<~XML)
        <?xml version="1.0" encoding="UTF-8"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets>
            <sheet name="Chart1" sheetId="1" r:id="rId1"/>
            <sheet name="DataSheet" sheetId="2" r:id="rId2"/>
          </sheets>
        </workbook>
      XML
      zw.add_entry("xl/_rels/workbook.xml.rels", <<~XML)
        <?xml version="1.0" encoding="UTF-8"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet1.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
        </Relationships>
      XML
      zw.add_entry("xl/chartsheets/sheet1.xml", <<~XML)
        <?xml version="1.0" encoding="UTF-8"?>
        <chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetViews><sheetView workbookViewId="0"/></sheetViews>
          <drawing r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
        </chartsheet>
      XML
      zw.add_entry("xl/worksheets/sheet1.xml", <<~XML)
        <?xml version="1.0" encoding="UTF-8"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData>
            <row r="1"><c r="A1"><v>42</v></c></row>
          </sheetData>
        </worksheet>
      XML
    end

    raw_bytes = buf.string.b

    # 1. Xlsxrb.load
    wb = Xlsxrb.load(raw_bytes)
    assert_equal(["DataSheet"], wb.sheets.map(&:name))
    assert_equal([[42]], wb.sheet("DataSheet").each_row_values.to_a)
    assert_equal(1, wb.unmapped_data[:non_worksheet_sheets].size)
    assert_equal("Chart1", wb.unmapped_data[:non_worksheet_sheets].first[:name])

    # 2. Ooxml::Reader
    Tempfile.create(["chartsheet_test", ".xlsx"]) do |tmp|
      tmp.binmode
      tmp.write(raw_bytes)
      tmp.flush

      reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
      assert_equal(["DataSheet"], reader.sheet_names)
      assert_equal({ "A1" => 42 }, reader.cells)
    end
  end

  # --- sheet_name & Validation ---

  test "Xlsxrb.sanitize_sheet_name public facade" do
    assert_equal("Sales_2026", Xlsxrb.sanitize_sheet_name("Sales:2026"))
    assert_equal("History_", Xlsxrb.sanitize_sheet_name("History"))
  end

  test "StreamWriter#sheet_names and WorksheetProxy#sheet_names report workbook sheets" do
    io = StringIO.new
    Xlsxrb.write(io) do |w|
      assert_equal([], w.sheet_names)
      w.sheet("Sheet1") do |s|
        s.row(["Hello"])
        assert_equal(["Sheet1"], s.sheet_names)
      end
      assert_equal(["Sheet1"], w.sheet_names)

      name = Xlsxrb.sanitize_sheet_name("Sales: 2026")
      assert_equal("Sales_ 2026", name)
      w.sheet(name) do |s|
        s.row([1])
        assert_equal(["Sheet1", "Sales_ 2026"], s.sheet_names)
      end
      assert_equal(["Sheet1", "Sales_ 2026"], w.sheet_names)
    end

    wb = Xlsxrb.load(io.string.b)
    assert_equal(["Sheet1", "Sales_ 2026"], wb.sheets.map(&:name))
  end

  test "WorkbookBuilder#sheet_names reports in-memory workbook sheets" do
    wb = Xlsxrb.build do |w|
      assert_equal([], w.sheet_names)
      w.sheet("Summary") do |s|
        s.row([1])
        assert_equal(["Summary"], w.sheet_names)
      end
      assert_equal(["Summary"], w.sheet_names)

      name = Xlsxrb.sanitize_sheet_name("Data [Q1]")
      assert_equal("Data _Q1_", name)
      w.sheet(name) { |s| s.row([2]) }
      assert_equal(["Summary", "Data _Q1_"], w.sheet_names)
    end
    assert_equal(["Summary", "Data _Q1_"], wb.sheets.map(&:name))
  end

  test "StreamWriter#sheet and WorkbookBuilder#sheet enforce Excel restrictions on colon, quotes, duplicates, and History" do
    # StreamWriter
    assert_raises(ArgumentError) do
      Xlsxrb.write(StringIO.new) { |w| w.sheet("Bad:Name") }
    end
    assert_raises(ArgumentError) do
      Xlsxrb.write(StringIO.new) { |w| w.sheet("'BadQuote") }
    end
    assert_raises(ArgumentError) do
      Xlsxrb.write(StringIO.new) { |w| w.sheet("BadQuote'") }
    end
    assert_raises(ArgumentError) do
      Xlsxrb.write(StringIO.new) { |w| w.sheet("History") }
    end
    assert_raises(ArgumentError) do
      Xlsxrb.write(StringIO.new) { |w| w.sheet("history") }
    end
    assert_raises(ArgumentError) do
      Xlsxrb.write(StringIO.new) do |w|
        w.sheet("Sales") { |s| s.row([1]) }
        w.sheet("sales") { |s| s.row([2]) }
      end
    end

    # Default consecutive sheet numbering in StreamWriter
    default_io = StringIO.new
    Xlsxrb.write(default_io) do |w|
      w.sheet { |s| s.row([1]) }
      w.sheet { |s| s.row([2]) }
    end
    assert_equal(%w[Sheet1 Sheet2], Xlsxrb.load(default_io.string.b).sheets.map(&:name))

    # WorkbookBuilder
    assert_raises(ArgumentError) do
      Xlsxrb.build { |w| w.sheet("Bad:Name") }
    end
    assert_raises(ArgumentError) do
      Xlsxrb.build { |w| w.sheet("'BadQuote") }
    end
    assert_raises(ArgumentError) do
      Xlsxrb.build { |w| w.sheet("BadQuote'") }
    end
    assert_raises(ArgumentError) do
      Xlsxrb.build { |w| w.sheet("History") }
    end
    assert_raises(ArgumentError) do
      Xlsxrb.build { |w| w.sheet("history") }
    end
    assert_raises(ArgumentError) do
      Xlsxrb.build do |w|
        w.sheet("Sales") { |s| s.row([1]) }
        w.sheet("sales") { |s| s.row([2]) }
      end
    end
  end

  test "builder getter and setter methods are separated and return assigned values" do
    # 1. WorkbookBuilder
    Xlsxrb.build do |b|
      # active_sheet
      b.sheet("Sheet1") { |s| s.row([1]) }
      b.sheet("Sheet2") { |s| s.row([2]) }

      assert_nil b.active_sheet
      assert_equal "Sheet1", (b.active_sheet = "Sheet1")
      assert_equal "Sheet1", b.active_sheet
      assert_nil(b.active_sheet = nil)
      assert_nil b.active_sheet
      assert_equal "Sheet1", b.send(:active_sheet=, "Sheet1")
      assert_equal "Sheet1", b.active_sheet

      # first_sheet
      assert_nil b.first_sheet
      assert_equal "Sheet2", (b.first_sheet = "Sheet2")
      assert_equal "Sheet2", b.first_sheet
      assert_nil(b.first_sheet = nil)
      assert_nil b.first_sheet
      assert_equal "Sheet1", b.send(:first_sheet=, "Sheet1")
      assert_equal "Sheet1", b.first_sheet

      # show_sheet_tabs
      assert_nil b.show_sheet_tabs
      assert_equal false, (b.show_sheet_tabs = false)
      assert_equal false, b.show_sheet_tabs
      assert_nil(b.show_sheet_tabs = nil)
      assert_nil b.show_sheet_tabs
      assert_equal false, b.send(:show_sheet_tabs=, false)

      # tab_ratio
      assert_nil b.tab_ratio
      assert_equal 500, (b.tab_ratio = 500)
      assert_equal 500, b.tab_ratio
      assert_nil(b.tab_ratio = nil)
      assert_nil b.tab_ratio
      assert_equal 600, b.send(:tab_ratio=, 600)

      # code_name
      assert_nil b.code_name
      assert_equal "WBCode", (b.code_name = "WBCode")
      assert_equal "WBCode", b.code_name
      assert_nil(b.code_name = nil)
      assert_nil b.code_name
      assert_equal "WBCode3", b.send(:code_name=, "WBCode3")
    end

    # 2. WorksheetBuilder
    Xlsxrb.build do |b|
      b.sheet("Sheet1") do |s|
        # tab_color
        assert_nil s.tab_color
        assert_equal "FF0000", (s.tab_color = "FF0000")
        assert_equal "FF0000", s.tab_color
        assert_nil(s.tab_color = nil)
        assert_nil s.tab_color
        assert_equal "0000FF", s.send(:tab_color=, "0000FF")

        # fit_to_page
        assert_equal false, s.fit_to_page
        assert_equal true, (s.fit_to_page = true)
        assert_equal true, s.fit_to_page
        assert_equal false, (s.fit_to_page = false)
        assert_equal false, s.fit_to_page
        assert_nil(s.fit_to_page = nil)
        assert_equal false, s.fit_to_page
        assert_equal true, s.send(:fit_to_page=, true)

        # code_name
        assert_nil s.code_name
        assert_equal "WSCode", (s.code_name = "WSCode")
        assert_equal "WSCode", s.code_name
        assert_nil(s.code_name = nil)
        assert_nil s.code_name
        assert_equal "WSCode3", s.send(:code_name=, "WSCode3")

        # sheet_calc_properties
        assert_nil s.sheet_calc_properties
        calc_hash = { full_calc_on_load: true }
        assert_equal calc_hash, (s.sheet_calc_properties = calc_hash)
        assert_equal Xlsxrb::Elements::SheetCalcProperties.new(**calc_hash), s.sheet_calc_properties
        assert_nil(s.sheet_calc_properties = nil)
        assert_nil s.sheet_calc_properties
        assert_equal calc_hash, s.send(:sheet_calc_properties=, calc_hash)

        # phonetic_properties
        assert_nil s.phonetic_properties
        props = { font_id: 2, type: :hiragana }
        expected_props = { font_id: 2, type: "Hiragana" }
        assert_equal props, (s.phonetic_properties = props)
        assert_equal Xlsxrb::Elements::PhoneticProperties.new(**expected_props), s.phonetic_properties
        assert_nil(s.phonetic_properties = nil)
        assert_nil s.phonetic_properties
        assert_equal props, s.send(:phonetic_properties=, props)

        # background_image
        assert_nil s.background_image
        assert_nil(s.background_image = nil)
        assert_nil s.background_image
        assert_nil s.send(:background_image=, nil)

        # state
        assert_equal :visible, s.state
        assert_equal :hidden, (s.state = :hidden)
        assert_equal :hidden, s.state
        assert_equal :very_hidden, (s.state = :very_hidden)
        assert_equal :very_hidden, s.state
        assert_equal "visible", (s.state = "visible")
        assert_equal :visible, s.state
        assert_equal :visible, s.send(:state=, "visible")
        assert_raises(ArgumentError) { s.state = :invalid_state }

        # TableBuilder totals_row / totals_row_count
        s.row([1, 2])
        s.table("A1:B1", columns: %w[A B]) do |t|
          assert_nil t.totals_row_count
          assert_nil t.totals_row
          assert_equal true, (t.totals_row = true)
          assert_equal true, t.totals_row
          assert_equal false, (t.totals_row = false)
          assert_equal false, t.totals_row
          assert_equal true, t.send(:totals_row=, true)
          assert_equal false, t.send(:totals_row=, false)

          assert_equal 0, t.totals_row_count
          assert_equal 1, (t.totals_row_count = 1)
          assert_equal 1, t.totals_row_count
          assert_equal 1, t.send(:totals_row_count=, 1)
        end
      end
    end

    # 3. StreamWriter
    buf = StringIO.new
    Xlsxrb.write(buf) do |w|
      # workbook-level setters on StreamWriter
      assert_nil w.active_sheet
      assert_equal "S1", (w.active_sheet = "S1")
      assert_equal "S1", w.active_sheet
      assert_nil(w.active_sheet = nil)
      assert_nil w.active_sheet
      assert_equal "S2", w.send(:active_sheet=, "S2")
      assert_equal "S1", w.send(:active_sheet=, "S1")

      assert_nil w.first_sheet
      assert_equal "S1", (w.first_sheet = "S1")
      assert_nil(w.first_sheet = nil)
      assert_nil w.first_sheet
      assert_equal "S1", w.send(:first_sheet=, "S1")

      assert_nil w.show_sheet_tabs
      assert_equal false, (w.show_sheet_tabs = false)
      assert_nil(w.show_sheet_tabs = nil)
      assert_nil w.show_sheet_tabs
      assert_equal true, w.send(:show_sheet_tabs=, true)

      assert_nil w.tab_ratio
      assert_equal 400, (w.tab_ratio = 400)
      assert_nil(w.tab_ratio = nil)
      assert_nil w.tab_ratio
      assert_equal 500, w.send(:tab_ratio=, 500)

      # sheet proxy setters
      w.sheet("S1") do |s|
        assert_nil s.tab_color
        assert_equal "AABBCC", (s.tab_color = "AABBCC")
        assert_equal "AABBCC", s.tab_color
        assert_nil(s.tab_color = nil)
        assert_nil s.tab_color
        assert_equal "112233", s.send(:tab_color=, "112233")

        assert_equal false, s.fit_to_page
        assert_equal true, (s.fit_to_page = true)
        assert_equal true, s.fit_to_page
        assert_equal false, (s.fit_to_page = false)
        assert_equal false, s.fit_to_page
        assert_nil(s.fit_to_page = nil)
        assert_equal false, s.fit_to_page
        assert_equal true, s.send(:fit_to_page=, true)

        assert_nil s.code_name
        assert_equal "StreamCode", (s.code_name = "StreamCode")
        assert_equal "StreamCode", s.code_name
        assert_nil(s.code_name = nil)
        assert_nil s.code_name
        assert_equal "StreamCode2", s.send(:code_name=, "StreamCode2")

        calc_hash = { full_calc_on_load: true }
        assert_nil s.sheet_calc_properties
        assert_equal calc_hash, (s.sheet_calc_properties = calc_hash)
        assert_equal Xlsxrb::Elements::SheetCalcProperties.new(**calc_hash), s.sheet_calc_properties
        assert_nil(s.sheet_calc_properties = nil)
        assert_nil s.sheet_calc_properties
        assert_equal calc_hash, s.send(:sheet_calc_properties=, calc_hash)

        props = { font_id: 3, type: :hiragana }
        expected_props = { font_id: 3, type: "Hiragana" }
        assert_nil s.phonetic_properties
        assert_equal props, (s.phonetic_properties = props)
        assert_equal Xlsxrb::Elements::PhoneticProperties.new(**expected_props), s.phonetic_properties
        assert_nil(s.phonetic_properties = nil)
        assert_nil s.phonetic_properties
        assert_equal props, s.send(:phonetic_properties=, props)

        assert_nil s.background_image
        assert_nil(s.background_image = nil)
        assert_nil s.background_image
        assert_nil s.send(:background_image=, nil)

        # state on WorksheetProxy
        assert_equal :visible, s.state
        assert_equal :hidden, (s.state = :hidden)
        assert_equal :hidden, s.state
        assert_equal :very_hidden, (s.state = :very_hidden)
        assert_equal :very_hidden, s.state
        assert_equal "visible", (s.state = "visible")
        assert_equal :visible, s.state
        assert_equal :visible, s.send(:state=, "visible")
        assert_raises(ArgumentError) { s.state = :invalid_state }

        # state on StreamWriter directly
        assert_equal :hidden, (w.state = :hidden)
        assert_equal :hidden, w.state
        assert_equal "very_hidden", (w.state = "very_hidden")
        assert_equal :very_hidden, w.state
        assert_equal :very_hidden, w.send(:state=, "very_hidden")
        assert_raises(ArgumentError) { w.state = :invalid_state }
        assert_equal :visible, (w.state = :visible)

        s.row([1])
      end
    end
  end

  test "state= configures sheet visibility on WorksheetBuilder, StreamWriter, and WorksheetProxy and roundtrips" do
    wb1 = Xlsxrb.build do |b|
      b.sheet("VisibleSheet") { |s| s.row([1]) }
      b.sheet("HiddenSheet") do |s|
        s.state = :hidden
        s.row([2])
      end
      b.sheet("VeryHiddenSheet") do |s|
        s.state = :very_hidden
        s.row([3])
      end
    end
    assert_equal :visible, wb1.sheet(0).state
    assert_equal :hidden, wb1.sheet(1).state
    assert_equal :very_hidden, wb1.sheet(2).state

    data1 = Xlsxrb.write(wb1)
    wb1_read = Xlsxrb.load(data1)
    assert_equal :visible, wb1_read.sheet(0).state
    assert_equal :hidden, wb1_read.sheet(1).state
    assert_equal :very_hidden, wb1_read.sheet(2).state

    data2 = Xlsxrb.write do |w|
      w.sheet("VisStream") { |s| s.row([1]) }
      w.sheet("HiddenProxy") do |s|
        s.state = :hidden
        s.row([2])
      end
      w.sheet("VeryHiddenStream")
      w.state = :very_hidden
      w.row([3])
    end
    wb2 = Xlsxrb.load(data2)
    assert_equal :visible, wb2.sheet(0).state
    assert_equal :hidden, wb2.sheet(1).state
    assert_equal :very_hidden, wb2.sheet(2).state
  end

  test "merge accepts coordinate range string, row_range and column_range report 0-based bounds, and bracket access supports 2D indexing" do
    wb_dom = Xlsxrb.build do |b|
      b.sheet("Grid") do |s|
        s.row([10, 20, 30])
        s.skip_rows(1)
        s.row([nil, 50])
        s.merge("A1:C1")
      end
    end
    data = Xlsxrb.write(wb_dom)

    wb = Xlsxrb.load(data)
    sheet = wb.sheet(0)

    # Merged cells roundtrip from range string
    assert_equal(["A1:C1"], sheet.merged_cells)

    # row_range and column_range
    assert_equal(0..2, sheet.row_range)
    assert_equal(0..2, sheet.column_range)

    # 2D bracket access and coordinate string access
    assert_equal(10, sheet[0, 0].value)
    assert_equal(20, sheet[0, 1].value)
    assert_equal(10, sheet[0, "A"].value)
    assert_equal(20, sheet[0, "B"].value)
    assert_equal(10, sheet["A1"].value)
    assert_equal(20, sheet["B1"].value)
    assert_nil(sheet[1, 0])
    assert_equal(50, sheet[2, 1].value)
    assert_equal(50, sheet[2, "B"].value)
    assert_equal(50, sheet["B3"].value)

    # Sparse row bracket access: row[col] accesses column index, not sparse array index
    row3 = sheet.row_at(2)
    assert_nil(row3[0])
    assert_equal(50, row3[1].value)
    assert_nil(row3["A"])
    assert_equal(50, row3["B"].value)
    assert_equal([nil, 50], row3.values)

    # StreamSheet dimension bounds and row_range/column_range (nil when no <dimension> tag is present)
    Xlsxrb.read(data) do |stream_sheet|
      assert_nil(stream_sheet.dimension)
      assert_nil(stream_sheet.row_range)
      assert_nil(stream_sheet.column_range)
      assert_equal(["A1:C1"], stream_sheet.merged_cells)
      assert_not_respond_to(stream_sheet, :[])
    end
  end

  def capture_stderr
    old_stderr = $stderr
    old_verbose = $VERBOSE
    $VERBOSE = false
    $stderr = StringIO.new
    yield
    $stderr.string
  ensure
    $stderr = old_stderr
    $VERBOSE = old_verbose
  end

  test "Commit 2: top-level Xlsxrb coordinate and date facade" do
    assert_equal(16_384, Xlsxrb::MAX_COLUMN)
    assert_equal([0, 0], Xlsxrb.parse_ref("A1"))
    assert_equal([9, 2], Xlsxrb.parse_ref("C10"))
    assert_equal("A1", Xlsxrb.cell_to_ref(0, 0))
    assert_equal("C10", Xlsxrb.cell_to_ref(9, 2))
    assert_equal(0, Xlsxrb.column_index("A"))
    assert_equal(26, Xlsxrb.column_index("AA"))
    assert_equal("A", Xlsxrb.column_letter(0))
    assert_equal("AA", Xlsxrb.column_letter(26))
    assert_equal([[0, 0], [9, 2]], Xlsxrb.parse_range("A1:C10"))
    assert_equal("A1:C10", Xlsxrb.range_to_ref(0, 0, 9, 2))
    assert_equal("A1:C10", Xlsxrb.range_to_ref([[0, 0], [9, 2]]))
    assert_equal("A1", Xlsxrb.range_to_ref(0, 0, 0, 0))
    assert_equal("A1", Xlsxrb.range_to_ref([[0, 0], [0, 0]]))

    assert_equal("100.00", Xlsxrb.format_value(100, "0.00"))
    t = Time.utc(2026, 1, 1, 12, 0, 0)
    serial = Xlsxrb.time_to_serial(t)
    assert_in_delta(46_023.5, serial, 0.0001)
    assert_equal(t, Xlsxrb.serial_to_time(serial))

    # Formula helpers
    sf = Xlsxrb.shared_formula("A1*2", shared_index: 1, ref: "B1:B10")
    assert_equal(:shared, sf.type)
    assert_equal(1, sf.shared_index)

    af1 = Xlsxrb.array_formula("SUM(A1:A10)", ref: "B1")
    assert_equal(:array, af1.type)
    assert_equal("B1", af1.ref)

    af2 = Xlsxrb.array_formula("SUM(A1:A10)", ref: "B1", cached_value: Xlsxrb::Elements::CellError::VALUE)
    assert_equal(:array, af2.type)
    assert_equal("B1", af2.ref)
    assert_equal("#VALUE!", af2.cached_value.to_s)
  end

  test "Commit 2: Elements::CellError constants" do
    assert_equal("#NULL!", Xlsxrb::Elements::CellError::NULL.code)
    assert_equal("#DIV/0!", Xlsxrb::Elements::CellError::DIV0.code)
    assert_equal("#VALUE!", Xlsxrb::Elements::CellError::VALUE.code)
    assert_equal("#REF!", Xlsxrb::Elements::CellError::REF.code)
    assert_equal("#NAME?", Xlsxrb::Elements::CellError::NAME.code)
    assert_equal("#NUM!", Xlsxrb::Elements::CellError::NUM.code)
    assert_equal("#N/A", Xlsxrb::Elements::CellError::NA.code)
    assert_equal("#GETTING_DATA", Xlsxrb::Elements::CellError::GETTING_DATA.code)
  end

  test "Commit 3: cell_styles and row_style in WorksheetBuilder#row with deprecation warnings" do
    builder = Xlsxrb::WorksheetBuilder.new("Sheet1")
    builder.style(:b, bold: true)
    builder.style(:r, font_color: "FF0000")

    # Canonical cell_styles and row_style
    builder.row([1, 2], cell_styles: [:b, nil], row_style: :r)
    rows = builder.instance_variable_get(:@rows)
    row = rows.last
    assert_equal(:r, row.unmapped_data[:style_index])
    assert_equal(:b, row.cells[0].style)
  end

  test "cell_styles and row_style in StreamWriter#row and WorksheetProxy#row" do
    io = StringIO.new
    writer = Xlsxrb::StreamWriter.new(io)
    writer.style(:b, bold: true)
    writer.style(:r, font_color: "FF0000")

    # Proxy with cell_styles and row_style
    writer.sheet("S1") do |proxy|
      proxy.row([1, 2], cell_styles: [:b, nil], row_style: :r)
    end

    writer.close
  end

  test "Elements::Cell.new style vs style_index" do
    # Cell.new with style:
    c1 = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: 42, style: :bold)
    assert_equal(:bold, c1.style)
    assert_nil(c1.style_index)

    # Cell.new with style_index Integer
    c2 = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: 42, style_index: 5)
    assert_nil(c2.style)
    assert_equal(5, c2.style_index)
  end

  test "singular builders" do
    wb = Xlsxrb.build do |b|
      b.workbook_protection(lock_structure: true)
      b.sheet("S1") do |s|
        s.data_validation("A1:A10") { |v| v.list = %w[X Y] }
        s.sheet_protection(sheet: true)
        s.selection("B2:C5")
        s.sheet_property(:tab_color, "FF0000")
        s.sheet_properties(enable_format_conditions_calculation: true)
        s.print_options(grid_lines: true)
        s.sheet_view(show_grid_lines: false)
      end
    end
    ws = wb.sheet(0)
    assert_equal(true, wb.workbook_protection.lock_structure)
    assert_equal(1, ws.data_validations.size)
    assert_equal(true, ws.sheet_protection.sheet)
    assert_equal "B2:C5", ws.selection.sqref
    assert_equal "FF0000", ws.sheet_properties.tab_color
    assert ws.sheet_properties.enable_format_conditions_calculation
    assert ws.print_options.grid_lines
    assert_equal false, ws.sheet_view.show_grid_lines
  end

  test "Commit 4: page breaks, column spelling, and image accessors" do
    wb = Xlsxrb.build do |b|
      b.sheet("S1") do |s|
        s.row_break(5)
        s.column_break(:C)
      end
    end
    ws = wb.sheet(0)
    assert_equal([{ id: 5, max: 16_383, man: 1 }], ws.row_breaks.map(&:to_h))
    assert_equal([{ id: 2, max: 1_048_575, man: 1 }], ws.column_breaks.map(&:to_h))

    # Elements::Image from_column / to_column
    img = Xlsxrb::Elements::Image.new(filename: "test.png", from_column: 1, to_column: 3)
    assert_equal 1, img.from_column
    assert_equal 3, img.to_column
  end

  test "Commit 4: filter_column coercion, defined_name scope, and DataValidationBuilder#list" do
    wb = Xlsxrb.build do |b|
      b.defined_name("GlobalRate", "0.15", scope: :workbook)
      b.sheet("S1") do |s|
        s.auto_filter("A1:D10")
        s.filter_column("B", %w[Active Pending])
        s.filter_column(:C, ["High"])
        s.defined_name("LocalRef", columns: "B", rows: 2, scope: "S1")
      end
    end
    ws = wb.sheet(0)
    fc = ws.unmapped_data.dig(:facade, :filter_columns) || {}
    assert fc.key?(1)
    assert fc.key?(2)

    # DataValidationBuilder#list and #list=
    builder = Xlsxrb::DataValidationBuilder.new("A1")
    assert_false(builder.respond_to?(:[]))
    assert_false(builder.respond_to?(:[]=))
    builder.list = %w[One Two Three]
    assert_equal %w[One Two Three], builder.list
    res = builder.to_h
    assert_equal %w[One Two Three], res[:list]
    assert_nil res[:in]
  end

  test "Commit 5: ChartBuilder and SeriesBuilder zero-arg readers and attr= writers" do
    cb = Xlsxrb::ChartBuilder.new
    cb.title = "Sales Report"
    assert_equal "Sales Report", cb.title
    assert_equal "Sales Report", cb.title

    cb.type = :bar
    assert_equal :bar, cb.type
    assert_equal :bar, cb.type

    cb.legend(position: :bottom)
    assert_equal({ position: :bottom }, cb.legend)

    sb = Xlsxrb::ChartBuilder::SeriesBuilder.new
    sb.name = "Revenue"
    assert_equal "Revenue", sb.name
    assert_equal "Revenue", sb.name

    sb.values = "Sheet1!B2:B10"
    assert_equal "Sheet1!B2:B10", sb.values
    assert_equal "Sheet1!B2:B10", sb.values

    sb.categories = "Sheet1!A2:A10"
    assert_equal "Sheet1!A2:A10", sb.categories
    assert_equal "Sheet1!A2:A10", sb.categories
  end

  test "Commit 5: StyleBuilder font_size, alignment, protection, and attr= writers" do
    sb = Xlsxrb::StyleBuilder.new("CustomStyle")
    sb.bold = true
    sb.italic = true
    sb.font_name = "Calibri"
    sb.font_size = 14
    sb.font_color = :red
    sb.underline = :double
    sb.strike = true

    assert_equal 14, sb.font_size
    assert_equal 14, sb.font_size

    sb.align_horizontal = :center
    sb.align_vertical = :top
    sb.wrap_text = true
    sb.text_rotation = 45
    sb.indent = 2
    assert_equal({ horizontal: :center, vertical: :top, wrap_text: true, text_rotation: 45, indent: 2 }, sb.alignment)

    # Fluent alignment kwargs
    sb.alignment(horizontal: :right)
    assert_equal(:right, sb.alignment[:horizontal])

    # Protection
    sb.locked = true
    sb.hidden = false
    assert_equal({ locked: true, hidden: false }, sb.protection)

    # Number format
    sb.number_format = "#,##0.00"
    assert_equal "#,##0.00", sb.number_format
  end

  test "Commit 5: ConditionalFormatBuilder dxf_id, font_size, and deprecations" do
    cf = Xlsxrb::ConditionalFormatBuilder.new("A1:A10")
    assert_false(cf.respond_to?(:[]))
    assert_false(cf.respond_to?(:[]=))
    assert_false(cf.respond_to?(:formulas))
    assert_false(cf.respond_to?(:formulas=))

    cf.dxf_id = 3
    assert_equal 3, cf.dxf_id
    assert_equal 3, cf.dxf_id
    assert_nil cf.to_h[:format_id]

    cf.font_size = 12
    assert_equal 12, cf.font_size
    assert_equal 12, cf.font_size
    assert_nil cf.to_h[:size]
  end

  test "Commit 5: WorkbookBuilder#style, StreamWriter#[], and DSL return types" do
    builder_styles = nil
    wb = Xlsxrb.build do |b|
      b.style(:headline, bold: true, font_size: 16)
      builder_styles = b.styles
      b.sheet("Main") do |s|
        ret_af = s.auto_filter("A1:C10")
        assert_nil ret_af
        ret_fc = s.filter_column("A", ["Val"])
        assert_nil ret_fc
        ret_sort = s.sort_state("A1:C10", [{ col_id: 0 }])
        assert_nil ret_sort
        ret_outline = s.outline_properties(summary_below: true)
        assert_nil ret_outline
        ret_split = s.split_pane(x_split: 100, y_split: 200)
        assert_nil ret_split
        ret_com = s.comment("A1", "Test comment")
        assert_nil ret_com
        ret_tc = s.threaded_comment("B2", "Threaded comment")
        assert_nil ret_tc
      end
    end
    assert(builder_styles.key?("headline") || builder_styles.key?(:headline))
    assert_not_nil wb.sheet(0)

    # StreamWriter#[]
    io = StringIO.new
    Xlsxrb.write(io) do |w|
      w.sheet("First") do |s|
        s.row([1, 2, 3])
      end
      proxy = w["First"]
      assert_instance_of Xlsxrb::StreamWriter::WorksheetProxy, proxy
      assert_equal "First", w.current_sheet
    end
  end

  test "sheet_view, sheet_format, sheet_properties, and print_options unification across WorksheetBuilder, StreamWriter, and WorksheetProxy" do
    # 1. WorksheetBuilder
    Xlsxrb.build do |b|
      b.sheet("BuilderSheet") do |s|
        # 0 arguments return Value Objects
        assert_instance_of(Xlsxrb::Elements::SheetView, s.sheet_view)
        assert_instance_of(Xlsxrb::Elements::SheetFormat, s.sheet_format)
        assert_instance_of(Xlsxrb::Elements::SheetProperties, s.sheet_properties)
        assert_instance_of(Xlsxrb::Elements::PrintOptions, s.print_options)

        # Keyword arguments return nil
        ret_sv = s.sheet_view(show_grid_lines: false, zoom_scale: 120)
        assert_nil(ret_sv)
        assert_equal(false, s.sheet_view.show_grid_lines)
        assert_equal(120, s.sheet_view.zoom_scale)

        ret_sf = s.sheet_format(default_row_height: 25.5, default_column_width: 14.0)
        assert_nil(ret_sf)
        assert_equal(25.5, s.sheet_format.default_row_height)
        assert_equal(14.0, s.sheet_format.default_column_width)

        ret_sp = s.sheet_properties(tab_color: "FF0000", filter_mode: true)
        assert_nil(ret_sp)
        assert_equal("FF0000", s.sheet_properties.tab_color)
        assert_equal(true, s.sheet_properties.filter_mode)

        ret_po = s.print_options(grid_lines: true, headings: false)
        assert_nil(ret_po)
        assert_equal(true, s.print_options.grid_lines)
        assert_equal(false, s.print_options.headings)

        # 2 positional arguments return nil
        assert_nil(s.sheet_view(:show_zeros, false))
        assert_equal(false, s.sheet_view.show_zeros)

        assert_nil(s.sheet_format(:outline_level_row, 3))
        assert_equal(3, s.sheet_format.outline_level_row)

        assert_nil(s.sheet_properties(:code_name, "SheetCode"))
        assert_equal("SheetCode", s.sheet_properties.code_name)

        assert_nil(s.print_options(:horizontal_centered, true))
        assert_equal(true, s.print_options.horizontal_centered)

        # 1 positional Hash argument returns nil
        assert_nil(s.sheet_view({ view: "pageLayout" }))
        assert_equal("pageLayout", s.sheet_view.view)

        assert_nil(s.sheet_format({ thick_top: true }))
        assert_equal(true, s.sheet_format.thick_top)

        assert_nil(s.sheet_properties({ fit_to_page: true }))
        assert_equal(true, s.sheet_properties.fit_to_page)

        assert_nil(s.print_options({ vertical_centered: true }))
        assert_equal(true, s.print_options.vertical_centered)

        # 1 positional Value Object argument returns nil
        existing_sv = s.sheet_view
        assert_nil(s.sheet_view(existing_sv))
        existing_sf = s.sheet_format
        assert_nil(s.sheet_format(existing_sf))
        existing_sp = s.sheet_properties
        assert_nil(s.sheet_properties(existing_sp))
        existing_po = s.print_options
        assert_nil(s.print_options(existing_po))

        # Setter methods = accept Hash and Value Object
        s.sheet_view = { zoom_scale: 150 }
        assert_equal(150, s.sheet_view.zoom_scale)
        s.sheet_view = Xlsxrb::Elements::SheetView.new(zoom_scale: 90)
        assert_equal(90, s.sheet_view.zoom_scale)

        s.sheet_format = { default_row_height: 30 }
        assert_equal(30, s.sheet_format.default_row_height)
        s.sheet_format = Xlsxrb::Elements::SheetFormat.new(default_row_height: 22)
        assert_equal(22, s.sheet_format.default_row_height)

        s.sheet_properties = { tab_color: "00FF00" }
        assert_equal("00FF00", s.sheet_properties.tab_color)
        s.sheet_properties = Xlsxrb::Elements::SheetProperties.new(tab_color: "0000FF")
        assert_equal("0000FF", s.sheet_properties.tab_color)

        s.print_options = { grid_lines: false }
        assert_equal(false, s.print_options.grid_lines)
        s.print_options = Xlsxrb::Elements::PrintOptions.new(grid_lines: true)
        assert_equal(true, s.print_options.grid_lines)
      end
    end

    # 2. StreamWriter and WorksheetProxy
    io = StringIO.new
    Xlsxrb.write(io) do |w|
      w.sheet("StreamSheet") do |proxy|
        # 0 arguments via proxy and direct writer
        assert_instance_of(Xlsxrb::Elements::SheetView, proxy.sheet_view)
        assert_instance_of(Xlsxrb::Elements::SheetFormat, proxy.sheet_format)
        assert_instance_of(Xlsxrb::Elements::SheetProperties, proxy.sheet_properties)
        assert_instance_of(Xlsxrb::Elements::PrintOptions, proxy.print_options)

        assert_instance_of(Xlsxrb::Elements::SheetView, w.sheet_view)
        assert_instance_of(Xlsxrb::Elements::SheetFormat, w.sheet_format)
        assert_instance_of(Xlsxrb::Elements::SheetProperties, w.sheet_properties)
        assert_instance_of(Xlsxrb::Elements::PrintOptions, w.print_options)

        # Keyword args return nil
        assert_nil(proxy.sheet_view(show_grid_lines: false, zoom_scale: 125))
        assert_equal(false, proxy.sheet_view.show_grid_lines)
        assert_equal(125, proxy.sheet_view.zoom_scale)

        assert_nil(proxy.sheet_format(default_row_height: 28.0, default_column_width: 16.0))
        assert_equal(28.0, proxy.sheet_format.default_row_height)
        assert_equal(16.0, proxy.sheet_format.default_column_width)

        assert_nil(proxy.sheet_properties(tab_color: "AABBCC"))
        assert_equal("AABBCC", proxy.sheet_properties.tab_color)

        assert_nil(proxy.print_options(headings: true))
        assert_equal(true, proxy.print_options.headings)

        # 2 positional args
        assert_nil(proxy.sheet_view(:tab_selected, true))
        assert_equal(true, proxy.sheet_view.tab_selected)

        assert_nil(proxy.sheet_format(:zero_height, true))
        assert_equal(true, proxy.sheet_format.zero_height)

        assert_nil(proxy.sheet_properties(:code_name, "CodeStream"))
        assert_equal("CodeStream", proxy.sheet_properties.code_name)

        assert_nil(proxy.print_options(:horizontal_centered, false))
        assert_equal(false, proxy.print_options.horizontal_centered)

        # Setters = on proxy and writer
        proxy.sheet_view = { zoom_scale: 110 }
        assert_equal(110, proxy.sheet_view.zoom_scale)
        proxy.sheet_format = { default_row_height: 32 }
        assert_equal(32, proxy.sheet_format.default_row_height)
        proxy.sheet_properties = { tab_color: "112233" }
        assert_equal("112233", proxy.sheet_properties.tab_color)
        proxy.print_options = { grid_lines: true }
        assert_equal(true, proxy.print_options.grid_lines)

        w.sheet_view = { zoom_scale: 100 }
        assert_equal(100, w.sheet_view.zoom_scale)
        w.sheet_format = { default_row_height: 20 }
        assert_equal(20, w.sheet_format.default_row_height)
        w.sheet_properties = { tab_color: "FFFFFF" }
        assert_equal("FFFFFF", w.sheet_properties.tab_color)
        w.print_options = { grid_lines: false }
        assert_equal(false, w.print_options.grid_lines)

        proxy.row([1, 2, 3])
      end
    end
  end

  test "Phase 6: WorkbookBuilder and StreamWriter symmetry and top-level utility facade" do
    # 1. WorkbookBuilder getters and escape_formulas
    wb = Xlsxrb.build(escape_formulas: true) do |b|
      assert_equal(true, b.escape_formulas?)
      assert_equal(true, b.escape_formulas)

      b.core_property(:creator, "Tester")
      assert_equal("Tester", b.core_properties.creator)
      assert_instance_of(Xlsxrb::Elements::CoreProperties, b.core_properties)

      b.app_property(:company, "Acme Corp")
      assert_equal("Acme Corp", b.app_properties.company)
      assert_instance_of(Xlsxrb::Elements::AppProperties, b.app_properties)

      b.custom_property("Department", "Engineering")
      assert_equal("Engineering", b.custom_properties["Department"].value)
      assert_instance_of(Xlsxrb::Elements::CustomProperty, b.custom_properties["Department"])

      b.workbook_protection(lock_structure: true)
      assert_equal(true, b.workbook_protection.lock_structure)
      assert_instance_of(Xlsxrb::Elements::WorkbookProtection, b.workbook_protection)
      b.workbook_protection = { lock_structure: false }
      assert_equal(false, b.workbook_protection.lock_structure)

      b.workbook_properties = { update_links: "never", filter_privacy: true }
      assert_equal(true, b.workbook_properties.filter_privacy)
      assert_instance_of(Xlsxrb::Elements::WorkbookProperties, b.workbook_properties)

      b.defined_name("TAX_RATE", "0.10")
      assert_equal(1, b.defined_names.size)
      assert_instance_of(Xlsxrb::Elements::DefinedName, b.defined_names.first)
      assert_equal("TAX_RATE", b.defined_names.first.name)

      b.style(:bold_num, bold: true, number_format: "#,##0")
      assert_include(b.styles.keys, "bold_num")

      b.sheet("SheetOne") do |s|
        assert_equal(true, s.escape_formulas?)
        s.row(["=1+1", 100])
      end
      assert_equal(["SheetOne"], b.sheet_names)
    end

    # 2. StreamWriter getters and symmetry
    Tempfile.create(["sw_symmetry", ".xlsx"]) do |f|
      Xlsxrb.write(f.path, escape_formulas: true) do |w|
        assert_equal(true, w.escape_formulas?)
        assert_equal(true, w.escape_formulas)

        w.core_property(:creator, "StreamTester")
        assert_equal("StreamTester", w.core_properties.creator)
        assert_instance_of(Xlsxrb::Elements::CoreProperties, w.core_properties)

        w.app_property(:company, "StreamCorp")
        assert_equal("StreamCorp", w.app_properties.company)
        assert_instance_of(Xlsxrb::Elements::AppProperties, w.app_properties)

        w.custom_property("Client", "Beta")
        assert_equal("Beta", w.custom_properties["Client"].value)
        assert_instance_of(Xlsxrb::Elements::CustomProperty, w.custom_properties["Client"])

        w.workbook_protection(lock_structure: true)
        assert_equal(true, w.workbook_protection.lock_structure)
        assert_instance_of(Xlsxrb::Elements::WorkbookProtection, w.workbook_protection)
        w.workbook_protection = { lock_structure: false }
        assert_equal(false, w.workbook_protection.lock_structure)

        w.workbook_properties = { update_links: "never", filter_privacy: true }
        assert_equal(true, w.workbook_properties.filter_privacy)
        assert_instance_of(Xlsxrb::Elements::WorkbookProperties, w.workbook_properties)

        w.defined_name("DISCOUNT", "0.05")
        assert_equal(1, w.defined_names.size)
        assert_instance_of(Xlsxrb::Elements::DefinedName, w.defined_names.first)
        assert_equal("DISCOUNT", w.defined_names.first.name)

        w.style(:stream_style, italic: true)
        assert_include(w.styles.keys, "stream_style")

        w.repeat_rows(0)
        w.repeat_columns("A", "B")

        w.sheet("StreamSheet") do |s|
          assert_equal(true, s.escape_formulas?)
          s.row(["Data", 50])
        end
        assert_equal(["StreamSheet"], w.sheet_names)
      end
    end

    # 3. Top-level Xlsxrb write
    Tempfile.create(["ww_test", ".xlsx"]) do |f|
      Xlsxrb.write(f.path, wb, shared_strings: true, compression_level: 6)
      assert_true(File.size(f.path).positive?)
      loaded = Xlsxrb.load(f.path)
      assert_equal(["SheetOne"], loaded.sheet_names)
    end

    data = Xlsxrb.write(wb, shared_strings: false, compression_level: 0)
    assert_instance_of(String, data)
    assert_true(data.start_with?("PK\x03\x04".b))
  end

  test "Xlsxrb.split_coordinate splits cell reference into column name and 1-based row number" do
    assert_equal(["A", 1], Xlsxrb.split_coordinate("A1"))
    assert_equal(["BC", 100], Xlsxrb.split_coordinate("BC100"))
    assert_equal(["Z", 5], Xlsxrb.split_coordinate("Z5"))
    assert_nil(Xlsxrb.split_coordinate(nil))
    assert_nil(Xlsxrb.split_coordinate("invalid"))
  end

  test "Xlsxrb.expand_formula shifts relative cell references by coordinates or delta" do
    # Simple shift
    assert_equal("A3*2", Xlsxrb.expand_formula("A2*2", from_ref: "B2", to_ref: "B3"))
    assert_equal("B2*2", Xlsxrb.expand_formula("A2*2", from_ref: "A2", to_ref: "B2"))

    # Direct deltas
    assert_equal("A4*2", Xlsxrb.expand_formula("A2*2", row_delta: 2, column_delta: 0))
    assert_equal("C2*2", Xlsxrb.expand_formula("A2*2", row_delta: 0, column_delta: 2))
    assert_equal("C4*2", Xlsxrb.expand_formula("A2*2", row_delta: 2, column_delta: 2))

    # Absolute and mixed references
    assert_equal("$A$1+A3", Xlsxrb.expand_formula("$A$1+A2", from_ref: "B2", to_ref: "B3"))
    assert_equal("A$1+B$1", Xlsxrb.expand_formula("A$1+B$1", from_ref: "A2", to_ref: "A3"))
    assert_equal("$A2+$A3", Xlsxrb.expand_formula("$A1+$A2", from_ref: "B1", to_ref: "B2"))

    # Range references
    assert_equal("SUM(A3:C3)", Xlsxrb.expand_formula("SUM(A2:C2)", from_ref: "D2", to_ref: "D3"))
    assert_equal("AVERAGE(B3:B10)", Xlsxrb.expand_formula("AVERAGE(A3:A10)", from_ref: "A1", to_ref: "B1"))

    # Preserves string literals
    formula = 'SUM(A2:B2) & " Text A1 here " & A2'
    assert_equal('SUM(A3:B3) & " Text A1 here " & A3', Xlsxrb.expand_formula(formula, from_ref: "C2", to_ref: "C3"))

    # Escaped double quotes in string literals
    f_esc = 'A2 & " ""A1"" test "'
    assert_equal('A3 & " ""A1"" test "', Xlsxrb.expand_formula(f_esc, from_ref: "B2", to_ref: "B3"))

    # Sheet prefix
    assert_equal("Sheet1!A3 + 'My Sheet'!B3", Xlsxrb.expand_formula("Sheet1!A2 + 'My Sheet'!B2", from_ref: "C2", to_ref: "C3"))

    # Functions with numbers
    assert_equal("LOG10(A3)", Xlsxrb.expand_formula("LOG10(A2)", from_ref: "B2", to_ref: "B3"))

    # No shift when delta is zero
    assert_equal("A2*2", Xlsxrb.expand_formula("A2*2", from_ref: "B2", to_ref: "B2"))
    assert_equal("A2*2", Xlsxrb.expand_formula("A2*2", row_delta: 0, column_delta: 0))

    # Empty or nil formula
    assert_nil(Xlsxrb.expand_formula(nil, from_ref: "A1", to_ref: "A2"))
    assert_equal("", Xlsxrb.expand_formula("", from_ref: "A1", to_ref: "A2"))

    # Invalid coordinates fallback to unmodified formula
    assert_equal("A2*2", Xlsxrb.expand_formula("A2*2", from_ref: "invalid", to_ref: "B2"))
  end
end
