# frozen_string_literal: true

require "test_helper"
require "tempfile"

class FacadeTest < Test::Unit::TestCase
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

      result = Xlsxrb.read(tmp.path)

      assert_instance_of(Xlsxrb::Elements::Workbook, result)
      assert_equal(1, result.sheets.size)
      assert_equal("TestSheet", result.sheets[0].name)
      assert_equal(2, result.sheets[0].rows.size)

      sheet = result.sheet(0)
      assert_equal("Hello", sheet.cell_value("A1"))
      assert_equal(42, sheet.cell_value("B1"))
      assert_equal(true, sheet.cell_value("A2"))
      assert_in_delta(3.14, sheet.cell_value("B2"))
    ensure
      tmp.close!
    end
  end

  test "Xlsxrb.write raises on nil target" do
    ws = Xlsxrb::Elements::Worksheet.new(name: "S")
    wb = Xlsxrb::Elements::Workbook.new(sheets: [ws])
    assert_raise(Xlsxrb::Error) { Xlsxrb.write(nil, wb) }
  end

  test "Xlsxrb.write raises on non-workbook" do
    assert_raise(Xlsxrb::Error) { Xlsxrb.write("/tmp/test.xlsx", "not a workbook") }
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
      result = Xlsxrb.read(tmp.path)

      sheet = result.sheet(0)
      assert_equal(0, sheet.cell_value("A1"))
      assert_equal(-99, sheet.cell_value("B1"))
      assert_equal(1_000_000, sheet.cell_value("C1"))
      assert_in_delta(1.5, sheet.cell_value("D1"))
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
      result = Xlsxrb.read(tmp.path)

      assert_equal(2, result.sheets.size)
      assert_equal(%w[First Second], result.sheet_names)
      assert_equal("sheet1", result.sheet("First").cell_value("A1"))
      assert_equal("sheet2", result.sheet("Second").cell_value("A1"))
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
      result = Xlsxrb.read(tmp.path)

      assert_equal(true, result.sheet(0).cell_value("A1"))
      assert_equal(false, result.sheet(0).cell_value("B1"))
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
      result = Xlsxrb.read(tmp.path)

      assert_equal("", result.sheet(0).cell_value("A1"))
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
      result = Xlsxrb.read(tmp.path)

      assert_equal(100, result.sheet(0).rows.size)
      assert_equal(0, result.sheet(0).cell_value("A1"))
      assert_equal(99, result.sheet(0).cell_value("A100"))
      assert_equal("row50", result.sheet(0).cell_value("B51"))
    ensure
      tmp.close!
    end
  end

  # --- foreach ---

  test "Xlsxrb.foreach yields rows one at a time" do
    tmp = Tempfile.new(["foreach_test", ".xlsx"])
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
      Xlsxrb.foreach(tmp.path) do |row|
        assert_instance_of(Xlsxrb::Elements::Row, row)
        collected << row.cells[0].value
      end

      assert_equal([0, 10, 20, 30, 40], collected)
    ensure
      tmp.close!
    end
  end

  test "Xlsxrb.foreach with sheet name" do
    tmp = Tempfile.new(["foreach_sheet", ".xlsx"])
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

      collected = []
      Xlsxrb.foreach(tmp.path, sheet: "Second") { |row| collected << row.cells[0].value }
      assert_equal(["B"], collected)
    ensure
      tmp.close!
    end
  end

  test "Xlsxrb.foreach returns enumerator without block" do
    tmp = Tempfile.new(["foreach_enum", ".xlsx"])
    begin
      ws = Xlsxrb::Elements::Worksheet.new(
        name: "S",
        rows: [Xlsxrb::Elements::Row.new(
          index: 0,
          cells: [Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: 1)]
        )]
      )
      wb = Xlsxrb::Elements::Workbook.new(sheets: [ws])
      Xlsxrb.write(tmp.path, wb)

      enum = Xlsxrb.foreach(tmp.path)
      assert_instance_of(Enumerator, enum)
      assert_equal(1, enum.first.cells[0].value)
    ensure
      tmp.close!
    end
  end

  # --- generate ---

  test "Xlsxrb.generate creates a valid XLSX" do
    tmp = Tempfile.new(["generate_test", ".xlsx"])
    begin
      Xlsxrb.generate(tmp.path) do |w|
        w.add_sheet("Output")
        w.add_row(%w[Name Score])
        w.add_row(["Alice", 95])
        w.add_row(["Bob", 87])
      end

      result = Xlsxrb.read(tmp.path)
      assert_equal(1, result.sheets.size)
      assert_equal("Output", result.sheet(0).name)
      assert_equal(3, result.sheet(0).rows.size)
      assert_equal("Name", result.sheet(0).cell_value("A1"))
      assert_equal(95, result.sheet(0).cell_value("B2"))
    ensure
      tmp.close!
    end
  end

  test "Xlsxrb.generate with multiple sheets" do
    tmp = Tempfile.new(["gen_multi", ".xlsx"])
    begin
      Xlsxrb.generate(tmp.path) do |w|
        w.add_sheet("Sheet1")
        w.add_row([1, 2, 3])
        w.add_sheet("Sheet2")
        w.add_row([4, 5, 6])
      end

      result = Xlsxrb.read(tmp.path)
      assert_equal(2, result.sheets.size)
      assert_equal(1, result.sheet("Sheet1").cell_value("A1"))
      assert_equal(4, result.sheet("Sheet2").cell_value("A1"))
    ensure
      tmp.close!
    end
  end

  test "Xlsxrb.generate without explicit add_sheet" do
    tmp = Tempfile.new(["gen_implicit", ".xlsx"])
    begin
      Xlsxrb.generate(tmp.path) do |w|
        w.add_row(["auto"])
      end

      result = Xlsxrb.read(tmp.path)
      assert_equal("Sheet1", result.sheet(0).name)
      assert_equal("auto", result.sheet(0).cell_value("A1"))
    ensure
      tmp.close!
    end
  end

  test "Xlsxrb.generate raises on nil target" do
    assert_raise(Xlsxrb::Error) { Xlsxrb.generate(nil) { |_w| } } # rubocop:disable Lint/EmptyBlock
  end

  test "Xlsxrb.generate raises without block" do
    assert_raise(Xlsxrb::Error) { Xlsxrb.generate("/tmp/test.xlsx") }
  end

  test "Xlsxrb.generate with booleans and nil" do
    tmp = Tempfile.new(["gen_types", ".xlsx"])
    begin
      Xlsxrb.generate(tmp.path) do |w|
        w.add_row([true, false, nil, "text", 42])
      end

      result = Xlsxrb.read(tmp.path)
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

  test "foreach processes large file without excessive memory" do
    tmp = Tempfile.new(["large_foreach", ".xlsx"])
    begin
      Xlsxrb.generate(tmp.path) do |w|
        w.add_sheet("Big")
        10_000.times do |i|
          w.add_row([i, "row#{i}", i * 0.5])
        end
      end

      count = 0
      sum = 0
      Xlsxrb.foreach(tmp.path) do |row|
        count += 1
        sum += row.cells[0].value.to_i
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
      Xlsxrb.generate(tmp.path) do |w|
        w.add_sheet("Large")
        10_000.times do |i|
          w.add_row([i, "data#{i}"])
        end
      end

      result = Xlsxrb.read(tmp.path)
      assert_equal(10_000, result.sheet(0).rows.size)
      assert_equal(0, result.sheet(0).cell_value("A1"))
      assert_equal(9999, result.sheet(0).cell_value("A10000"))
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

      wb2 = Xlsxrb.read(tmp1.path)
      Xlsxrb.write(tmp2.path, wb2)

      wb3 = Xlsxrb.read(tmp2.path)

      assert_equal("A", wb3.sheet(0).cell_value("A1"))
      assert_equal(1, wb3.sheet(0).cell_value("B1"))
      assert_equal("B", wb3.sheet(0).cell_value("A2"))
      assert_equal(2, wb3.sheet(0).cell_value("B2"))
    ensure
      tmp1.close!
      tmp2.close!
    end
  end

  test "generate then foreach round-trip" do
    tmp = Tempfile.new(["gen_foreach_rt", ".xlsx"])
    begin
      Xlsxrb.generate(tmp.path) do |w|
        w.add_row(["x", 1])
        w.add_row(["y", 2])
        w.add_row(["z", 3])
      end

      values = []
      Xlsxrb.foreach(tmp.path) do |row|
        values << row.values
      end

      assert_equal([["x", 1], ["y", 2], ["z", 3]], values)
    ensure
      tmp.close!
    end
  end

  test "add_chart works in Streaming generate API" do
    tmp = Tempfile.new(["facade_chart_stream", ".xlsx"])
    Xlsxrb.generate(tmp.path) do |w|
      w.add_sheet("Sales") do |s|
        s.add_row(["Month", "Value"])
        s.add_row(["Jan", 100])
        s.add_row(["Feb", 200])
        w.add_chart(type: :bar, title: "Sales Data", series: [{ cat_ref: "Sales!$A$2:$A$3", val_ref: "Sales!$B$2:$B$3" }])
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

  test "add_chart works in In-Memory build API" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("Sales") do |s|
        s.add_row(["Month", "Value"])
        s.add_row(["Jan", 100])
        s.add_row(["Feb", 200])
        s.add_chart(type: :pie, title: "Sales Pie", series: [{ cat_ref: "Sales!$A$2:$A$3", val_ref: "Sales!$B$2:$B$3" }])
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
end
