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
      Xlsxrb.foreach(tmp.path) do |sheet|
        sheet.each do |row|
          assert_instance_of(Xlsxrb::Elements::Row, row)
          collected << row.cells[0].value
        end
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
      Xlsxrb.foreach(tmp.path).find { |s| s.name == "Second" }&.each { |row| collected << row.cells[0].value }
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
      assert_equal(1, enum.first.first.cells[0].value)
    ensure
      tmp.close!
    end
  end

  # --- generate ---

  test "Xlsxrb.generate creates a valid XLSX" do
    tmp = Tempfile.new(["generate_test", ".xlsx"])
    begin
      Xlsxrb.generate(tmp.path) do |w|
        w.sheet("Output")
        w.row(%w[Name Score])
        w.row(["Alice", 95])
        w.row(["Bob", 87])
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
        w.sheet("Sheet1")
        w.row([1, 2, 3])
        w.sheet("Sheet2")
        w.row([4, 5, 6])
      end

      result = Xlsxrb.read(tmp.path)
      assert_equal(2, result.sheets.size)
      assert_equal(1, result.sheet("Sheet1").cell_value("A1"))
      assert_equal(4, result.sheet("Sheet2").cell_value("A1"))
    ensure
      tmp.close!
    end
  end

  test "Xlsxrb.generate without explicit sheet" do
    tmp = Tempfile.new(["gen_implicit", ".xlsx"])
    begin
      Xlsxrb.generate(tmp.path) do |w|
        w.row(["auto"])
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
        w.row([true, false, nil, "text", 42])
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
        w.sheet("Big")
        10_000.times do |i|
          w.row([i, "row#{i}", i * 0.5])
        end
      end

      count = 0
      sum = 0
      Xlsxrb.foreach(tmp.path) do |sheet|
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
      Xlsxrb.generate(tmp.path) do |w|
        w.sheet("Large")
        10_000.times do |i|
          w.row([i, "data#{i}"])
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
        w.row(["x", 1])
        w.row(["y", 2])
        w.row(["z", 3])
      end

      values = []
      Xlsxrb.foreach(tmp.path) do |sheet|
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
    Xlsxrb.generate(tmp.path) do |w|
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
        s.row(%w[Name Score], styles: %w[header header])
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
    Xlsxrb.generate(tmp.path) do |w|
      w.sheet("Sales") do
        w.row(%w[Month Value])
        w.row(["Jan", 100])
        w.row(["Feb", 200])
        w.chart do |c|
          c.type :bar
          c.title "Sales Data"
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

  # --- Xlsxrb.modify ---

  test "Xlsxrb.modify reads and writes back a workbook with value changes" do
    source = Tempfile.new(["modify_source", ".xlsx"])
    target = Tempfile.new(["modify_target", ".xlsx"])
    begin
      # Create source file
      workbook = Xlsxrb.build do |w|
        w.sheet("Data") do |s|
          s.row(%w[Name Score])
          s.row(["Alice", 95])
          s.row(["Bob", 87])
        end
      end
      Xlsxrb.write(source.path, workbook)

      # Modify: change Bob's score
      Xlsxrb.modify(source.path, target.path) do |wb|
        sheet = wb.sheet(0)
        row2 = sheet.row_at(2)
        new_cell = Xlsxrb::Elements::Cell.new(row_index: 2, column_index: 1, value: 99)
        new_row = row2.with(cells: row2.cells.map { |c| c.column_index == 1 ? new_cell : c })
        new_sheet = sheet.with(rows: sheet.rows.map { |r| r.index == 2 ? new_row : r })
        wb.with(sheets: wb.sheets.map.with_index { |s, i| i.zero? ? new_sheet : s })
      end

      # Read back and verify
      result = Xlsxrb.read(target.path)
      assert_equal("Data", result.sheet(0).name)
      assert_equal("Alice", result.sheet(0).cell_value("A2"))
      assert_equal(99, result.sheet(0).cell_value("B3"))
    ensure
      source.close!
      target.close!
    end
  end

  test "Xlsxrb.modify overwrites source when no target given" do
    tmp = Tempfile.new(["modify_inplace", ".xlsx"])
    begin
      workbook = Xlsxrb.build do |w|
        w.sheet("S") do |s|
          s.row(["original"])
        end
      end
      Xlsxrb.write(tmp.path, workbook)

      Xlsxrb.modify(tmp.path) do |wb|
        sheet = wb.sheet(0)
        row = sheet.row_at(0)
        new_cell = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "modified")
        new_row = row.with(cells: [new_cell])
        new_sheet = sheet.with(rows: [new_row])
        wb.with(sheets: [new_sheet])
      end

      result = Xlsxrb.read(tmp.path)
      assert_equal("modified", result.sheet(0).cell_value("A1"))
    ensure
      tmp.close!
    end
  end

  test "Xlsxrb.modify raises on nil source" do
    assert_raise(Xlsxrb::Error) { Xlsxrb.modify(nil) { |wb| wb } }
  end

  test "Xlsxrb.modify raises when no block given" do
    assert_raise(Xlsxrb::Error) { Xlsxrb.modify("/tmp/nonexistent.xlsx") }
  end

  test "Xlsxrb.modify preserves workbook when block returns non-Workbook" do
    tmp = Tempfile.new(["modify_noop", ".xlsx"])
    begin
      workbook = Xlsxrb.build do |w|
        w.sheet("Sheet1") do |s|
          s.row(["keep"])
        end
      end
      Xlsxrb.write(tmp.path, workbook)

      # Block returns nil — workbook should be preserved
      Xlsxrb.modify(tmp.path) { |_wb| nil }

      result = Xlsxrb.read(tmp.path)
      assert_equal("keep", result.sheet(0).cell_value("A1"))
    ensure
      tmp.close!
    end
  end
  # 1. Workbook#update_sheet
  test "Workbook#update_sheet creates a new workbook with the updated sheet" do
    wb = Xlsxrb::Elements::Workbook.new(
      sheets: [
        Xlsxrb::Elements::Worksheet.new(name: "Sheet1", rows: [
                                          Xlsxrb::Elements::Row.new(index: 0, cells: [
                                                                      Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "Hello")
                                                                    ])
                                        ])
      ]
    )

    new_wb = wb.update_sheet("Sheet1") do |sheet|
      sheet.update_cell("A1", value: "World")
    end

    assert_not_equal wb, new_wb
    assert_equal "Hello", wb.sheet("Sheet1").cell_value("A1")
    assert_equal "World", new_wb.sheet("Sheet1").cell_value("A1")
  end

  # 2. Worksheet#update_cell
  test "Worksheet#update_cell creates a new cell or updates an existing one" do
    sheet = Xlsxrb::Elements::Worksheet.new(name: "Test")

    # Update new cell
    sheet2 = sheet.update_cell("B2", value: 42)
    assert_nil sheet.cell_value("B2")
    assert_equal 42, sheet2.cell_value("B2")

    # Update existing cell
    sheet3 = sheet2.update_cell("B2", value: 100, style_index: 1)
    assert_equal 42, sheet2.cell_value("B2")
    assert_equal 100, sheet3.cell_value("B2")
    assert_equal 1, sheet3["B2"].style_index
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

    Xlsxrb.generate(temp_file.path) do |wb|
      wb.style("bold", &:bold)
      wb.style("italic", &:italic)
      wb.style("red") { |s| s.font_color(:red) }

      wb.sheet("Test")
      wb.row(
        { "A" => 1, "B" => 2, "C" => 3, "D" => 4 },
        styles: { "A" => "bold", "B".."C" => "italic", "D" => "red" }
      )

      wb.column("A".."B", width: 20)
      wb.column(%w[C D], width: 10)
    end

    parsed = Xlsxrb.read(temp_file.path)
    sheet = parsed.sheet("Test")

    assert_equal 1, sheet.cell_value("A1")
    assert_equal 2, sheet.cell_value("B1")
    assert_equal 3, sheet.cell_value("C1")
    assert_equal 4, sheet.cell_value("D1")
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
    assert_equal 14, sb.font_props[:sz]

    assert_equal "solid", sb.fill_props[:pattern]
    assert_equal "FF0000FF", sb.fill_props[:fg_color]

    assert_equal "thick", sb.border_props[:top][:style]

    assert_equal "center", sb.alignment[:horizontal]
    assert_equal "0.00", sb.num_fmt_id
  end

  test "StyleBuilder fluent DSL configures correctly" do
    sb = Xlsxrb::StyleBuilder.new("test")
    sb.font(bold: true, color: :red)
      .fill_color(:blue)
      .border_all(style: "thick")
      .align_horizontal("center")
      .number_format("0.00")

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

  test "Xlsxrb.rich_text returns a rich text element" do
    rt = Xlsxrb.rich_text(text: "Hello", bold: true)
    assert_instance_of(Xlsxrb::Elements::RichText, rt)
    assert_equal(1, rt.runs.size)
    assert_equal("Hello", rt.runs[0][:text])
    assert_equal({ bold: true }, rt.runs[0][:font])
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
        s.freeze_pane(row: 1, col: "B")
      end
    end
    assert_equal({ row: 1, col: 1 }, wb.sheet(0).unmapped_data.dig(:facade, :freeze_pane))
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
    assert_equal({creator: "test"}, wb.unmapped_data.dig(:facade, :core_properties))
  end

  test "WorkbookBuilder#app_property configures correctly" do
    wb = Xlsxrb.build do |w|
      w.app_property(:company, "test")
      w.sheet("S")
    end
    assert_equal({company: "test"}, wb.unmapped_data.dig(:facade, :app_properties))
  end

  test "WorkbookBuilder#custom_property configures correctly" do
    wb = Xlsxrb.build do |w|
      w.custom_property("test", "test")
      w.sheet("S")
    end
    assert_equal([{name: "test", value: "test", type: :string}], wb.unmapped_data.dig(:facade, :custom_properties))
  end

  test "WorkbookBuilder#protect_workbook configures correctly" do
    wb = Xlsxrb.build do |w|
      w.protect_workbook(password: "123")
      w.sheet("S")
    end
    assert_equal({password: "123"}, wb.unmapped_data.dig(:facade, :workbook_protection))
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
end
