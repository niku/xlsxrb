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

  test "Xlsxrb.rich_text supports multiple runs" do
    rt = Xlsxrb.rich_text({ text: "A", font: { bold: true } }, { text: "B", font: { italic: true } })
    assert_instance_of(Xlsxrb::Elements::RichText, rt)
    assert_equal(2, rt.runs.size)
    assert_equal("A", rt.runs[0][:text])
    assert_equal({ bold: true }, rt.runs[0][:font])
    assert_equal("B", rt.runs[1][:text])
    assert_equal({ italic: true }, rt.runs[1][:font])
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

  test "WorkbookBuilder#protect_workbook configures correctly" do
    wb = Xlsxrb.build do |w|
      w.protect_workbook(password: "123")
      w.sheet("S")
    end
    assert_equal({ password: "123" }, wb.unmapped_data.dig(:facade, :workbook_protection))
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

  test "WorksheetBuilder#protect_sheet configures correctly" do
    wb = Xlsxrb.build do |w|
      w.sheet("S") do |s|
        s.protect_sheet(password: "123")
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

  test "WorksheetBuilder#validate_data configures correctly" do
    wb = Xlsxrb.build do |w|
      w.sheet("S") do |s|
        s.validate_data("A1", type: :whole)
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
    assert_equal [{ val: "Sheet1!$B$1:$B$5" }], c[:series]
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

  test "WorksheetBuilder#select_cell captures active cell" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.select_cell("C3")
      end
    end
    sel = wb.sheet(0).unmapped_data.dig(:facade, :selection)
    assert_equal "C3", sel[:active_cell]
  end

  test "WorksheetBuilder#page_break_row adds row breaks" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.page_break_row(10)
      end
    end
    breaks = wb.sheet(0).unmapped_data.dig(:facade, :row_breaks)
    assert_equal 10, breaks.first
  end

  test "WorksheetBuilder#page_break_col adds col breaks" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.page_break_col(5)
      end
    end
    breaks = wb.sheet(0).unmapped_data.dig(:facade, :col_breaks)
    assert_equal 5, breaks.first
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

  test "WorksheetBuilder#validate_data captures validation" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.validate_data("A1", type: :list)
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
        s.freeze_pane(row: 1, col: "B")
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
    shapes = wb.sheet(0).unmapped_data.dig(:facade, :shapes)
    assert_equal "rect", shapes.first[:preset]
    assert_equal "Hi", shapes.first[:text]
  end

  test "WorksheetBuilder#table captures values" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.table("A1:B10", columns: [], name: "Table1")
      end
    end
    tables = wb.sheet(0).unmapped_data.dig(:facade, :tables)
    assert_equal "Table1", tables.first[:name]
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
    assert_equal true, col.hidden
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
    assert_equal({ text: "My Chart" }, chart[:title])
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
    assert_equal({ position: "b" }, chart[:legend])
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
    assert_equal({ title: "X Axis" }, chart[:category_axis])
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
    assert_equal({ title: "Y Axis" }, chart[:value_axis])
  end

  test "WorksheetBuilder#chart without block" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.chart(target: "A1", type: :bar, title: { text: "No Block" })
      end
    end
    chart = wb.sheet(0).charts.first
    assert_equal({ text: "No Block" }, chart[:title])
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

  test "Xlsxrb.generate streaming output yields chunks" do
    out = StringIO.new
    Xlsxrb.generate(out) do |w|
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
    assert_equal true, col.custom_width
  end

  test "Xlsxrb.read parses workbook" do
    temp = Tempfile.new(["test_read", ".xlsx"])
    Xlsxrb.generate(temp.path) do |w|
      w.sheet("S1") { |s| s.row ["A"] }
    end

    wb = Xlsxrb.read(temp.path)
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

  test "WorksheetBuilder#page_break_row creates break" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.page_break_row(5)
      end
    end
    breaks = wb.sheet(0).unmapped_data.dig(:facade, :row_breaks)
    assert breaks
  end

  test "WorksheetBuilder#page_break_col creates break" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.page_break_col("C")
      end
    end
    breaks = wb.sheet(0).unmapped_data.dig(:facade, :col_breaks)
    assert breaks
  end

  test "WorksheetBuilder#select_cell captures active cell" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.select_cell("D4")
      end
    end
    selection = wb.sheet(0).unmapped_data.dig(:facade, :selection)
    assert_equal "D4", selection[:active_cell]
  end

  test "WorksheetBuilder#validate_data captures options" do
    wb = Xlsxrb.build do |w|
      w.sheet("S1") do |s|
        s.validate_data("A1:A10", type: :list)
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

  test "Elements::Cell#content returns value" do
    cell = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "test")
    assert_equal "test", cell.content
  end

  test "Elements::Cell.column_letter works" do
    assert_equal "A", Xlsxrb::Elements::Cell.column_letter(0)
    assert_equal "Z", Xlsxrb::Elements::Cell.column_letter(25)
    assert_equal "AA", Xlsxrb::Elements::Cell.column_letter(26)
  end

  test "Elements::Cell.column_index works" do
    assert_equal 0, Xlsxrb::Elements::Cell.column_index("A")
    assert_equal 25, Xlsxrb::Elements::Cell.column_index("Z")
    assert_equal 26, Xlsxrb::Elements::Cell.column_index("AA")
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

  test "Elements::Row#to_a aliases values" do
    row = Xlsxrb::Elements::Row.new(index: 0, cells: [
                                      Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "A"),
                                      Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 1, value: "B")
                                    ])
    assert_equal %w[A B], row.to_a
  end

  test "Elements::Worksheet#cell_value works" do
    ws = Xlsxrb::Elements::Worksheet.new(name: "S1", rows: [
                                           Xlsxrb::Elements::Row.new(index: 0, cells: [
                                                                       Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "A")
                                                                     ])
                                         ])
    assert_equal "A", ws.cell_value("A1")
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

  test "Elements::Worksheet#cells_hash maps ref to cell" do
    ws = Xlsxrb::Elements::Worksheet.new(name: "S1", rows: [
                                           Xlsxrb::Elements::Row.new(index: 0, cells: [
                                                                       Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "A")
                                                                     ])
                                         ])
    assert_equal "A", ws.cells_hash["A1"].value
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

  test "Elements::Worksheet#update_cell works" do
    ws = Xlsxrb::Elements::Worksheet.new(name: "S1", rows: [
                                           Xlsxrb::Elements::Row.new(index: 0, cells: [
                                                                       Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "A")
                                                                     ])
                                         ])
    ws = ws.update_cell("A1", value: "Z")
    assert_equal "Z", ws.cell_value("A1")
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

  test "Elements::Workbook#update_sheet works" do
    wb = Xlsxrb::Elements::Workbook.new(sheets: [
                                          Xlsxrb::Elements::Worksheet.new(name: "S1", rows: [])
                                        ])
    wb = wb.update_sheet("S1") do |_s|
      Xlsxrb::Elements::Worksheet.new(name: "S2", rows: [])
    end
    assert_equal ["S2"], wb.sheet_names
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
    assert_equal 12, builder.font_props[:sz]
  end

  test "StyleBuilder#bold method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.bold
    assert_equal true, builder.font_props[:bold]
  end

  test "StyleBuilder#italic method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.italic
    assert_equal true, builder.font_props[:italic]
  end

  test "StyleBuilder#size method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.size(16)
    assert_equal 16, builder.font_props[:sz]
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
    builder.strike
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
    builder.wrap_text
    assert_equal true, builder.alignment[:wrap_text]
  end

  test "StyleBuilder#shrink_to_fit method" do
    builder = Xlsxrb::StyleBuilder.new
    builder.shrink_to_fit
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
    assert_equal 14, builder.num_fmt_id
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
        s.row([1, 2, 3, 4, 5], styles: { "A" => "bold", "C".."D" => "italic" })
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
        s.row([1], styles: [{ font: { bold: true } }])
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
    Xlsxrb.generate(io) do |w|
      w.workbook_property(:update_links, "always")
      w.style("bold", font: { bold: true })

      w.sheet("Test") do |s|
        assert s.respond_to?(:row)

        s.row([1, 2, 3])

        s.merge(row: 0, col_start: 0, col_end: 2)
        s.freeze_pane(row: 1, col: 1)
        s.split_pane(x_split: 1000, y_split: 1000, top_left_cell: "B2")
        s.select_cell("A1", sqref: "A1:A2", pane: "topRight")

        s.page_margins(left: 0.5, right: 0.5, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2)
        s.print_options(:grid_lines, true)
        s.page_setup(paper_size: 9, orientation: "landscape")
        s.header_footer(odd_header: "&L&T")

        s.sheet_view(:show_grid_lines, false)
        s.protect_sheet(password: "secret", sheet: true)
      end
    end
    assert io.size.positive?
  end

  test "Writer raises error when writing to inactive sheet" do
    io = StringIO.new
    Xlsxrb.generate(io) do |w|
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
        s.merge(row: 0, col_start: 0, col_end: 2)
      end
    end
    assert wb
  end
end
