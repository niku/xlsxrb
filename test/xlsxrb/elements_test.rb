# frozen_string_literal: true

require "test_helper"

class ElementsTest < Test::Unit::TestCase
  # --- Cell ---

  test "cell creates a valid cell" do
    cell = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "hello")
    assert(cell.valid?)
    assert_equal("hello", cell.value)
    assert_equal(0, cell.row_index)
    assert_equal(0, cell.column_index)
    assert_equal("A1", cell.ref)
  end

  test "cell with negative row_index is invalid" do
    cell = Xlsxrb::Elements::Cell.new(row_index: -1, column_index: 0, value: "x")
    assert_include(cell.errors, "row_index must be >= 0")
  end

  test "cell with negative column_index is invalid" do
    cell = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: -1, value: "x")
    assert_include(cell.errors, "column_index must be >= 0")
  end

  test "cell with too large row_index is invalid" do
    cell = Xlsxrb::Elements::Cell.new(row_index: 1_048_576, column_index: 0)
    assert_include(cell.errors, "row_index must be < 1048576")
  end

  test "cell with too large column_index is invalid" do
    cell = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 16_384)
    assert_include(cell.errors, "column_index must be < 16384")
  end

  test "cell column_letter converts index to Excel column" do
    assert_equal("A", Xlsxrb::Elements::Cell.column_letter(0))
    assert_equal("Z", Xlsxrb::Elements::Cell.column_letter(25))
    assert_equal("AA", Xlsxrb::Elements::Cell.column_letter(26))
    assert_equal("AZ", Xlsxrb::Elements::Cell.column_letter(51))
    assert_equal("XFD", Xlsxrb::Elements::Cell.column_letter(16_383))
  end

  test "cell parse_ref converts A1-style to indices" do
    assert_equal([0, 0], Xlsxrb::Elements::Cell.parse_ref("A1"))
    assert_equal([9, 1], Xlsxrb::Elements::Cell.parse_ref("B10"))
    assert_equal([0, 26], Xlsxrb::Elements::Cell.parse_ref("AA1"))
  end

  test "cell ref round-trips" do
    cell = Xlsxrb::Elements::Cell.new(row_index: 5, column_index: 27, value: 42)
    parsed = Xlsxrb::Elements::Cell.parse_ref(cell.ref)
    assert_equal([5, 27], parsed)
  end

  test "cell unmapped_data defaults to empty hash" do
    cell = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0)
    assert_equal({}, cell.unmapped_data)
  end

  test "cell supports various value types" do
    [42, 3.14, "text", true, false, nil].each do |val|
      cell = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: val)
      assert(cell.valid?, "Expected valid cell for value \#{val.inspect}")
    end
  end

  test "cell with unsupported value type is invalid" do
    cell = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: [1, 2, 3])
    assert(cell.errors.any? { |e| e.include?("unsupported value type") })
  end

  test "cell with formula" do
    cell = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: 10, formula: "SUM(A2:A5)")
    assert_equal("SUM(A2:A5)", cell.formula)
    assert(cell.valid?)
  end

  # --- Row ---

  test "row creates a valid row" do
    cells = [Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "a")]
    row = Xlsxrb::Elements::Row.new(index: 0, cells: cells)
    assert(row.valid?)
    assert_equal(0, row.index)
    assert_equal(1, row.cells.size)
  end

  test "row with negative index is invalid" do
    Xlsxrb::Elements::Row.new(index: -1)
  end

  test "row cell_at returns cell by column index" do
    c1 = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "A")
    c2 = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 2, value: "C")
    row = Xlsxrb::Elements::Row.new(index: 0, cells: [c1, c2])

    assert_equal("A", row.cell_at(0).value)
    assert_equal("C", row.cell_at(2).value)
    assert_nil(row.cell_at(1))
  end

  test "row values returns array with nils for gaps" do
    c1 = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: 1)
    c2 = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 2, value: 3)
    row = Xlsxrb::Elements::Row.new(index: 0, cells: [c1, c2])

    assert_equal([1, nil, 3], row.values)
  end

  test "row attributes" do
    row = Xlsxrb::Elements::Row.new(index: 5, height: 25.0, hidden: true)
    assert_in_delta(25.0, row.height)
    assert(row.hidden)
  end

  # --- Column ---

  test "column creates a valid column" do
    col = Xlsxrb::Elements::Column.new(index: 0, width: 15.5)
    assert(col.valid?)
    assert_equal(0, col.index)
    assert_in_delta(15.5, col.width)
  end

  test "column with negative index is invalid" do
    Xlsxrb::Elements::Column.new(index: -1)
  end

  test "column with too large index is invalid" do
    Xlsxrb::Elements::Column.new(index: 16_384)
  end

  # --- Worksheet ---

  test "worksheet creates a valid worksheet" do
    c = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "test")
    r = Xlsxrb::Elements::Row.new(index: 0, cells: [c])
    ws = Xlsxrb::Elements::Worksheet.new(name: "Sheet1", rows: [r])
    assert(ws.valid?)
    assert_equal("Sheet1", ws.name)
  end

  test "worksheet with empty name is invalid" do
    Xlsxrb::Elements::Worksheet.new(name: "")
  end

  test "worksheet with duplicate row indices is invalid" do
    r1 = Xlsxrb::Elements::Row.new(index: 0)
    r2 = Xlsxrb::Elements::Row.new(index: 0)
    ws = Xlsxrb::Elements::Worksheet.new(name: "Sheet1", rows: [r1, r2])
    assert_include(ws.errors, "duplicate row indices")
  end

  test "worksheet row_at returns row by index" do
    r0 = Xlsxrb::Elements::Row.new(index: 0)
    r5 = Xlsxrb::Elements::Row.new(index: 5)
    ws = Xlsxrb::Elements::Worksheet.new(name: "Sheet1", rows: [r0, r5])

    assert_equal(0, ws.row_at(0).index)
    assert_equal(5, ws.row_at(5).index)
    assert_nil(ws.row_at(1))
  end

  test "worksheet cell_value with A1 reference" do
    c = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 1, value: 42)
    r = Xlsxrb::Elements::Row.new(index: 0, cells: [c])
    ws = Xlsxrb::Elements::Worksheet.new(name: "S", rows: [r])

    assert_equal(42, ws.cell_value("B1"))
    assert_nil(ws.cell_value("A1"))
  end

  # --- Workbook ---

  test "workbook creates a valid workbook" do
    ws = Xlsxrb::Elements::Worksheet.new(name: "Sheet1")
    wb = Xlsxrb::Elements::Workbook.new(sheets: [ws])
    assert(wb.valid?)
  end

  test "workbook with no sheets is invalid" do
    wb = Xlsxrb::Elements::Workbook.new(sheets: [])
    assert_include(wb.errors, "workbook must have at least one sheet")
  end

  test "workbook with duplicate sheet names is invalid" do
    ws1 = Xlsxrb::Elements::Worksheet.new(name: "Sheet1")
    ws2 = Xlsxrb::Elements::Worksheet.new(name: "Sheet1")
    wb = Xlsxrb::Elements::Workbook.new(sheets: [ws1, ws2])
    assert_include(wb.errors, "duplicate sheet names")
  end

  test "workbook sheet by index" do
    ws = Xlsxrb::Elements::Worksheet.new(name: "Data")
    wb = Xlsxrb::Elements::Workbook.new(sheets: [ws])
    assert_equal("Data", wb.sheet(0).name)
  end

  test "workbook sheet by name" do
    ws1 = Xlsxrb::Elements::Worksheet.new(name: "Summary")
    ws2 = Xlsxrb::Elements::Worksheet.new(name: "Details")
    wb = Xlsxrb::Elements::Workbook.new(sheets: [ws1, ws2])
    assert_equal("Details", wb.sheet("Details").name)
  end

  test "workbook sheet_names returns all sheet names" do
    ws1 = Xlsxrb::Elements::Worksheet.new(name: "A")
    ws2 = Xlsxrb::Elements::Worksheet.new(name: "B")
    wb = Xlsxrb::Elements::Workbook.new(sheets: [ws1, ws2])
    assert_equal(%w[A B], wb.sheet_names)
  end
end
