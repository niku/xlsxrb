# frozen_string_literal: true

require "test_helper"

class UtilsTest < Test::Unit::TestCase
  cover Xlsxrb::Utils

  test "ref_to_row_col converts cell references to 0-based row and column coordinates" do
    assert_equal([0, 0], Xlsxrb::Utils.ref_to_row_col("A1"))
    assert_equal([0, 1], Xlsxrb::Utils.ref_to_row_col("B1"))
    assert_equal([9, 2], Xlsxrb::Utils.ref_to_row_col("C10"))
    assert_equal([99, 54], Xlsxrb::Utils.ref_to_row_col("BC100"))
    assert_equal([0, 26], Xlsxrb::Utils.ref_to_row_col("AA1"))
    assert_nil(Xlsxrb::Utils.ref_to_row_col(nil))
    assert_nil(Xlsxrb::Utils.ref_to_row_col("invalid"))
  end

  test "row_col_to_ref converts 0-based row and col coordinates to A1-style reference" do
    assert_equal("A1", Xlsxrb::Utils.row_col_to_ref(0, 0))
    assert_equal("B1", Xlsxrb::Utils.row_col_to_ref(0, 1))
    assert_equal("C10", Xlsxrb::Utils.row_col_to_ref(9, 2))
    assert_equal("BC100", Xlsxrb::Utils.row_col_to_ref(99, 54))
    assert_equal("AA1", Xlsxrb::Utils.row_col_to_ref(0, 26))
  end

  test "parse_range converts cell range references to 0-based coordinate pairs" do
    assert_equal([[0, 0], [9, 2]], Xlsxrb::Utils.parse_range("A1:C10"))
    assert_equal([[0, 0], [9, 2]], Xlsxrb.parse_range("A1:C10"))
    assert_equal([[4, 1], [4, 1]], Xlsxrb::Utils.parse_range("B5"))
    assert_equal([[0, 0], [9, 2]], Xlsxrb::Utils.parse_range("C10:A1"))
    assert_nil(Xlsxrb::Utils.parse_range(nil))
    assert_nil(Xlsxrb::Utils.parse_range(""))
    assert_nil(Xlsxrb::Utils.parse_range("invalid"))
    assert_nil(Xlsxrb::Utils.parse_range("invalid:range"))
  end

  test "range_to_ref converts 0-based coordinates to A1-style range reference" do
    assert_equal("A1:C10", Xlsxrb::Utils.range_to_ref(0, 0, 9, 2))
    assert_equal("A1:C10", Xlsxrb.range_to_ref(0, 0, 9, 2))
    assert_equal("B5", Xlsxrb::Utils.range_to_ref(4, 1, 4, 1))
  end

  test "col_name_to_index converts column names to 0-based column index" do
    assert_equal(0, Xlsxrb::Utils.col_name_to_index("A"))
    assert_equal(1, Xlsxrb::Utils.col_name_to_index("B"))
    assert_equal(25, Xlsxrb::Utils.col_name_to_index("Z"))
    assert_equal(26, Xlsxrb::Utils.col_name_to_index("AA"))
    assert_equal(54, Xlsxrb::Utils.col_name_to_index("BC"))
    assert_equal(0, Xlsxrb::Utils.col_name_to_index(:A))
    assert_equal(5, Xlsxrb::Utils.col_name_to_index(5))
  end

  test "col_index_to_name converts 0-based column index to letter" do
    assert_equal("A", Xlsxrb::Utils.col_index_to_name(0))
    assert_equal("B", Xlsxrb::Utils.col_index_to_name(1))
    assert_equal("Z", Xlsxrb::Utils.col_index_to_name(25))
    assert_equal("AA", Xlsxrb::Utils.col_index_to_name(26))
    assert_equal("BC", Xlsxrb::Utils.col_index_to_name(54))
  end

  test "split_coordinate splits cell reference into column name and 1-based row number" do
    assert_equal(["A", 1], Xlsxrb::Utils.split_coordinate("A1"))
    assert_equal(["BC", 100], Xlsxrb::Utils.split_coordinate("BC100"))
    assert_equal(["Z", 5], Xlsxrb::Utils.split_coordinate("Z5"))
    assert_nil(Xlsxrb::Utils.split_coordinate(nil))
    assert_nil(Xlsxrb::Utils.split_coordinate("invalid"))
  end

  test "date_to_serial and serial_to_date convert dates across 1900 and 1904 systems" do
    d = Date.new(2026, 1, 1)
    s1900 = Xlsxrb::Utils.date_to_serial(d)
    assert_equal(46_023, s1900)
    assert_equal(d, Xlsxrb::Utils.serial_to_date(s1900))

    s1904 = Xlsxrb::Utils.date_to_serial(d, date1904: true)
    assert_equal(44_561, s1904)
    assert_equal(d, Xlsxrb::Utils.serial_to_date(s1904, date1904: true))

    # date_to_serial with Time instance
    t = Time.utc(2026, 1, 1, 12, 0, 0)
    assert_in_delta(46_023.5, Xlsxrb::Utils.date_to_serial(t), 1e-6)
  end

  test "datetime_to_serial and serial_to_time convert times across 1900 and 1904 systems" do
    t = Time.utc(2026, 1, 1, 12, 0, 0)
    s1900 = Xlsxrb::Utils.datetime_to_serial(t)
    assert_in_delta(46_023.5, s1900, 1e-6)
    assert_equal(t, Xlsxrb::Utils.serial_to_time(s1900))
    assert_equal(t, Xlsxrb::Utils.serial_to_datetime(s1900))

    s1904 = Xlsxrb::Utils.datetime_to_serial(t, date1904: true)
    assert_in_delta(44_561.5, s1904, 1e-6)
    assert_equal(t, Xlsxrb::Utils.serial_to_time(s1904, date1904: true))

    # datetime_to_serial with Date instance
    d = Date.new(2026, 1, 1)
    assert_in_delta(46_023.0, Xlsxrb::Utils.datetime_to_serial(d), 1e-6)
  end

  test "valid_excel_date? delegates to Ooxml::Utils" do
    assert_true(Xlsxrb::Utils.valid_excel_date?(46_023))
    assert_false(Xlsxrb::Utils.valid_excel_date?(-1))
    assert_true(Xlsxrb::Utils.valid_excel_date?(44_561, date1904: true))
    assert_false(Xlsxrb::Utils.valid_excel_date?(3_000_000, date1904: true))
  end

  test "expand_formula shifts relative cell references by coordinates or delta" do
    # Simple shift
    assert_equal("A3*2", Xlsxrb::Utils.expand_formula("A2*2", from_ref: "B2", to_ref: "B3"))
    assert_equal("B2*2", Xlsxrb::Utils.expand_formula("A2*2", from_ref: "A2", to_ref: "B2"))

    # Direct deltas
    assert_equal("A4*2", Xlsxrb::Utils.expand_formula("A2*2", row_delta: 2, col_delta: 0))
    assert_equal("C2*2", Xlsxrb::Utils.expand_formula("A2*2", row_delta: 0, col_delta: 2))
    assert_equal("C4*2", Xlsxrb::Utils.expand_formula("A2*2", row_delta: 2, col_delta: 2))

    # Absolute and mixed references
    assert_equal("$A$1+A3", Xlsxrb::Utils.expand_formula("$A$1+A2", from_ref: "B2", to_ref: "B3"))
    assert_equal("A$1+B$1", Xlsxrb::Utils.expand_formula("A$1+B$1", from_ref: "A2", to_ref: "A3"))
    assert_equal("$A2+$A3", Xlsxrb::Utils.expand_formula("$A1+$A2", from_ref: "B1", to_ref: "B2"))

    # Range references
    assert_equal("SUM(A3:C3)", Xlsxrb::Utils.expand_formula("SUM(A2:C2)", from_ref: "D2", to_ref: "D3"))
    assert_equal("AVERAGE(B3:B10)", Xlsxrb::Utils.expand_formula("AVERAGE(A3:A10)", from_ref: "A1", to_ref: "B1"))

    # Preserves string literals
    formula = 'SUM(A2:B2) & " Text A1 here " & A2'
    assert_equal('SUM(A3:B3) & " Text A1 here " & A3', Xlsxrb::Utils.expand_formula(formula, from_ref: "C2", to_ref: "C3"))

    # Escaped double quotes in string literals
    f_esc = 'A2 & " ""A1"" test "'
    assert_equal('A3 & " ""A1"" test "', Xlsxrb::Utils.expand_formula(f_esc, from_ref: "B2", to_ref: "B3"))

    # Sheet prefix
    assert_equal("Sheet1!A3 + 'My Sheet'!B3", Xlsxrb::Utils.expand_formula("Sheet1!A2 + 'My Sheet'!B2", from_ref: "C2", to_ref: "C3"))

    # Functions with numbers
    assert_equal("LOG10(A3)", Xlsxrb::Utils.expand_formula("LOG10(A2)", from_ref: "B2", to_ref: "B3"))

    # No shift when delta is zero
    assert_equal("A2*2", Xlsxrb::Utils.expand_formula("A2*2", from_ref: "B2", to_ref: "B2"))
    assert_equal("A2*2", Xlsxrb::Utils.expand_formula("A2*2", row_delta: 0, col_delta: 0))

    # Empty or nil formula
    assert_nil(Xlsxrb::Utils.expand_formula(nil, from_ref: "A1", to_ref: "A2"))
    assert_equal("", Xlsxrb::Utils.expand_formula("", from_ref: "A1", to_ref: "A2"))

    # Invalid coordinates fallback to unmodified formula
    assert_equal("A2*2", Xlsxrb::Utils.expand_formula("A2*2", from_ref: "invalid", to_ref: "B2"))
  end
end
