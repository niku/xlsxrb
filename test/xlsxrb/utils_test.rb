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
end
