# frozen_string_literal: true

require "test_helper"

class DslHelpersTest < Test::Unit::TestCase
  cover Xlsxrb::DslHelpers

  # --- normalize_merge_range ---

  test "normalize_merge_range with string range" do
    assert_equal("A1:C3", Xlsxrb::DslHelpers.normalize_merge_range("A1:C3"))
    assert_equal("A1", Xlsxrb::DslHelpers.normalize_merge_range("A1"))

    # Invalid range in strict mode
    assert_raises(ArgumentError) { Xlsxrb::DslHelpers.normalize_merge_range("invalid_range") }
    assert_raises(ArgumentError) { Xlsxrb::DslHelpers.normalize_merge_range("123:456") }

    # Lenient mode bypasses format check
    assert_equal("custom_range", Xlsxrb::DslHelpers.normalize_merge_range("custom_range", strict_excel_mode: false))
  end

  # --- normalize_protection_options ---

  test "normalize_protection_options hashes plain text password" do
    opts = { select_locked_cells: true, password: "mypassword" }
    normalized = Xlsxrb::DslHelpers.normalize_protection_options(opts)

    assert_nil(normalized[:password])
    assert_equal("SHA-512", normalized[:algorithm_name])
    assert_not_nil(normalized[:hash_value])
    assert_not_nil(normalized[:salt_value])
    assert_equal(100_000, normalized[:spin_count])
    assert_equal(true, normalized[:select_locked_cells])

    # Original opts must not be mutated
    assert_equal("mypassword", opts[:password])
    refute(opts.key?(:hash_value))

    # Legacy 4-hex digit password is not hashed
    legacy_opts = { password: "04B2" }
    assert_equal({ password: "04B2" }, Xlsxrb::DslHelpers.normalize_protection_options(legacy_opts))

    # Already hashed options are preserved
    pre_hashed = { password: "secret", hash_value: "already_hashed" }
    assert_equal(pre_hashed, Xlsxrb::DslHelpers.normalize_protection_options(pre_hashed))
  end

  # --- normalize_column_indices ---

  test "normalize_column_indices handles single, array, and range arguments" do
    # Single value
    assert_equal([0], Xlsxrb::DslHelpers.normalize_column_indices(0))
    assert_equal([0], Xlsxrb::DslHelpers.normalize_column_indices("A"))
    assert_equal([1], Xlsxrb::DslHelpers.normalize_column_indices(:B))

    # Range
    assert_equal([1, 2, 3], Xlsxrb::DslHelpers.normalize_column_indices(1..3))
    assert_equal([0, 1, 2], Xlsxrb::DslHelpers.normalize_column_indices("A".."C"))

    # Array
    assert_equal([0, 2, 4], Xlsxrb::DslHelpers.normalize_column_indices([0, "C", :E]))
  end

  # --- normalize_page_margins ---

  test "normalize_page_margins compacts nil values" do
    margins = Xlsxrb::DslHelpers.normalize_page_margins(top: 0.75, bottom: 0.75)
    assert_equal({ top: 0.75, bottom: 0.75 }, margins)

    all_margins = Xlsxrb::DslHelpers.normalize_page_margins(
      left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3
    )
    assert_equal(6, all_margins.size)
  end

  # --- normalize_table_options ---

  test "normalize_table_options builds canonical hash" do
    tbl = Xlsxrb::DslHelpers.normalize_table_options(
      "A1:C10",
      columns: %w[ID Name Price],
      name: "Products",
      display_name: "ProductsTable",
      style: "TableStyleMedium2",
      show_filter: true
    )
    assert_equal("A1:C10", tbl[:ref])
    assert_equal(%w[ID Name Price], tbl[:columns])
    assert_equal("Products", tbl[:name])
    assert_equal("ProductsTable", tbl[:display_name])
    assert_equal("TableStyleMedium2", tbl[:style])
    assert_equal(true, tbl[:show_filter])

    # Minimal table options without optional keys
    min_tbl = Xlsxrb::DslHelpers.normalize_table_options("B2:D5", columns: %w[X Y])
    assert_equal("B2:D5", min_tbl[:ref])
    assert_equal(%w[X Y], min_tbl[:columns])
    refute(min_tbl.key?(:name))
    refute(min_tbl.key?(:display_name))
    refute(min_tbl.key?(:style))

    # Symbol columns and totals_row normalization
    sym_tbl = Xlsxrb::DslHelpers.normalize_table_options(
      "A1:C10",
      columns: [:id, { name: :total, totals_row_function: :sum }, "status"],
      totals_row: true,
      auto_filter: false
    )
    assert_equal("A1:C10", sym_tbl[:ref])
    assert_equal("id", sym_tbl[:columns][0])
    assert_equal({ name: "total", totals_row_function: "sum" }, sym_tbl[:columns][1])
    assert_equal("status", sym_tbl[:columns][2])
    assert_equal(1, sym_tbl[:totals_row_count])
    assert_equal(false, sym_tbl[:auto_filter])

    # totals_row: false sets totals_row_count: 0
    tr_false = Xlsxrb::DslHelpers.normalize_table_options("A1:B2", columns: %w[A B], totals_row: false)
    assert_equal(0, tr_false[:totals_row_count])

    # totals_row: true / false
    trs_true = Xlsxrb::DslHelpers.normalize_table_options("A1:B2", columns: %w[A B], totals_row: true)
    assert_equal(1, trs_true[:totals_row_count])

    # totals_row_count integer
    tr_cnt = Xlsxrb::DslHelpers.normalize_table_options("A1:B2", columns: %w[A B], totals_row_count: 2)
    assert_equal(2, tr_cnt[:totals_row_count])

    # Single column not in array
    single_col_tbl = Xlsxrb::DslHelpers.normalize_table_options("A1:A2", columns: "OnlyCol")
    assert_equal(["OnlyCol"], single_col_tbl[:columns])

    # totals_row_function fallback for unknown symbol or string
    unknown_fn_tbl = Xlsxrb::DslHelpers.normalize_table_options(
      "A1:A2",
      columns: [{ name: "C", totals_row_function: "unknownFn" }, { name: "D", totals_row_function: :custom_unknown }]
    )
    assert_equal("unknownFn", unknown_fn_tbl[:columns][0][:totals_row_function])
    assert_equal("custom_unknown", unknown_fn_tbl[:columns][1][:totals_row_function])

    # Column function mapping for all known symbols
    Xlsxrb::DslHelpers::TOTALS_ROW_FUNCTIONS.each do |sym, expected_str|
      col_tbl = Xlsxrb::DslHelpers.normalize_table_options("A1:A5", columns: [{ name: "Col", totals_row_function: sym }])
      assert_equal(expected_str, col_tbl[:columns].first[:totals_row_function])
    end

    # Calculated column formulas and totals row formulas aliases
    calc_tbl1 = Xlsxrb::DslHelpers.normalize_table_options(
      "A1:B5",
      columns: [
        { name: "C1", formula: "[@Qty]*[@Price]" },
        { name: "C2", calculated_column_formula: "[@Qty]*0.1" },
        { name: "C3", calculated_column_formula: "[@Qty]+1" }
      ]
    )
    assert_equal("[@Qty]*[@Price]", calc_tbl1[:columns][0][:calculated_column_formula])
    assert_equal("[@Qty]*0.1", calc_tbl1[:columns][1][:calculated_column_formula])
    assert_equal("[@Qty]+1", calc_tbl1[:columns][2][:calculated_column_formula])

    tot_tbl1 = Xlsxrb::DslHelpers.normalize_table_options(
      "A1:B5",
      columns: [
        { name: "T1", totals_row_formula: "SUM([T1])" },
        { name: "T2", totals_row_formula: "SUBTOTAL(109,[T2])", totals_row_function: :sum }
      ]
    )
    assert_equal("SUM([T1])", tot_tbl1[:columns][0][:totals_row_formula])
    assert_equal("custom", tot_tbl1[:columns][0][:totals_row_function])
    assert_equal("SUBTOTAL(109,[T2])", tot_tbl1[:columns][1][:totals_row_formula])
    assert_equal("sum", tot_tbl1[:columns][1][:totals_row_function])
  end

  # --- absolute_range ---

  test "absolute_range converts range and single cell to absolute references" do
    assert_equal("$A$1:$B$2", Xlsxrb::DslHelpers.absolute_range("A1:B2"))
    assert_equal("$A$1", Xlsxrb::DslHelpers.absolute_range("A1"))
    assert_equal("$AA$10:$ZZ$99", Xlsxrb::DslHelpers.absolute_range("AA10:ZZ99"))

    # Idempotent on already absolute references
    assert_equal("$A$1:$B$2", Xlsxrb::DslHelpers.absolute_range("$A$1:$B$2"))
    assert_equal("$A$1", Xlsxrb::DslHelpers.absolute_range("$A$1"))
    assert_equal("$A$1", Xlsxrb::DslHelpers.absolute_range("$A1"))
    assert_equal("$A$1", Xlsxrb::DslHelpers.absolute_range("A$1"))
  end

  # --- normalize_row_values ---

  test "normalize_row_values returns array as-is without allocations" do
    arr = [1, "two", 3.0]
    assert_same(arr, Xlsxrb::DslHelpers.normalize_row_values(arr))
    assert_nil(Xlsxrb::DslHelpers.normalize_row_values(nil))
  end

  test "normalize_row_values converts column-keyed hash to sparse array" do
    assert_equal([], Xlsxrb::DslHelpers.normalize_row_values({}))
    assert_equal([10, nil, 30], Xlsxrb::DslHelpers.normalize_row_values({ "A" => 10, "C" => 30 }))
    assert_equal(["x", nil, "y"], Xlsxrb::DslHelpers.normalize_row_values({ 0 => "x", 2 => "y" }))
    assert_equal([42], Xlsxrb::DslHelpers.normalize_row_values({ A: 42 }))

    # Hash subclass support
    subclass_hash = Class.new(Hash).new
    subclass_hash["B"] = 99
    assert_equal([nil, 99], Xlsxrb::DslHelpers.normalize_row_values(subclass_hash))
  end

  # --- normalize_row_styles ---

  test "normalize_row_styles returns non-hash values as-is without allocations" do
    arr = %i[s1 s2]
    assert_same(arr, Xlsxrb::DslHelpers.normalize_row_styles(arr))
    assert_nil(Xlsxrb::DslHelpers.normalize_row_styles(nil))
    assert_equal(:bold, Xlsxrb::DslHelpers.normalize_row_styles(:bold))
  end

  test "normalize_row_styles converts hash with ranges and arrays to sparse array" do
    assert_equal([], Xlsxrb::DslHelpers.normalize_row_styles({}))
    assert_equal([:header], Xlsxrb::DslHelpers.normalize_row_styles({ "A" => :header }))
    assert_equal([:s1, :s1, :s1, nil, :s2], Xlsxrb::DslHelpers.normalize_row_styles({ 0..2 => :s1, 4 => :s2 }))
    assert_equal([nil, :s3, nil, :s3], Xlsxrb::DslHelpers.normalize_row_styles({ [1, 3] => :s3 }))

    # Column letter array / range requiring column_index conversion
    assert_equal([nil, :s_b, nil, :s_b], Xlsxrb::DslHelpers.normalize_row_styles({ %w[B D] => :s_b }))
    assert_equal(%i[s_ab s_ab], Xlsxrb::DslHelpers.normalize_row_styles({ ("A".."B") => :s_ab }))

    # Hash, Range, and Array subclass support
    subclass_hash = Class.new(Hash).new
    subclass_hash["A"] = :h1
    assert_equal([:h1], Xlsxrb::DslHelpers.normalize_row_styles(subclass_hash))

    subclass_range = Class.new(Range).new(0, 1)
    assert_equal(%i[r1 r1], Xlsxrb::DslHelpers.normalize_row_styles({ subclass_range => :r1 }))

    subclass_array = Class.new(Array).new
    subclass_array << 0
    assert_equal([:a1], Xlsxrb::DslHelpers.normalize_row_styles({ subclass_array => :a1 }))
  end

  # --- validate_row_bounds! ---

  test "validate_row_bounds! validates limits under strict mode" do
    # Valid boundaries
    assert_nil(Xlsxrb::DslHelpers.validate_row_bounds!(0, 20.0))
    assert_nil(Xlsxrb::DslHelpers.validate_row_bounds!(0, 20.0, strict_excel_mode: true))
    assert_nil(Xlsxrb::DslHelpers.validate_row_bounds!(1_048_575, 409, strict_excel_mode: true))
    assert_nil(Xlsxrb::DslHelpers.validate_row_bounds!(1_048_575, 0, strict_excel_mode: true))
    assert_nil(Xlsxrb::DslHelpers.validate_row_bounds!(0, nil, strict_excel_mode: true))

    # Row index exceeds Excel limit (boundary and strictly greater, with and without default param)
    err = assert_raises(ArgumentError) do
      Xlsxrb::DslHelpers.validate_row_bounds!(1_048_576, 20.0)
    end
    assert_equal("Row index 1048576 exceeds Excel limit of 1,048,576 rows", err.message)

    err = assert_raises(ArgumentError) do
      Xlsxrb::DslHelpers.validate_row_bounds!(1_048_577, 20.0, strict_excel_mode: true)
    end
    assert_equal("Row index 1048577 exceeds Excel limit of 1,048,576 rows", err.message)

    # Row height negative
    err = assert_raises(ArgumentError) do
      Xlsxrb::DslHelpers.validate_row_bounds!(0, -0.1, strict_excel_mode: true)
    end
    assert_equal("Row height -0.1 must be between 0 and 409 points (Excel limitation)", err.message)

    # Row height exceeds 409
    err = assert_raises(ArgumentError) do
      Xlsxrb::DslHelpers.validate_row_bounds!(0, 409.1, strict_excel_mode: true)
    end
    assert_equal("Row height 409.1 must be between 0 and 409 points (Excel limitation)", err.message)

    # Lenient mode bypasses all checks
    assert_nil(Xlsxrb::DslHelpers.validate_row_bounds!(2_000_000, 1000, strict_excel_mode: false))
    assert_nil(Xlsxrb::DslHelpers.validate_row_bounds!(2_000_000, -10, strict_excel_mode: false))
  end

  # --- compress_runs & complement_runs ---

  test "compress_runs compresses integers and ranges into minimal disjoint runs" do
    assert_equal([], Xlsxrb::DslHelpers.compress_runs(nil))
    assert_equal([], Xlsxrb::DslHelpers.compress_runs([]))
    assert_equal([[5, 5]], Xlsxrb::DslHelpers.compress_runs(5))
    assert_equal([[2, 10]], Xlsxrb::DslHelpers.compress_runs(2..10))
    assert_equal([[2, 9]], Xlsxrb::DslHelpers.compress_runs(2...10))
    assert_equal([], Xlsxrb::DslHelpers.compress_runs(10..2))
    assert_equal([[3, 3]], Xlsxrb::DslHelpers.compress_runs(3..3))

    assert_equal([[1, 3], [5, 6], [8, 8]], Xlsxrb::DslHelpers.compress_runs([1, 2, 3, 5, 6, 8]))
    assert_equal([[1, 5], [8, 10]], Xlsxrb::DslHelpers.compress_runs([8..10, 1..3, 3..5]))
    # Contained interval (e <= merged.last[1]) and non-integer element
    assert_equal([[1, 10]], Xlsxrb::DslHelpers.compress_runs([1..10, 3..5, 2...4, "ignored"]))

    # Subclass support for Array and Range
    subclass_range = Class.new(Range).new(1, 3)
    subclass_array = Class.new(Array).new([1, 2, 5])
    assert_equal([[1, 3]], Xlsxrb::DslHelpers.compress_runs(subclass_range))
    assert_equal([[1, 3]], Xlsxrb::DslHelpers.compress_runs([subclass_range]))
    assert_equal([[1, 2], [5, 5]], Xlsxrb::DslHelpers.compress_runs(subclass_array))
  end

  test "complement_runs computes intervals not in excluded" do
    assert_equal([[4, 4], [7, 8], [10, 10]], Xlsxrb::DslHelpers.complement_runs([2, 3, 5, 6, 9, 11], 2, 11))
    assert_equal([[3, 5]], Xlsxrb::DslHelpers.complement_runs([2], 2, 5))
    assert_equal([[1, 4]], Xlsxrb::DslHelpers.complement_runs([5], 1, 5))
    assert_equal([[5, 5]], Xlsxrb::DslHelpers.complement_runs([1..4], 1, 5))
    assert_equal([[6, 10]], Xlsxrb::DslHelpers.complement_runs([1..5], 5, 10))
    assert_equal([[1, 4]], Xlsxrb::DslHelpers.complement_runs([5..10], 1, 5))
    assert_equal([[1, 5]], Xlsxrb::DslHelpers.complement_runs([6, 8], 1, 5))
    assert_equal([[5, 10]], Xlsxrb::DslHelpers.complement_runs([1, 2], 5, 10))
    assert_equal([[1, 10]], Xlsxrb::DslHelpers.complement_runs([], 1, 10))
    assert_equal([[1, 10]], Xlsxrb::DslHelpers.complement_runs(nil, 1, 10))
    assert_equal([], Xlsxrb::DslHelpers.complement_runs(1..10, 1, 10))
    assert_equal([[1, 2], [5, 6]], Xlsxrb::DslHelpers.complement_runs([3..4], 1, 6))
    assert_equal([[1, 10]], Xlsxrb::DslHelpers.complement_runs([15..20, -5..0], 1, 10))
  end

  # --- area_ref & chunked_sqrefs ---

  test "area_ref formats single cell and rectangular range" do
    assert_equal("A1", Xlsxrb::DslHelpers.area_ref([0, 0], [0, 0]))
    assert_equal("A1:A5", Xlsxrb::DslHelpers.area_ref([0, 0], [0, 4]))
    assert_equal("A1:C1", Xlsxrb::DslHelpers.area_ref([0, 2], [0, 0]))
    assert_equal("B2:C5", Xlsxrb::DslHelpers.area_ref([1, 2], [1, 4]))
  end

  test "chunked_sqrefs chunks references by limit" do
    assert_equal([], Xlsxrb::DslHelpers.chunked_sqrefs([], []))
    col_runs = [[0, 0], [2, 2]]
    row_runs = [[1, 2], [4, 5]]
    # 2 cols * 2 rows = 4 areas: A2:A3, A5:A6, C2:C3, C5:C6
    assert_equal(["A2:A3 A5:A6 C2:C3 C5:C6"], Xlsxrb::DslHelpers.chunked_sqrefs(col_runs, row_runs))
    assert_equal(["A2:A3 A5:A6", "C2:C3 C5:C6"], Xlsxrb::DslHelpers.chunked_sqrefs(col_runs, row_runs, 2))
  end

  # --- quote_sheet_name, format_absolute_reference & normalize_defined_name_value ---

  test "quote_sheet_name wraps and escapes single quotes" do
    assert_equal("'Sheet1'", Xlsxrb::DslHelpers.quote_sheet_name("Sheet1"))
    assert_equal("'Sheet1'", Xlsxrb::DslHelpers.quote_sheet_name(:Sheet1))
    assert_equal("'O''Connor'", Xlsxrb::DslHelpers.quote_sheet_name("O'Connor"))
    assert_equal("'O''Connor''s Pub'", Xlsxrb::DslHelpers.quote_sheet_name("O'Connor's Pub"))
  end

  test "format_absolute_reference converts valid coordinates and rejects non-coordinates" do
    assert_equal("$A$1", Xlsxrb::DslHelpers.format_absolute_reference("a1"))
    assert_equal("$A$1:$B$2", Xlsxrb::DslHelpers.format_absolute_reference("a1:b2"))
    assert_equal("$XFD$1:$XFD$1048576", Xlsxrb::DslHelpers.format_absolute_reference("xfd1:xfd1048576"))
    assert_equal("$XFD$1048576", Xlsxrb::DslHelpers.format_absolute_reference("$xfd$1048576"))
    assert_equal("$A:$B", Xlsxrb::DslHelpers.format_absolute_reference("a:b"))
    assert_equal("$XFD:$XFD", Xlsxrb::DslHelpers.format_absolute_reference("$xfd:$xfd"))
    assert_equal("$1:$10", Xlsxrb::DslHelpers.format_absolute_reference("1:10"))
    assert_equal("$10:$20", Xlsxrb::DslHelpers.format_absolute_reference("$10:$20"))

    assert_nil(Xlsxrb::DslHelpers.format_absolute_reference("ABCD1"))
    assert_nil(Xlsxrb::DslHelpers.format_absolute_reference("A1:ABCD2"))
    assert_nil(Xlsxrb::DslHelpers.format_absolute_reference("ABCD:B"))
    assert_nil(Xlsxrb::DslHelpers.format_absolute_reference("A:ABCD"))
    assert_nil(Xlsxrb::DslHelpers.format_absolute_reference("A1\nB2"))
    assert_nil(Xlsxrb::DslHelpers.format_absolute_reference("A:B\nC:D"))
    assert_nil(Xlsxrb::DslHelpers.format_absolute_reference("1:10\n2:20"))
    assert_nil(Xlsxrb::DslHelpers.format_absolute_reference("0.05"))
  end

  test "normalize_defined_name_value auto-formats bare references and coordinates" do
    assert_equal("", Xlsxrb::DslHelpers.normalize_defined_name_value)
    assert_equal("$A$1:$A$10", Xlsxrb::DslHelpers.normalize_defined_name_value("A1:A10"))
    assert_equal("'Data'!$A$1:$A$10", Xlsxrb::DslHelpers.normalize_defined_name_value("A1:A10", sheet: "Data"))
    assert_equal("'Data'!$A$1", Xlsxrb::DslHelpers.normalize_defined_name_value("A1", sheet: "Data"))
    assert_equal("'O''Connor'!$A$1:$A$10", Xlsxrb::DslHelpers.normalize_defined_name_value(range: :"A1:A10", sheet: "O'Connor"))
    assert_equal("'Sheet1'!$A$2:$A$10", Xlsxrb::DslHelpers.normalize_defined_name_value(columns: 0, rows: 1..9, sheet: "Sheet1"))
    assert_equal("'Sheet1'!$C$2:$C$9", Xlsxrb::DslHelpers.normalize_defined_name_value(columns: "C", rows: 1...9, sheet: "Sheet1"))
    assert_equal("'Sheet1'!$A$2:$A$5", Xlsxrb::DslHelpers.normalize_defined_name_value(rows: Class.new(Range).new(1, 4), sheet: "Sheet1"))
    assert_equal("'Sheet1'!$C$5", Xlsxrb::DslHelpers.normalize_defined_name_value(columns: "C", rows: 4, sheet: "Sheet1"))
    assert_equal("'Sheet1'!$A$5", Xlsxrb::DslHelpers.normalize_defined_name_value(rows: 4, sheet: "Sheet1"))
    assert_equal("'Sheet1'!$C:$C", Xlsxrb::DslHelpers.normalize_defined_name_value(columns: "C", sheet: "Sheet1"))
    assert_equal("'Sheet1'!$A:$B", Xlsxrb::DslHelpers.normalize_defined_name_value("A:B", sheet: "Sheet1"))
    assert_equal("'Sheet1'!$1:$10", Xlsxrb::DslHelpers.normalize_defined_name_value("1:10", sheet: "Sheet1"))

    # Leaves formulas, constants, and already-qualified references untouched
    assert_equal("0.10", Xlsxrb::DslHelpers.normalize_defined_name_value("0.10", sheet: "Data"))
    assert_equal("=SUM($A$1:$A$10)", Xlsxrb::DslHelpers.normalize_defined_name_value("=SUM($A$1:$A$10)", sheet: "Data"))
    assert_equal("Sheet1!$A$1:$B$2", Xlsxrb::DslHelpers.normalize_defined_name_value("Sheet1!$A$1:$B$2"))
    assert_equal("'Other'!$A$1:$B$2", Xlsxrb::DslHelpers.normalize_defined_name_value("'Other'!$A$1:$B$2", sheet: "Data"))
  end

  # --- normalize_data_validation ---

  test "normalize_data_validation handles list array and message newlines" do
    res = Xlsxrb::DslHelpers.normalize_data_validation("A1:A10", list: %w[Active Inactive], prompt: "Line1\nLine2", error: "Err1\r\nErr2")
    assert_equal(1, res.size)
    dv = res.first
    assert_equal("A1:A10", dv[:sqref])
    assert_equal("list", dv[:type])
    assert_equal('"Active,Inactive"', dv[:formula1])
    assert_equal("Line1_x000a_Line2", dv[:prompt])
    assert_equal("Err1_x000a_Err2", dv[:error])
    assert_equal(true, dv[:show_input_message])
    assert_equal(true, dv[:show_error_message])
  end

  test "normalize_data_validation respects explicit show_message false" do
    res = Xlsxrb::DslHelpers.normalize_data_validation("A1", prompt: "Hello", show_input_message: false)
    assert_equal(false, res.first[:show_input_message])
  end

  test "normalize_data_validation handles coordinates and except_rows and deletes columns keys" do
    res = Xlsxrb::DslHelpers.normalize_data_validation(columns: 0, rows: 1..4, except_rows: 2, list: %w[Yes No])
    assert_equal(1, res.size)
    assert_equal("A2 A4:A5", res.first[:sqref])
    assert_nil(res.first[:columns])
    assert_nil(res.first[:rows])
  end

  # --- sheet_name ---

  test "sanitize_sheet_name with nil, empty string, or whitespace defaults to 'Sheet'" do
    assert_equal("Sheet", Xlsxrb::DslHelpers.sanitize_sheet_name(nil))
    assert_equal("Sheet", Xlsxrb::DslHelpers.sanitize_sheet_name(""))
    assert_equal("Sheet", Xlsxrb::DslHelpers.sanitize_sheet_name("   "))
    assert_equal("Sales", Xlsxrb::DslHelpers.sanitize_sheet_name("  Sales  "))
  end

  test "sanitize_sheet_name replaces forbidden characters and control characters" do
    assert_equal("Data_1__2026_10_24__", Xlsxrb::DslHelpers.sanitize_sheet_name("Data[1]:2026/10/24*?"))
    assert_equal("Back_slash", Xlsxrb::DslHelpers.sanitize_sheet_name('Back\slash'))
    assert_equal("Ctrl_char", Xlsxrb::DslHelpers.sanitize_sheet_name("Ctrl\x00char"))
    assert_equal("Ctrl_char", Xlsxrb::DslHelpers.sanitize_sheet_name("Ctrl\x08char"))
    assert_equal("Ctrl_char", Xlsxrb::DslHelpers.sanitize_sheet_name("Ctrl\x1Fchar"))
  end

  test "sanitize_sheet_name replaces leading and trailing apostrophes but keeps inner apostrophes" do
    assert_equal("_Sales_", Xlsxrb::DslHelpers.sanitize_sheet_name("'Sales'"))
    assert_equal("__Sales__", Xlsxrb::DslHelpers.sanitize_sheet_name("''Sales''"))
    assert_equal("John's Sales", Xlsxrb::DslHelpers.sanitize_sheet_name("John's Sales"))
    assert_equal("_", Xlsxrb::DslHelpers.sanitize_sheet_name("'"))
    assert_equal("__", Xlsxrb::DslHelpers.sanitize_sheet_name("''"))
  end

  test "sanitize_sheet_name truncates to 31 chars without premature truncation" do
    long_name = "A" * 50
    assert_equal("A" * 31, Xlsxrb::DslHelpers.sanitize_sheet_name(long_name))

    # Replaces trailing apostrophe if 31-char truncation leaves one at the end
    trunc_apostrophe = "#{"A" * 30}'DEF"
    assert_equal("#{"A" * 30}_", Xlsxrb::DslHelpers.sanitize_sheet_name(trunc_apostrophe))
    trunc_multi_apostrophe = "#{"A" * 29}''DEF"
    assert_equal("#{"A" * 29}__", Xlsxrb::DslHelpers.sanitize_sheet_name(trunc_multi_apostrophe))
  end

  test "sanitize_sheet_name avoids reserved name History by appending replace_char" do
    assert_equal("History_", Xlsxrb::DslHelpers.sanitize_sheet_name("History"))
    assert_equal("history_", Xlsxrb::DslHelpers.sanitize_sheet_name("history"))
    assert_equal("HISTORY_", Xlsxrb::DslHelpers.sanitize_sheet_name("HISTORY"))
    assert_equal("History-", Xlsxrb::DslHelpers.sanitize_sheet_name("History", replace_char: "-"))
  end

  test "sanitize_sheet_name supports custom replace_char" do
    assert_equal("Sales-2026", Xlsxrb::DslHelpers.sanitize_sheet_name("Sales:2026", replace_char: "-"))
    assert_equal("-Sales", Xlsxrb::DslHelpers.sanitize_sheet_name("'Sales", replace_char: "-"))
    assert_equal("--Sales", Xlsxrb::DslHelpers.sanitize_sheet_name("''Sales", replace_char: "-"))
  end

  # --- extract_dxf_from_rule! ---

  test "extract_dxf_from_rule! extracts font and fill styles from rule" do
    rule = {
      type: :cell_is,
      font_color: "FF0000",
      bold: true,
      italic: false,
      underline: "single",
      font_size: 12,
      fill_color: "FFFF00",
      formula: "100"
    }
    dxf = Xlsxrb::DslHelpers.extract_dxf_from_rule!(rule)
    assert_equal({ color: "FF0000", bold: true, italic: false, underline: "single", font_size: 12 }, dxf[:font])
    assert_equal({ pattern: "solid", fg_color: "FFFF00" }, dxf[:fill])
    assert_equal({ type: :cell_is, formula: "100" }, rule)

    # With font_size key
    rule_fs = { font_size: 14 }
    dxf_fs = Xlsxrb::DslHelpers.extract_dxf_from_rule!(rule_fs)
    assert_equal({ font_size: 14 }, dxf_fs[:font])

    # Empty when no font or fill
    empty_rule = { type: :cell_is, formula: "50" }
    assert_nil(Xlsxrb::DslHelpers.extract_dxf_from_rule!(empty_rule))
  end

  # --- normalize_conditional_format ---

  test "normalize_conditional_format with sqref and symbol options" do
    rules = Xlsxrb::DslHelpers.normalize_conditional_format(
      "A1:A10",
      type: :cell_is,
      operator: :greater_than,
      formula: 100
    )
    assert_equal(1, rules.size)
    assert_equal("A1:A10", rules.first[:sqref])
    assert_equal(:cell_is, rules.first[:type])
    assert_equal(:greater_than, rules.first[:operator])
    assert_equal(100, rules.first[:formula])
  end

  test "normalize_conditional_format with coordinate targeting" do
    rules = Xlsxrb::DslHelpers.normalize_conditional_format(
      columns: 0..1,
      rows: 1..9,
      type: :cell_is,
      operator: :less_than,
      formula: 50
    )
    assert_equal(1, rules.size)
    assert_equal("A2:B10", rules.first[:sqref])
    assert_equal(:cell_is, rules.first[:type])
    assert_equal(:less_than, rules.first[:operator])

    # with except_rows
    rules_exc = Xlsxrb::DslHelpers.normalize_conditional_format(
      columns: 2,
      rows: 0..4,
      except_rows: 2,
      type: :expression,
      formula: "C1>0"
    )
    assert_equal("C1:C2 C4:C5", rules_exc.first[:sqref])
  end

  test "normalize_conditional_format auto-completes 2-color and 3-color scales" do
    # 2-color scale
    rules2 = Xlsxrb::DslHelpers.normalize_conditional_format(
      "B1:B20",
      type: :color_scale,
      colors: %w[FFFFFF 5B9BD5]
    )
    cs2 = rules2.first[:color_scale]
    assert_equal(%w[FFFFFFFF FF5B9BD5], cs2[:colors])
    assert_equal([{ type: "min" }, { type: "max" }], cs2[:cfvo])

    # 3-color scale
    rules3 = Xlsxrb::DslHelpers.normalize_conditional_format(
      "B1:B20",
      color_scale: { colors: %w[F8696B FFEB84 63BE7B] }
    )
    cs3 = rules3.first[:color_scale]
    assert_equal(%w[FFF8696B FFFFEB84 FF63BE7B], cs3[:colors])
    assert_equal([{ type: "min" }, { type: "percent", val: "50" }, { type: "max" }], cs3[:cfvo])
  end

  test "normalize_conditional_format auto-completes icon_set with 3, 4, 5 icons" do
    rules3 = Xlsxrb::DslHelpers.normalize_conditional_format("A1:A5", type: :icon_set, icon_style: :three_arrows)
    is3 = rules3.first[:icon_set]
    assert_equal("3Arrows", is3[:icon_set])
    assert_equal(3, is3[:cfvo].size)

    rules4 = Xlsxrb::DslHelpers.normalize_conditional_format("A1:A5", type: :icon_set, icon_style: :four_rating)
    is4 = rules4.first[:icon_set]
    assert_equal("4Rating", is4[:icon_set])
    assert_equal(4, is4[:cfvo].size)
    assert_equal(%w[0 25 50 75], is4[:cfvo].map { |c| c[:val] })

    rules5 = Xlsxrb::DslHelpers.normalize_conditional_format("A1:A5", type: :icon_set, icon_style: :five_quarters)
    is5 = rules5.first[:icon_set]
    assert_equal("5Quarters", is5[:icon_set])
    assert_equal(5, is5[:cfvo].size)
    assert_equal(%w[0 20 40 60 80], is5[:cfvo].map { |c| c[:val] })
  end

  test "normalize_conditional_format with block form" do
    rules = Xlsxrb::DslHelpers.normalize_conditional_format("C1:C10") do |cf|
      cf.type = :cell_is
      cf.operator = :greater_than
      cf.formula = 100
      cf.bold = true
      cf.fill_color = "FF0000"
      cf.above_average = true
      cf.equal_average = true
      cf.std_dev = 2
      cf.rank = 5
      cf.percent = true
      cf.bottom = true
      cf.text = "urgent"
      cf.time_period = :this_month
    end
    assert_equal(1, rules.size)
    assert_equal("C1:C10", rules.first[:sqref])
    assert_equal(:cell_is, rules.first[:type])
    assert_equal(:greater_than, rules.first[:operator])
    assert_equal(100, rules.first[:formula])
    assert_equal(true, rules.first[:bold])
    assert_equal("FF0000", rules.first[:fill_color])
    assert_equal(true, rules.first[:above_average])
    assert_equal(true, rules.first[:equal_average])
    assert_equal(2, rules.first[:std_dev])
    assert_equal(5, rules.first[:rank])
    assert_equal(true, rules.first[:percent])
    assert_equal(true, rules.first[:bottom])
    assert_equal("urgent", rules.first[:text])
    assert_equal("thisMonth", rules.first[:time_period])
  end

  test "normalize_conditional_format infers type and normalizes time_period" do
    # Time period inference and symbol mapping
    rules_tp = Xlsxrb::DslHelpers.normalize_conditional_format("D1:D10", time_period: :last_7_days)
    assert_equal(:time_period, rules_tp.first[:type])
    assert_equal("last7Days", rules_tp.first[:time_period])

    # Rank inference -> top10
    rules_top = Xlsxrb::DslHelpers.normalize_conditional_format("E1:E10", rank: 10, percent: true)
    assert_equal(:top10, rules_top.first[:type])
    assert_equal(10, rules_top.first[:rank])
    assert_equal(true, rules_top.first[:percent])

    # Above average inference
    rules_avg = Xlsxrb::DslHelpers.normalize_conditional_format("F1:F10", above_average: false, equal_average: true)
    assert_equal(:above_average, rules_avg.first[:type])
    assert_equal(false, rules_avg.first[:above_average])
    assert_equal(true, rules_avg.first[:equal_average])

    # Text inference
    rules_txt = Xlsxrb::DslHelpers.normalize_conditional_format("G1:G10", text: "error")
    assert_equal(:contains_text, rules_txt.first[:type])
    assert_equal("error", rules_txt.first[:text])

    # Error rule types
    rules_err = Xlsxrb::DslHelpers.normalize_conditional_format("H1:H10", type: :contains_errors)
    assert_equal(:contains_errors, rules_err.first[:type])
  end

  test "normalize_conditional_format raises when neither sqref nor coordinates provided" do
    assert_raises(ArgumentError) do
      Xlsxrb::DslHelpers.normalize_conditional_format(type: :cell_is)
    end
  end
end
