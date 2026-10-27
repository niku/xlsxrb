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

  test "normalize_merge_range with hash argument" do
    # Single row range with col indices
    range_hash1 = { row: 0, col_start: 0, col_end: 2 }
    assert_equal("A1:C1", Xlsxrb::DslHelpers.normalize_merge_range(range_hash1))

    # Multi-row range with col letters
    range_hash2 = { row_start: 1, row_end: 4, col_start: "B", col_end: "D" }
    assert_equal("B2:D5", Xlsxrb::DslHelpers.normalize_merge_range(range_hash2))
  end

  test "normalize_merge_range with keyword arguments" do
    res = Xlsxrb::DslHelpers.normalize_merge_range(row: 2, col_start: "A", col_end: :C)
    assert_equal("A3:C3", res)

    res_rect = Xlsxrb::DslHelpers.normalize_merge_range(row_start: 0, row_end: 2, col_start: 0, col_end: 1)
    assert_equal("A1:B3", res_rect)
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
    assert_equal("A1", Xlsxrb::DslHelpers.area_ref([0, 0], [1, 1]))
    assert_equal("A1:A5", Xlsxrb::DslHelpers.area_ref([0, 0], [1, 5]))
    assert_equal("A1:C1", Xlsxrb::DslHelpers.area_ref([0, 2], [1, 1]))
    assert_equal("B2:C5", Xlsxrb::DslHelpers.area_ref([1, 2], [2, 5]))
  end

  test "chunked_sqrefs chunks references by limit" do
    assert_equal([], Xlsxrb::DslHelpers.chunked_sqrefs([], []))
    col_runs = [[0, 0], [2, 2]]
    row_runs = [[2, 3], [5, 6]]
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
    assert_equal("'Sheet1'!$A$2:$A$10", Xlsxrb::DslHelpers.normalize_defined_name_value(col: 0, rows: 2..10, sheet: "Sheet1"))
    assert_equal("'Sheet1'!$C$2:$C$9", Xlsxrb::DslHelpers.normalize_defined_name_value(col: "C", rows: 2...10, sheet: "Sheet1"))
    assert_equal("'Sheet1'!$A$2:$A$5", Xlsxrb::DslHelpers.normalize_defined_name_value(rows: Class.new(Range).new(2, 5), sheet: "Sheet1"))
    assert_equal("'Sheet1'!$C$5", Xlsxrb::DslHelpers.normalize_defined_name_value(col: "C", rows: 5, sheet: "Sheet1"))
    assert_equal("'Sheet1'!$A$5", Xlsxrb::DslHelpers.normalize_defined_name_value(rows: 5, sheet: "Sheet1"))
    assert_equal("'Sheet1'!$C:$C", Xlsxrb::DslHelpers.normalize_defined_name_value(col: "C", sheet: "Sheet1"))
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
    res = Xlsxrb::DslHelpers.normalize_data_validation("A1:A10", in: %w[Active Inactive], prompt: "Line1\nLine2", error: "Err1\r\nErr2")
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

  test "normalize_data_validation handles coordinates and except_rows and deletes both col and columns keys" do
    res = Xlsxrb::DslHelpers.normalize_data_validation(columns: 0, col: 0, rows: 2..5, row: 2..5, except_rows: 3, in: %w[Yes No])
    assert_equal(1, res.size)
    assert_equal("A2 A4:A5", res.first[:sqref])
    assert_nil(res.first[:columns])
    assert_nil(res.first[:col])
    assert_nil(res.first[:rows])
    assert_nil(res.first[:row])
  end

  # --- sheet_name ---

  test "sanitize_sheet_name with nil, empty string, or whitespace defaults to 'Sheet'" do
    assert_equal("Sheet", Xlsxrb::DslHelpers.sanitize_sheet_name(nil))
    assert_equal("Sheet", Xlsxrb::DslHelpers.sanitize_sheet_name(""))
    assert_equal("Sheet", Xlsxrb::DslHelpers.sanitize_sheet_name("   "))
    assert_equal("Sales", Xlsxrb::DslHelpers.sanitize_sheet_name("  Sales  "))
    # Alias sheet_name
    assert_equal("Sheet", Xlsxrb::DslHelpers.sheet_name(""))
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
end
