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
end
