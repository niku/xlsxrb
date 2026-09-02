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
end
