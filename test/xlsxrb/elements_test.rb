# frozen_string_literal: true

require "test_helper"

class ElementsTest < Test::Unit::TestCase
  cover Xlsxrb::Elements::Cell
  cover Xlsxrb::Elements::Row
  cover Xlsxrb::Elements::Column
  cover Xlsxrb::Elements::Worksheet
  cover Xlsxrb::Elements::CoordinateAccess
  cover Xlsxrb::Elements::Workbook
  cover Xlsxrb::Elements::CellError
  cover Xlsxrb::Elements::RichText
  cover Xlsxrb::Elements::Formula
  # --- Cell ---

  test "cell creates a valid cell" do
    cell = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "hello")
    assert(cell.valid?)
    assert_equal("hello", cell.value)
    assert_equal(0, cell.row_index)
    assert_equal(0, cell.column_index)
    assert_equal("A1", cell.ref)
  end

  test "cell handles max boundary coordinates (XFD1048576)" do
    max_cell = Xlsxrb::Elements::Cell.new(row_index: 1_048_575, column_index: 16_383, value: "max")
    assert(max_cell.valid?)
    assert_equal("XFD1048576", max_cell.ref)
    assert_equal([1_048_575, 16_383], Xlsxrb::Elements::Cell.parse_ref("XFD1048576"))

    # Out of boundary
    over_cell = Xlsxrb::Elements::Cell.new(row_index: 1_048_576, column_index: 16_384, value: "over")
    refute(over_cell.valid?)
    assert(over_cell.errors.any? { |e| e.include?("row_index must be < 1048576") })
    assert(over_cell.errors.any? { |e| e.include?("column_index must be < 16384") })
  end

  test "cell valid_coordinates? checks boundaries and types" do
    # Valid lower and upper boundaries
    assert_true(Xlsxrb::Elements::Cell.valid_coordinates?(0, 0))
    assert_true(Xlsxrb::Elements::Cell.valid_coordinates?(0, 16_383))
    assert_true(Xlsxrb::Elements::Cell.valid_coordinates?(1_048_575, 0))
    assert_true(Xlsxrb::Elements::Cell.valid_coordinates?(1_048_575, 16_383))
    assert_true(Xlsxrb::Elements::Cell.valid_coordinates?(100, 100))

    # Negative boundaries
    assert_false(Xlsxrb::Elements::Cell.valid_coordinates?(-1, 0))
    assert_false(Xlsxrb::Elements::Cell.valid_coordinates?(0, -1))
    assert_false(Xlsxrb::Elements::Cell.valid_coordinates?(-1, -1))

    # Upper boundary limit violations
    assert_false(Xlsxrb::Elements::Cell.valid_coordinates?(1_048_576, 0))
    assert_false(Xlsxrb::Elements::Cell.valid_coordinates?(0, 16_384))
    assert_false(Xlsxrb::Elements::Cell.valid_coordinates?(1_048_576, 16_384))

    # Non-integer types for row_index
    assert_false(Xlsxrb::Elements::Cell.valid_coordinates?("0", 0))
    assert_false(Xlsxrb::Elements::Cell.valid_coordinates?(nil, 0))
    assert_false(Xlsxrb::Elements::Cell.valid_coordinates?(0.0, 0))
    assert_false(Xlsxrb::Elements::Cell.valid_coordinates?(1.5, 0))
    assert_false(Xlsxrb::Elements::Cell.valid_coordinates?(:zero, 0))
    assert_false(Xlsxrb::Elements::Cell.valid_coordinates?([], 0))

    # Non-integer types for column_index
    assert_false(Xlsxrb::Elements::Cell.valid_coordinates?(0, "0"))
    assert_false(Xlsxrb::Elements::Cell.valid_coordinates?(0, nil))
    assert_false(Xlsxrb::Elements::Cell.valid_coordinates?(0, 0.0))
    assert_false(Xlsxrb::Elements::Cell.valid_coordinates?(0, 1.5))
    assert_false(Xlsxrb::Elements::Cell.valid_coordinates?(0, :zero))
    assert_false(Xlsxrb::Elements::Cell.valid_coordinates?(0, []))

    # Both non-integer
    assert_false(Xlsxrb::Elements::Cell.valid_coordinates?(nil, nil))
  end

  test "cell valid_value? checks supported and unsupported value types" do
    formula = Xlsxrb::Elements::Formula.new(expression: "SUM(A1:A10)")
    rich_text = Xlsxrb::Elements::RichText.new(runs: [{ text: "rich" }])
    cell_err = Xlsxrb::Elements::CellError.new(code: "#VALUE!")

    assert_true(Xlsxrb::Elements::Cell.valid_value?(nil))
    assert_true(Xlsxrb::Elements::Cell.valid_value?("test"))
    assert_true(Xlsxrb::Elements::Cell.valid_value?(100))
    assert_true(Xlsxrb::Elements::Cell.valid_value?(100.5))
    assert_true(Xlsxrb::Elements::Cell.valid_value?(true))
    assert_true(Xlsxrb::Elements::Cell.valid_value?(false))
    assert_true(Xlsxrb::Elements::Cell.valid_value?(Date.new(2026, 1, 1)))
    assert_true(Xlsxrb::Elements::Cell.valid_value?(Time.now))
    assert_true(Xlsxrb::Elements::Cell.valid_value?(formula))
    assert_true(Xlsxrb::Elements::Cell.valid_value?({ formula: "A1" }))
    assert_true(Xlsxrb::Elements::Cell.valid_value?(rich_text))
    assert_true(Xlsxrb::Elements::Cell.valid_value?(cell_err))

    assert_false(Xlsxrb::Elements::Cell.valid_value?([1, 2, 3]))
    assert_false(Xlsxrb::Elements::Cell.valid_value?(:symbol))
    assert_false(Xlsxrb::Elements::Cell.valid_value?(Object.new))
    assert_false(Xlsxrb::Elements::Cell.valid_value?({ not_formula: 1 }))
  end

  test "cell accessors, type conversions, and formula handling" do
    formula = Xlsxrb::Elements::Formula.new(expression: "SUM(A1:A10)")
    cell = Xlsxrb::Elements::Cell.new(row_index: 2, column_index: 3, value: "123.45", formula: formula, style_index: 5)

    assert_equal("123.45", cell.content)
    assert_equal("123.45", cell.to_s)
    cell_int = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: 42)
    assert_equal("42", cell_int.to_s)
    assert_instance_of(String, cell_int.to_s)
    cell_nil = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: nil)
    assert_equal("", cell_nil.to_s)
    assert_equal(123, cell.to_i)
    assert_instance_of(Float, cell.to_f)
    assert_in_delta(123.45, cell.to_f)
    assert_equal("123.45", cell[:value])
    assert_equal("SUM(A1:A10)", cell[:formula])
    assert_equal(5, cell[:style_index])
    assert_equal("D3", cell[:ref])
    assert_equal(3, cell[:column_index])
    assert_equal(2, cell[:row_index])
    assert_equal("s", cell[:type])

    # String formula
    cell_str_f = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: true, formula: "A1+1")
    assert_equal("A1+1", cell_str_f[:formula])
    assert_equal("b", cell_str_f[:type])

    # Boolean false and numeric type
    cell_false = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: false)
    assert_equal("b", cell_false[:type])
    cell_num = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: 42)
    assert_nil(cell_num[:type])
    assert_nil(cell_num[:unknown_key])

    # to_date conversions
    date_val = Date.new(2025, 5, 20)
    assert_equal(date_val, Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: date_val).to_date)
    assert_equal(Date.new(2026, 1, 1), Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: 46_023).to_date)
    assert_equal(Date.new(2025, 12, 31), Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "2025-12-31").to_date)
    assert_nil(Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "invalid-date").to_date)

    # to_time conversions
    time_val = Time.new(2025, 5, 20, 10, 30, 0, "+00:00")
    assert_equal(time_val, Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: time_val).to_time)
    assert_equal(Time.new(2026, 1, 1, 12, 0, 0, "+00:00"), Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: 46_023.5).to_time)
    assert_equal(Time.parse("2025-12-31 15:00:00"), Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "2025-12-31 15:00:00").to_time)
    assert_nil(Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "invalid-time").to_time)
  end

  test "cell formula predicate and expression methods" do
    cell_no_f = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: 42)
    assert_false(cell_no_f.formula?)
    assert_nil(cell_no_f.formula_expression)

    cell_str_f = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: 42, formula: "SUM(A1:A10)")
    assert_true(cell_str_f.formula?)
    assert_equal("SUM(A1:A10)", cell_str_f.formula_expression)

    f_obj = Xlsxrb::Elements::Formula.new(expression: "B1*2", cached_value: 100)
    cell_obj_f = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: 100, formula: f_obj)
    assert_true(cell_obj_f.formula?)
    assert_equal("B1*2", cell_obj_f.formula_expression)

    cell_empty_f = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: 42, formula: "")
    assert_false(cell_empty_f.formula?)
    assert_nil(cell_empty_f.formula_expression)
  end

  test "cell equality, hashing, with copy, and pattern matching" do
    c1 = Xlsxrb::Elements::Cell.new(row_index: 1, column_index: 1, value: 10, formula: "A1", style_index: 1)
    c2 = Xlsxrb::Elements::Cell.new(row_index: 1, column_index: 1, value: 10, formula: "A1", style_index: 1)
    c3 = Xlsxrb::Elements::Cell.new(row_index: 1, column_index: 1, value: 20, formula: "A1", style_index: 1)

    assert_equal(c1, c2)
    assert(c1.eql?(c2))
    assert_equal(c1.hash, c2.hash)
    refute_equal(c1, c3)
    refute_equal(c1, "not-a-cell")

    # with
    c1_modified = c1.with(value: 99, style_index: 2)
    assert_equal(99, c1_modified.value)
    assert_equal(2, c1_modified.style_index)
    assert_equal(c1.row_index, c1_modified.row_index)
    assert_equal(c1.formula, c1_modified.formula)

    # pattern matching (deconstruct & deconstruct_keys)
    assert_equal([1, 1, 10, "A1", 1, {}, [], nil, nil, nil, nil], c1.deconstruct)
    assert_equal({ row_index: 1, value: 10 }, c1.deconstruct_keys(%i[row_index value]))
    assert_equal({ row_index: 1, column_index: 1, value: 10, formula: "A1", style_index: 1, unmapped_data: {}, errors: [], hyperlink: nil, comment: nil, raw_value: nil, format_code: nil }, c1.deconstruct_keys(nil))
  end

  test "cell hyperlink predicate and url accessor" do
    cell_no_link = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "Text")
    assert_false(cell_no_link.link?)
    assert_false(cell_no_link.hyperlink?)
    assert_nil(cell_no_link.url)
    assert_nil(cell_no_link.hyperlink)
    assert_nil(cell_no_link[:hyperlink])

    cell_link = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "Google", hyperlink: { url: "https://google.com", tooltip: "Search" })
    assert_true(cell_link.link?)
    assert_true(cell_link.hyperlink?)
    assert_equal("https://google.com", cell_link.url)
    assert_equal({ url: "https://google.com", tooltip: "Search" }, cell_link.hyperlink)
    assert_equal({ url: "https://google.com", tooltip: "Search" }, cell_link[:hyperlink])

    cell_loc = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, hyperlink: { location: "Sheet2!A1" })
    assert_true(cell_loc.link?)
    assert_equal("Sheet2!A1", cell_loc.url)

    cell_str = cell_no_link.with(hyperlink: "https://ruby-lang.org")
    assert_true(cell_str.link?)
    assert_equal("https://ruby-lang.org", cell_str.url)

    assert_false(Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, hyperlink: "").link?)
    assert_false(Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, hyperlink: {}).link?)
    assert_nil(Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, hyperlink: "").url)
  end

  test "cell comment predicate and text accessor" do
    cell_no_comment = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "Text")
    assert_false(cell_no_comment.comment?)
    assert_nil(cell_no_comment.comment)
    assert_nil(cell_no_comment.comment_text)
    assert_nil(cell_no_comment[:comment])

    cell_comment = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "Text", comment: { author: "Author", text: "Important Note" })
    assert_true(cell_comment.comment?)
    assert_equal("Important Note", cell_comment.comment_text)
    assert_equal({ author: "Author", text: "Important Note" }, cell_comment.comment)
    assert_equal({ author: "Author", text: "Important Note" }, cell_comment[:comment])

    cell_str = cell_no_comment.with(comment: "Inline note")
    assert_true(cell_str.comment?)
    assert_equal("Inline note", cell_str.comment_text)

    assert_false(Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, comment: "").comment?)
    assert_false(Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, comment: {}).comment?)
    assert_nil(Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, comment: "").comment_text)
  end

  test "worksheet hyperlinks and hyperlink coordinate lookup" do
    links = { "A1" => { url: "https://example.com" }, "B2" => { location: "Sheet2!A1" } }
    ws = Xlsxrb::Elements::Worksheet.new(name: "Sheet1", hyperlinks: links)

    assert_equal(links, ws.hyperlinks)
    assert_equal({ url: "https://example.com" }, ws.hyperlink("A1"))
    assert_equal({ url: "https://example.com" }, ws.hyperlink(:A1))
    assert_equal({ url: "https://example.com" }, ws.hyperlink(0, 0))
    assert_equal({ location: "Sheet2!A1" }, ws.hyperlink("B2"))
    assert_equal({ location: "Sheet2!A1" }, ws.hyperlink(1, 1))
    assert_nil(ws.hyperlink("C3"))

    updated = ws.update_cell("C3", value: "Test", hyperlink: { url: "https://updated.com" })
    assert_equal({ url: "https://updated.com" }, updated.hyperlink("C3"))
    assert_true(updated["C3"].link?)
    assert_equal("https://updated.com", updated["C3"].url)
  end

  test "worksheet comments and comment coordinate lookup" do
    comments = [{ ref: "A1", author: "Tester", text: "Hello" }, { ref: "B2", author: "Reviewer", text: "Check this" }]
    ws = Xlsxrb::Elements::Worksheet.new(name: "Sheet1", comments: comments)

    assert_equal(comments, ws.comments)
    assert_equal({ ref: "A1", author: "Tester", text: "Hello" }, ws.comment("A1"))
    assert_equal({ ref: "A1", author: "Tester", text: "Hello" }, ws.comment(:A1))
    assert_equal({ ref: "A1", author: "Tester", text: "Hello" }, ws.comment(0, 0))
    assert_equal({ ref: "B2", author: "Reviewer", text: "Check this" }, ws.comment("B2"))
    assert_equal({ ref: "B2", author: "Reviewer", text: "Check this" }, ws.comment(1, 1))
    assert_nil(ws.comment("C3"))

    updated = ws.update_cell("C3", value: "Test", comment: { author: "Me", text: "New comment" })
    assert_equal({ ref: "C3", author: "Me", text: "New comment" }, updated.comment("C3"))
    assert_true(updated["C3"].comment?)
    assert_equal("New comment", updated["C3"].comment_text)
  end

  test "cell with negative row_index is invalid" do
    cell = Xlsxrb::Elements::Cell.new(row_index: -1, column_index: 0, value: "x")
    assert(cell.errors.any? { |e| e.include?("row_index must be a non-negative Integer") && e.include?("-1") })
  end

  test "cell with negative column_index is invalid" do
    cell = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: -1, value: "x")
    assert(cell.errors.any? { |e| e.include?("column_index must be a non-negative Integer") && e.include?("-1") })
  end

  test "cell with too large row_index is invalid" do
    cell = Xlsxrb::Elements::Cell.new(row_index: 1_048_576, column_index: 0)
    assert(cell.errors.any? { |e| e.include?("row_index must be < 1048576") && e.include?("1048576") })
  end

  test "cell with too large column_index is invalid" do
    cell = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 16_384)
    assert(cell.errors.any? { |e| e.include?("column_index must be < 16384") && e.include?("16384") })
  end

  test "cell validate edge cases and all supported/unsupported types" do
    # Non-integer indices
    errs1 = Xlsxrb::Elements::Cell.validate("0", :zero, "val")
    assert(errs1.any? { |e| e.include?("row_index must be a non-negative Integer") })
    assert(errs1.any? { |e| e.include?("column_index must be a non-negative Integer") })

    # Unsupported value type
    errs2 = Xlsxrb::Elements::Cell.validate(0, 0, Object.new)
    assert(errs2.any? { |e| e.include?("unsupported value type: Object") })

    # All valid types return empty errors
    assert_empty(Xlsxrb::Elements::Cell.validate(0, 0, nil))
    assert_empty(Xlsxrb::Elements::Cell.validate(0, 0, "text"))
    assert_empty(Xlsxrb::Elements::Cell.validate(0, 0, 100))
    assert_empty(Xlsxrb::Elements::Cell.validate(0, 0, 12.34))
    assert_empty(Xlsxrb::Elements::Cell.validate(0, 0, true))
    assert_empty(Xlsxrb::Elements::Cell.validate(0, 0, false))
    assert_empty(Xlsxrb::Elements::Cell.validate(0, 0, Date.today))
    assert_empty(Xlsxrb::Elements::Cell.validate(0, 0, Time.now))
    assert_empty(Xlsxrb::Elements::Cell.validate(0, 0, Xlsxrb::Elements::Formula.new(expression: "SUM(A1)")))
    assert_empty(Xlsxrb::Elements::Cell.validate(0, 0, { formula: "SUM(A1)" }))
    assert_empty(Xlsxrb::Elements::Cell.validate(0, 0, Xlsxrb::Elements::RichText.new(runs: [])))
    assert_empty(Xlsxrb::Elements::Cell.validate(0, 0, Xlsxrb::Elements::CellError.new("#REF!")))

    # Combined: invalid row with supported Formula value should only report row error
    errs3 = Xlsxrb::Elements::Cell.validate(-1, 0, Xlsxrb::Elements::Formula.new(expression: "A1"))
    assert_equal(1, errs3.size)
    assert(errs3.first.include?("row_index must be a non-negative Integer"))
  end

  test "cell column_letter converts index to Excel column" do
    assert_equal("A", Xlsxrb::Elements::Cell.column_letter(0))
    assert_equal("Z", Xlsxrb::Elements::Cell.column_letter(25))
    assert_equal("AA", Xlsxrb::Elements::Cell.column_letter(26))
    assert_equal("AZ", Xlsxrb::Elements::Cell.column_letter(51))
    assert_equal("XFD", Xlsxrb::Elements::Cell.column_letter(16_383))
    assert_equal("XFE", Xlsxrb::Elements::Cell.column_letter(16_384))
    assert_equal("ZZZ", Xlsxrb::Elements::Cell.column_letter(18_277))

    # Cached array identity
    assert_same(Xlsxrb::Elements::Cell.column_letter(0), Xlsxrb::Elements::Cell.column_letter(0))
    assert_same(Xlsxrb::Elements::Cell.column_letter(25), Xlsxrb::Elements::Cell.column_letter(25))
  end

  test "cell column_letter raises ArgumentError for invalid index" do
    err_neg = assert_raises(ArgumentError) { Xlsxrb::Elements::Cell.column_letter(-1) }
    assert_equal("Column index must be a non-negative Integer, got -1", err_neg.message)

    err_str = assert_raises(ArgumentError) { Xlsxrb::Elements::Cell.column_letter("0") }
    assert_equal('Column index must be a non-negative Integer, got "0"', err_str.message)

    err_sym = assert_raises(ArgumentError) { Xlsxrb::Elements::Cell.column_letter(:foo) }
    assert_equal("Column index must be a non-negative Integer, got :foo", err_sym.message)

    err_flt = assert_raises(ArgumentError) { Xlsxrb::Elements::Cell.column_letter(1.5) }
    assert_equal("Column index must be a non-negative Integer, got 1.5", err_flt.message)
  end

  test "cell column_index converts string digits correctly" do
    assert_equal(0, Xlsxrb::Elements::Cell.column_index("0"))
    assert_equal(25, Xlsxrb::Elements::Cell.column_index("25"))
  end

  test "cell column_index converts letters, symbols, and integers correctly" do
    # Integer pass-through
    assert_equal(0, Xlsxrb::Elements::Cell.column_index(0))
    assert_equal(26, Xlsxrb::Elements::Cell.column_index(26))
    err_neg_int = assert_raises(ArgumentError) { Xlsxrb::Elements::Cell.column_index(-1) }
    assert_equal("Column index must be >= 0, got -1", err_neg_int.message)

    # Symbol conversion
    assert_equal(0, Xlsxrb::Elements::Cell.column_index(:A))
    assert_equal(25, Xlsxrb::Elements::Cell.column_index(:Z))
    assert_equal(26, Xlsxrb::Elements::Cell.column_index(:AA))

    # Letter string conversion (both upper and lowercase)
    assert_equal(0, Xlsxrb::Elements::Cell.column_index("A"))
    assert_equal(0, Xlsxrb::Elements::Cell.column_index("a"))
    assert_equal(25, Xlsxrb::Elements::Cell.column_index("z"))
    assert_equal(26, Xlsxrb::Elements::Cell.column_index("aa"))
    assert_equal(51, Xlsxrb::Elements::Cell.column_index("AZ"))
    assert_equal(16_383, Xlsxrb::Elements::Cell.column_index("XFD"))
  end

  test "cell column_index raises ArgumentError for invalid values" do
    err_neg_str = assert_raises(ArgumentError) { Xlsxrb::Elements::Cell.column_index("-1") }
    assert_equal("Column index must be >= 0, got -1", err_neg_str.message)

    err_excl = assert_raises(ArgumentError) { Xlsxrb::Elements::Cell.column_index("!") }
    assert_equal('Invalid column letter: "!"', err_excl.message)

    err_mixed = assert_raises(ArgumentError) { Xlsxrb::Elements::Cell.column_index("A1") }
    assert_equal('Invalid column letter: "A1"', err_mixed.message)

    err_empty = assert_raises(ArgumentError) { Xlsxrb::Elements::Cell.column_index("") }
    assert_equal('Invalid column letter: ""', err_empty.message)

    err_sym = assert_raises(ArgumentError) { Xlsxrb::Elements::Cell.column_index(:bad!) }
    assert_equal("Invalid column letter: :bad!", err_sym.message)
  end

  test "cell parse_ref converts A1-style to indices" do
    assert_equal([0, 0], Xlsxrb::Elements::Cell.parse_ref("A1"))
    assert_equal([9, 1], Xlsxrb::Elements::Cell.parse_ref("B10"))
    assert_equal([0, 26], Xlsxrb::Elements::Cell.parse_ref("AA1"))
  end

  test "cell parse_ref handles lowercase, edge cases, and invalid inputs" do
    assert_nil(Xlsxrb::Elements::Cell.parse_ref(nil))
    assert_nil(Xlsxrb::Elements::Cell.parse_ref(""))
    assert_nil(Xlsxrb::Elements::Cell.parse_ref("123"))
    assert_nil(Xlsxrb::Elements::Cell.parse_ref("1A"))
    assert_nil(Xlsxrb::Elements::Cell.parse_ref("A"))
    assert_nil(Xlsxrb::Elements::Cell.parse_ref("ABC"))
    assert_nil(Xlsxrb::Elements::Cell.parse_ref("A0"))
    assert_nil(Xlsxrb::Elements::Cell.parse_ref("A-1"))
    assert_nil(Xlsxrb::Elements::Cell.parse_ref("invalid_ref"))
    assert_nil(Xlsxrb::Elements::Cell.parse_ref("A1_trailing"))

    # Lowercase reference
    assert_equal([0, 0], Xlsxrb::Elements::Cell.parse_ref("a1"))
    assert_equal([9, 1], Xlsxrb::Elements::Cell.parse_ref("b10"))
    assert_equal([99, 26], Xlsxrb::Elements::Cell.parse_ref("aa100"))
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
    row = Xlsxrb::Elements::Row.new(index: -1)
    refute(row.valid?)
    assert(row.errors.any? { |e| e.include?("index must be a non-negative Integer") })
  end

  test "row with too large index is invalid" do
    row = Xlsxrb::Elements::Row.new(index: 1_048_576)
    refute(row.valid?)
    assert(row.errors.any? { |e| e.include?("index must be < 1048576") && e.include?("max row is 1048575") })
  end

  test "row validate checks non-integer index and non-array cells" do
    errs = Xlsxrb::Elements::Row.validate("invalid", "not-an-array")
    assert_equal(["index must be a non-negative Integer (got \"invalid\")", "cells must be an Array (got String)"], errs)
    assert_equal(["index must be a non-negative Integer (got nil)"], Xlsxrb::Elements::Row.validate(nil, []))
    assert_equal(["index must be a non-negative Integer (got -1)"], Xlsxrb::Elements::Row.validate(-1, []))
    assert_equal(["index must be < 1048576 (got 1048576, max row is 1048575)"], Xlsxrb::Elements::Row.validate(1_048_576, []))
    assert_equal([], Xlsxrb::Elements::Row.validate(0, []))
    assert_equal([], Xlsxrb::Elements::Row.validate(1_048_575, []))
  end

  test "Row.valid_index? checks OOXML row bounds" do
    assert_equal(true, Xlsxrb::Elements::Row.valid_index?(0))
    assert_equal(true, Xlsxrb::Elements::Row.valid_index?(1_048_575))
    assert_equal(false, Xlsxrb::Elements::Row.valid_index?(-1))
    assert_equal(false, Xlsxrb::Elements::Row.valid_index?(1_048_576))
    assert_equal(false, Xlsxrb::Elements::Row.valid_index?("0"))
    assert_equal(false, Xlsxrb::Elements::Row.valid_index?(nil))
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

  test "row accessors, enumeration, and conversion to_a" do
    c1 = Xlsxrb::Elements::Cell.new(row_index: 2, column_index: 0, value: "hello")
    c2 = Xlsxrb::Elements::Cell.new(row_index: 2, column_index: 2, value: "world")
    row = Xlsxrb::Elements::Row.new(index: 2, cells: [c1, c2], height: 30.0, hidden: true, custom_height: true, outline_level: 1)

    assert(row.valid?)
    assert_equal(c1, row[0])
    assert_equal(c2, row[1])
    assert_nil(row[2])

    # cell_at looks up by column_index
    assert_equal(c1, row.cell_at(0))
    assert_nil(row.cell_at(1))
    assert_equal(c2, row.cell_at(2))

    # Symbol accessors
    assert_equal([c1, c2], row[:cells])
    assert_equal(2, row[:index])
    assert_in_delta(30.0, row[:height])
    assert_equal(true, row[:hidden])
    assert_equal(true, row[:custom_height])
    assert_equal(1, row[:outline_level])
    assert_equal({ height: 30.0, hidden: true, custom_height: true, outline_level: 1 }, row[:attrs])

    # Enumeration
    assert_kind_of(Enumerator, row.each)
    assert_kind_of(Enumerator, row.each_cell)
    assert_equal(%w[hello world], row.map(&:value))

    collected_cells = []
    row.each_cell { |c| collected_cells << c.value }
    assert_equal(%w[hello world], collected_cells)

    # to_a and values
    assert_equal(["hello", nil, "world"], row.to_a)
    assert_equal(["hello", nil, "world"], row.values)

    # empty row
    empty_row = Xlsxrb::Elements::Row.new(index: 0, cells: [])
    assert_equal([], empty_row.to_a)
    assert_equal([], empty_row.values)
  end

  test "row style_index and attrs with style_index" do
    row_without_style = Xlsxrb::Elements::Row.new(index: 0)
    assert_nil(row_without_style.style_index)
    assert_nil(row_without_style[:style_index])
    assert_equal({ height: nil, hidden: false, custom_height: false, outline_level: nil }, row_without_style[:attrs])

    row_with_style = Xlsxrb::Elements::Row.new(index: 1, unmapped_data: { style_index: 4 })
    assert_equal(4, row_with_style.style_index)
    assert_equal(4, row_with_style[:style_index])
    assert_equal({ height: nil, hidden: false, custom_height: false, outline_level: nil, style_index: 4 }, row_with_style[:attrs])
  end

  # --- Column ---

  test "column creates a valid column" do
    col = Xlsxrb::Elements::Column.new(index: 0, width: 15.5, hidden: true, custom_width: true, outline_level: 2)
    assert(col.valid?)
    assert_equal(0, col.index)
    assert_in_delta(15.5, col.width)
    assert_equal(true, col.hidden)
    assert_equal(true, col.custom_width)
    assert_equal(2, col.outline_level)
  end

  test "column style_index returns style from unmapped_data" do
    col1 = Xlsxrb::Elements::Column.new(index: 0)
    assert_nil(col1.style_index)

    col2 = Xlsxrb::Elements::Column.new(index: 1, unmapped_data: { style_index: 2 })
    assert_equal(2, col2.style_index)

    col3 = Xlsxrb::Elements::Column.new(index: 2, unmapped_data: { style: 5 })
    assert_equal(5, col3.style_index)
  end

  test "column with negative index is invalid" do
    col = Xlsxrb::Elements::Column.new(index: -1)
    refute(col.valid?)
    assert(col.errors.any? { |e| e.include?("index must be a non-negative Integer") && e.include?("-1") })
  end

  test "column with too large index is invalid" do
    col = Xlsxrb::Elements::Column.new(index: 16_384)
    refute(col.valid?)
    assert(col.errors.any? { |e| e.include?("index must be < 16384") && e.include?("16384") })
  end

  test "column validate checks non-integer index" do
    errs = Xlsxrb::Elements::Column.validate("bad_index")
    assert_equal(["index must be a non-negative Integer (got \"bad_index\")"], errs)
    assert_equal(["index must be a non-negative Integer (got nil)"], Xlsxrb::Elements::Column.validate(nil))
    assert_equal(["index must be a non-negative Integer (got -1)"], Xlsxrb::Elements::Column.validate(-1))
    assert_equal(["index must be < 16384 (got 16384, max column is XFD=16383)"], Xlsxrb::Elements::Column.validate(16_384))
    assert_equal([], Xlsxrb::Elements::Column.validate(0))
    assert_equal([], Xlsxrb::Elements::Column.validate(16_383))
  end

  test "Column.valid_index? checks OOXML column bounds" do
    assert_equal(true, Xlsxrb::Elements::Column.valid_index?(0))
    assert_equal(true, Xlsxrb::Elements::Column.valid_index?(16_383))
    assert_equal(false, Xlsxrb::Elements::Column.valid_index?(-1))
    assert_equal(false, Xlsxrb::Elements::Column.valid_index?(16_384))
    assert_equal(false, Xlsxrb::Elements::Column.valid_index?("0"))
    assert_equal(false, Xlsxrb::Elements::Column.valid_index?(nil))
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
    ws = Xlsxrb::Elements::Worksheet.new(name: "")
    refute(ws.valid?)
    assert(ws.errors.any? { |e| e.include?("worksheet name must be a non-empty String") })
  end

  test "worksheet validate checks name length and forbidden characters" do
    # Name longer than 31 characters
    long_name = "A" * 32
    errs_long = Xlsxrb::Elements::Worksheet.validate(long_name, [])
    assert(errs_long.any? { |e| e.include?("cannot exceed 31 characters") && e.include?("32") })

    # Exactly 31 characters is valid
    assert_empty(Xlsxrb::Elements::Worksheet.validate("A" * 31, []))

    # Forbidden characters: \ / ? * [ ]
    %w[\\ / ? * [ ]].each do |char|
      errs_char = Xlsxrb::Elements::Worksheet.validate("Sheet#{char}1", [])
      assert(errs_char.any? { |e| e.include?("cannot contain \\, /, ?, *, [, or ]") }, "Expected error for character: #{char}")
    end

    # Non-array rows
    errs_rows = Xlsxrb::Elements::Worksheet.validate("Sheet1", "invalid_rows")
    assert(errs_rows.any? { |e| e.include?("rows must be an Array") })
  end

  test "worksheet valid_name? checks length, forbidden chars, and non-strings" do
    assert_true(Xlsxrb::Elements::Worksheet.valid_name?("Sheet1"))
    assert_true(Xlsxrb::Elements::Worksheet.valid_name?("A" * 31))
    assert_false(Xlsxrb::Elements::Worksheet.valid_name?("A" * 32))
    assert_false(Xlsxrb::Elements::Worksheet.valid_name?(""))
    assert_false(Xlsxrb::Elements::Worksheet.valid_name?(nil))
    assert_false(Xlsxrb::Elements::Worksheet.valid_name?(123))
    assert_false(Xlsxrb::Elements::Worksheet.valid_name?(:Sheet1))
    %w[\\ / ? * [ ]].each do |char|
      assert_false(Xlsxrb::Elements::Worksheet.valid_name?("Sheet#{char}1"))
    end
  end

  test "worksheet update_cell creates or modifies cells in rows" do
    ws = Xlsxrb::Elements::Worksheet.new(name: "Test")

    # Update cell in brand new row
    ws = ws.update_cell("B2", value: "first_val", style_index: 3)
    assert_equal("first_val", ws.cell_value("B2"))

    # Update cell in existing row (new column)
    ws = ws.update_cell("C2", value: "second_val")
    assert_equal("second_val", ws.cell_value("C2"))

    # Overwrite existing cell in existing row
    ws = ws.update_cell("B2", value: "updated_val")
    assert_equal("updated_val", ws.cell_value("B2"))

    # Invalid ref raises ArgumentError
    assert_raises(ArgumentError) { ws.update_cell("invalid_ref", value: 1) }
  end

  test "worksheet with duplicate row indices is invalid" do
    r1 = Xlsxrb::Elements::Row.new(index: 0)
    r2 = Xlsxrb::Elements::Row.new(index: 0)
    ws = Xlsxrb::Elements::Worksheet.new(name: "Sheet1", rows: [r1, r2])
    assert(ws.errors.any? { |e| e.include?("duplicate row index") && e.include?("0") })
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

  test "coordinate access methods on worksheet (cells, cells_hash, bracket access, first_row, last_row)" do
    c1 = Xlsxrb::Elements::Cell.new(row_index: 3, column_index: 1, value: "B4")
    c2 = Xlsxrb::Elements::Cell.new(row_index: 1, column_index: 0, value: "A2")
    c3 = Xlsxrb::Elements::Cell.new(row_index: 1, column_index: 2, value: "C2")
    c4 = Xlsxrb::Elements::Cell.new(row_index: 2, column_index: 0, value: "A3")
    r1 = Xlsxrb::Elements::Row.new(index: 1, cells: [c3, c2]) # col 2 before col 0
    r2 = Xlsxrb::Elements::Row.new(index: 2, cells: [c4])
    r3 = Xlsxrb::Elements::Row.new(index: 3, cells: [c1])
    ws = Xlsxrb::Elements::Worksheet.new(name: "Sheet1", rows: [r3, r2, r1])

    # first_row and last_row
    assert_equal(1, ws.first_row.index)
    assert_equal(3, ws.last_row.index)

    # first_column and last_column (1-based)
    assert_equal(1, ws.first_column)
    assert_equal(1, ws.first_col)
    assert_equal(3, ws.last_column)
    assert_equal(3, ws.last_col)

    # computed dimension
    assert_equal("A2:C4", ws.dimension)

    # cells_hash
    hash = ws.cells_hash
    assert_equal(c2, hash["A2"])
    assert_equal(c3, hash["C2"])
    assert_equal(c4, hash["A3"])
    assert_equal(c1, hash["B4"])
    assert_nil(hash["Z99"])

    # cells strictly sorted by [row_index, column_index]
    assert_equal([c2, c3, c4, c1], ws.cells)

    # [] bracket access with String and Symbol (case insensitive)
    assert_equal(c2, ws["A2"])
    assert_equal(c2, ws["a2"])
    assert_equal(c1, ws[:B4])
    assert_equal(c1, ws[:b4])
    assert_nil(ws["Z99"])

    # cell_value edge cases
    assert_equal("A2", ws.cell_value("A2"))
    assert_equal("A2", ws.cell_value("a2"))
    assert_nil(ws.cell_value("A4"))
    assert_nil(ws.cell_value("B2"))
    assert_nil(ws.cell_value("invalid_ref"))

    # Empty worksheet
    empty_ws = Xlsxrb::Elements::Worksheet.new(name: "Empty", rows: [])
    assert_nil(empty_ws.first_row)
    assert_nil(empty_ws.last_row)
    assert_nil(empty_ws.first_column)
    assert_nil(empty_ws.first_col)
    assert_nil(empty_ws.last_column)
    assert_nil(empty_ws.last_col)
    assert_nil(empty_ws.dimension)
    assert_equal({}, empty_ws.cells_hash)
    assert_equal([], empty_ws.cells)
    assert_nil(empty_ws["A1"])
  end

  test "worksheet dimension property, with, and deconstruct_keys" do
    ws = Xlsxrb::Elements::Worksheet.new(name: "Sheet1", dimension: "A1:Z100")
    assert_equal("A1:Z100", ws.dimension)

    # with replaces dimension
    ws2 = ws.with(dimension: "B2:D10")
    assert_equal("B2:D10", ws2.dimension)
    assert_equal("A1:Z100", ws.dimension)

    # deconstruct_keys includes dimension
    assert_equal("A1:Z100", ws.deconstruct_keys(nil)[:dimension])

    # single cell dimension computation
    single_cell = Xlsxrb::Elements::Cell.new(row_index: 2, column_index: 2, value: 42)
    single_row = Xlsxrb::Elements::Row.new(index: 2, cells: [single_cell])
    single_ws = Xlsxrb::Elements::Worksheet.new(name: "Single", rows: [single_row])
    assert_equal("C3", single_ws.dimension)
  end

  test "worksheet handles first-class conditional_formatting and data_validations" do
    ws = Xlsxrb::Elements::Worksheet.new(name: "Sheet1")
    assert_equal([], ws.conditional_formatting)
    assert_equal([], ws.conditional_formats)
    assert_equal([], ws.data_validations)
    assert_predicate(ws.conditional_formatting, :frozen?)
    assert_predicate(ws.data_validations, :frozen?)

    cf_rule = [{ sqref: "A1:A10", type: "cellIs", operator: "greaterThan", formula: "100" }]
    dv_rule = [{ sqref: "B1:B10", type: "whole", operator: "between", formula1: "1", formula2: "10" }]

    ws_with_rules = Xlsxrb::Elements::Worksheet.new(
      name: "Sheet1",
      conditional_formatting: cf_rule,
      data_validations: dv_rule
    )
    assert_equal(cf_rule, ws_with_rules.conditional_formatting)
    assert_equal(cf_rule, ws_with_rules.conditional_formats)
    assert_equal(dv_rule, ws_with_rules.data_validations)

    # with
    new_cf = [{ sqref: "C1:C10", type: "containsText" }]
    ws_updated = ws_with_rules.with(conditional_formatting: new_cf)
    assert_equal(new_cf, ws_updated.conditional_formatting)
    assert_equal(dv_rule, ws_updated.data_validations)

    ws_updated_alias = ws_with_rules.with(conditional_formats: new_cf)
    assert_equal(new_cf, ws_updated_alias.conditional_formatting)

    new_dv = [{ sqref: "D1:D10", type: "list", formula1: '"A,B,C"' }]
    ws_updated_dv = ws_with_rules.with(data_validations: new_dv)
    assert_equal(cf_rule, ws_updated_dv.conditional_formatting)
    assert_equal(new_dv, ws_updated_dv.data_validations)

    # deconstruct_keys
    keys = ws_with_rules.deconstruct_keys(nil)
    assert_equal(cf_rule, keys[:conditional_formatting])
    assert_equal(cf_rule, keys[:conditional_formats])
    assert_equal(dv_rule, keys[:data_validations])

    # == and hash
    same_ws = Xlsxrb::Elements::Worksheet.new(name: "Sheet1", conditional_formatting: cf_rule, data_validations: dv_rule)
    diff_cf_ws = Xlsxrb::Elements::Worksheet.new(name: "Sheet1", conditional_formatting: new_cf, data_validations: dv_rule)
    diff_dv_ws = Xlsxrb::Elements::Worksheet.new(name: "Sheet1", conditional_formatting: cf_rule, data_validations: new_dv)

    assert_equal(ws_with_rules, same_ws)
    assert_equal(ws_with_rules.hash, same_ws.hash)
    refute_equal(ws_with_rules, diff_cf_ws)
    refute_equal(ws_with_rules, diff_dv_ws)
  end

  test "worksheet handles visibility state and hidden/visible predicates" do
    ws_default = Xlsxrb::Elements::Worksheet.new(name: "VisibleSheet")
    assert_equal(:visible, ws_default.state)
    assert_true(ws_default.visible?)
    assert_false(ws_default.hidden?)

    ws_hidden = Xlsxrb::Elements::Worksheet.new(name: "HiddenSheet", state: :hidden)
    assert_equal(:hidden, ws_hidden.state)
    assert_false(ws_hidden.visible?)
    assert_true(ws_hidden.hidden?)

    ws_very_hidden = Xlsxrb::Elements::Worksheet.new(name: "VeryHiddenSheet", state: :very_hidden)
    assert_equal(:very_hidden, ws_very_hidden.state)
    assert_false(ws_very_hidden.visible?)
    assert_true(ws_very_hidden.hidden?)

    # with
    ws_toggled = ws_default.with(state: :hidden)
    assert_equal(:hidden, ws_toggled.state)
    assert_true(ws_toggled.hidden?)

    # deconstruct_keys
    assert_equal(:hidden, ws_hidden.deconstruct_keys(nil)[:state])

    # == and hash
    refute_equal(ws_default, ws_default.with(state: :hidden))
    same_hidden = Xlsxrb::Elements::Worksheet.new(name: "HiddenSheet", state: :hidden)
    assert_equal(ws_hidden, same_hidden)
    assert_equal(ws_hidden.hash, same_hidden.hash)
  end

  # --- Workbook ---

  test "workbook creates a valid workbook" do
    ws = Xlsxrb::Elements::Worksheet.new(name: "Sheet1")
    wb = Xlsxrb::Elements::Workbook.new(sheets: [ws])
    assert(wb.valid?)
  end

  test "workbook with no sheets is invalid" do
    wb = Xlsxrb::Elements::Workbook.new(sheets: [])
    refute(wb.valid?)
    assert_include(wb.errors, "workbook must have at least one sheet")
  end

  test "workbook with duplicate sheet names is invalid" do
    ws1 = Xlsxrb::Elements::Worksheet.new(name: "Sheet1")
    ws2 = Xlsxrb::Elements::Worksheet.new(name: "Sheet1")
    wb = Xlsxrb::Elements::Workbook.new(sheets: [ws1, ws2])
    refute(wb.valid?)
    assert(wb.errors.any? { |e| e.include?("duplicate sheet name") && e.include?("Sheet1") })
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

  test "workbook defined_names and defined_name lookup" do
    ws1 = Xlsxrb::Elements::Worksheet.new(name: "Summary")
    ws2 = Xlsxrb::Elements::Worksheet.new(name: "Details")
    dns = [
      { name: "TotalRevenue", value: "Summary!$B$10", local_sheet_id: nil },
      { name: "LocalRate", value: "Details!$A$1", local_sheet_id: 1 }
    ]
    wb = Xlsxrb::Elements::Workbook.new(sheets: [ws1, ws2], defined_names: dns)

    assert_equal(dns, wb.defined_names)
    assert_equal("Summary!$B$10", wb.defined_name("TotalRevenue")[:value])
    assert_nil(wb.defined_name("NonExistent"))

    # Sheet-specific lookup
    assert_equal("Details!$A$1", wb.defined_name("LocalRate", sheet: "Details")[:value])
    assert_equal("Details!$A$1", wb.defined_name("LocalRate", sheet: 1)[:value])
    assert_nil(wb.defined_name("LocalRate", sheet: "Summary"))

    # Fallback to unmapped_data if defined_names not passed explicitly
    wb_unmapped = Xlsxrb::Elements::Workbook.new(sheets: [ws1], unmapped_data: { defined_names: dns })
    assert_equal(dns, wb_unmapped.defined_names)

    # load preserves defined_names
    loaded = wb.load
    assert_equal(dns, loaded.defined_names)
  end

  test "workbook validate checks non-array sheets" do
    errs = Xlsxrb::Elements::Workbook.validate("invalid_sheets")
    assert(errs.any? { |e| e.include?("sheets must be an Array") })
  end

  test "workbook sheet lookups, bracket access, and enumeration" do
    ws1 = Xlsxrb::Elements::Worksheet.new(name: "Sales")
    ws2 = Xlsxrb::Elements::Worksheet.new(name: "Expenses")
    wb = Xlsxrb::Elements::Workbook.new(sheets: [ws1, ws2])

    # index lookup
    assert_equal(ws1, wb.sheet(0))
    assert_equal(ws2, wb.sheet(1))
    assert_nil(wb.sheet(2))

    # name lookup and bracket alias
    assert_equal(ws1, wb.sheet("Sales"))
    assert_equal(ws2, wb["Expenses"])
    assert_nil(wb.sheet("Missing"))
    assert_nil(wb[:symbol_not_supported])

    # Enumerable each and each_sheet
    assert_equal(%w[Sales Expenses], wb.map(&:name))
    collected = []
    wb.each_sheet { |s| collected << s.name }
    assert_equal(%w[Sales Expenses], collected)
  end

  test "workbook update_sheet modifies matching sheet and validates block" do
    ws1 = Xlsxrb::Elements::Worksheet.new(name: "OldName")
    wb = Xlsxrb::Elements::Workbook.new(sheets: [ws1])

    # Successful update
    updated_wb = wb.update_sheet("OldName") do |sheet|
      sheet.with(name: "NewName")
    end
    assert_equal("NewName", updated_wb.sheet(0).name)
    assert_equal("OldName", wb.sheet(0).name) # immutability

    # Block missing
    assert_raises(ArgumentError) { wb.update_sheet("OldName") }

    # Sheet not found
    assert_raises(ArgumentError) { wb.update_sheet("NonExistent") { |s| s } }

    # Invalid return type from block
    expected_errors = [TypeError]
    expected_errors << RBS::Test::Tester::TypeError if defined?(RBS::Test::Tester::TypeError)
    assert_raises(*expected_errors) { wb.update_sheet("OldName") { "not_a_worksheet" } }
  end

  # --- Formula helper ---

  test "Xlsxrb.formula creates a Formula without cached value" do
    f = Xlsxrb.formula("SUM(A1:A10)")
    assert_instance_of(Xlsxrb::Elements::Formula, f)
    assert_equal("SUM(A1:A10)", f.expression)
    assert_nil(f.cached_value)
    assert_equal(true, f.calculate_always)
  end

  test "Xlsxrb.formula creates a Formula with cached value" do
    f = Xlsxrb.formula("SUM(A1:A10)", cached_value: "55")
    assert_equal("SUM(A1:A10)", f.expression)
    assert_equal("55", f.cached_value)
    assert_nil(f.calculate_always)
  end

  # --- CellError & RichText ---

  test "cell_error valid error codes, to_s, equality, and validation" do
    Xlsxrb::Elements::VALID_ERROR_CODES.each do |code|
      err = Xlsxrb::Elements::CellError.new(code: code)
      assert_equal(code, err.code)
      assert_equal(code, err.to_s)
    end

    err1 = Xlsxrb::Elements::CellError.new(code: "#REF!")
    err2 = Xlsxrb::Elements::CellError.new(code: "#REF!")
    err3 = Xlsxrb::Elements::CellError.new(code: "#N/A")
    assert_equal(err1, err2)
    refute_equal(err1, err3)

    assert_raises(ArgumentError) { Xlsxrb::Elements::CellError.new(code: "#INVALID!") }
    assert_raises(ArgumentError) { Xlsxrb::Elements::CellError.new(code: "") }
  end

  test "rich_text to_s concatenation and runs handling" do
    rt = Xlsxrb::Elements::RichText.new(runs: [
                                          { text: "Hello ", font: { bold: true } },
                                          { text: "World", font: { italic: true } }
                                        ])
    assert_equal("Hello World", rt.to_s)

    empty_rt = Xlsxrb::Elements::RichText.new(runs: [])
    assert_equal("", empty_rt.to_s)
  end

  test "rich_text_run initialization, attributes, to_s, and bracket access" do
    run = Xlsxrb::Elements::RichTextRun.new(text: "Hello", font: { bold: true })
    assert_equal("Hello", run.text)
    assert_equal({ bold: true }, run.font)
    assert_equal("Hello", run[:text])
    assert_equal({ bold: true }, run[:font])
    assert_nil(run[:unknown])
    assert_equal("Hello", run.to_s)

    # empty run
    empty_run = Xlsxrb::Elements::RichTextRun.new
    assert_equal("", empty_run.text)
    assert_nil(empty_run.font)

    # combined in RichText
    rt = Xlsxrb::Elements::RichText.new(runs: [run, { text: " World" }])
    assert_equal("Hello World", rt.to_s)

    # Xlsxrb.rich_text helper accepting RichTextRun
    rt2 = Xlsxrb.rich_text(run, { text: " World" })
    assert_equal("Hello World", rt2.to_s)
  end

  # --- Error message quality ---

  test "cell error message includes actual value for unsupported type" do
    cell = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: [1, 2, 3])
    err = cell.errors.find { |e| e.include?("unsupported value type") }
    assert_not_nil(err, "Expected error about unsupported value type")
    assert_match(/Array/, err)
    assert_match(/supported types/, err)
  end

  test "cell error message includes actual row_index value" do
    cell = Xlsxrb::Elements::Cell.new(row_index: "bad", column_index: 0)
    err = cell.errors.find { |e| e.include?("row_index") }
    assert_not_nil(err)
    assert_match(/"bad"/, err)
  end

  test "cell error message includes actual column_index value for out of range" do
    cell = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 20_000)
    err = cell.errors.find { |e| e.include?("column_index must be < 16384") }
    assert_not_nil(err)
    assert_match(/20000/, err)
    assert_match(/XFD/, err)
  end

  test "row error message includes actual index value" do
    row = Xlsxrb::Elements::Row.new(index: -5)
    err = row.errors.find { |e| e.include?("index must be a non-negative Integer") }
    assert_not_nil(err)
    assert_match(/-5/, err)
  end

  test "column error message includes actual index for out of range" do
    col = Xlsxrb::Elements::Column.new(index: 99_999)
    err = col.errors.find { |e| e.include?("index must be < 16384") }
    assert_not_nil(err)
    assert_match(/99999/, err)
  end

  test "worksheet error message includes actual name value" do
    ws = Xlsxrb::Elements::Worksheet.new(name: nil)
    err = ws.errors.find { |e| e.include?("worksheet name") }
    assert_not_nil(err)
    assert_match(/nil/, err)
  end

  test "worksheet error message shows which row indices are duplicated" do
    r1 = Xlsxrb::Elements::Row.new(index: 3)
    r2 = Xlsxrb::Elements::Row.new(index: 3)
    ws = Xlsxrb::Elements::Worksheet.new(name: "S", rows: [r1, r2])
    err = ws.errors.find { |e| e.include?("duplicate row index") }
    assert_not_nil(err)
    assert_match(/3/, err)
    assert_match(/unique/, err)
  end

  test "workbook error message shows which sheet names are duplicated" do
    ws1 = Xlsxrb::Elements::Worksheet.new(name: "Sales")
    ws2 = Xlsxrb::Elements::Worksheet.new(name: "Sales")
    wb = Xlsxrb::Elements::Workbook.new(sheets: [ws1, ws2])
    err = wb.errors.find { |e| e.include?("duplicate sheet name") }
    assert_not_nil(err)
    assert_match(/Sales/, err)
    assert_match(/unique/, err)
  end

  test "build_row_from_raw includes source location in error context" do
    raw_cell = {
      ref: "A1",
      type: "n",
      style_index: 0,
      value: [1, 2, 3],
      source: { part: "xl/worksheets/sheet1.xml", row: 0, cell: "A1" }
    }
    raw_row = {
      index: 0,
      cells: [raw_cell],
      attrs: {},
      source: { part: "xl/worksheets/sheet1.xml", row: 0 }
    }

    row = Xlsxrb.send(:build_row_from_raw, raw_row)
    assert_equal(1, row.cells.size)
    cell = row.cells.first
    assert_false(cell.valid?)
    assert_match(%r{at xl/worksheets/sheet1.xml row 1 cell A1}, cell.errors.first)
  end

  test "cell raw_value, format_code, and formatted_value" do
    c_empty = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: nil)
    assert_nil(c_empty.formatted_value)
    assert_nil(c_empty.raw_value)
    assert_nil(c_empty.format_code)
    assert_nil(c_empty[:raw_value])
    assert_nil(c_empty[:format_code])
    assert_nil(c_empty[:formatted_value])

    c_num = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: 1042, raw_value: "1042", format_code: "#,##0.00")
    assert_equal("1042", c_num.raw_value)
    assert_equal("#,##0.00", c_num.format_code)
    assert_equal("1,042.00", c_num.formatted_value)
    assert_equal("1,042.00", c_num[:formatted_value])
    assert_equal("1042", c_num[:raw_value])
    assert_equal("#,##0.00", c_num[:format_code])
    assert_equal("$1,042.00", c_num.formatted_value("\"$\"#,##0.00"))

    c_date = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 1, value: Date.new(2014, 6, 1), format_code: "yyyy-mm-dd")
    assert_equal("2014-06-01", c_date.formatted_value)
    assert_equal("06-01-14", c_date.formatted_value("mm-dd-yy"))

    c_bool = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 2, value: true)
    assert_equal("TRUE", c_bool.formatted_value)

    c_err = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 3, value: Xlsxrb::Error::DIV0)
    assert_equal("#DIV/0!", c_err.formatted_value)
  end

  test "row raw_values, formatted_values, and formulas" do
    c1 = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: 1042, raw_value: "1042", format_code: "#,##0")
    c2 = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 2, value: 50, raw_value: "50", formula: "A1*2")
    row = Xlsxrb::Elements::Row.new(index: 0, cells: [c1, c2])

    assert_equal([1042, nil, 50], row.values)
    assert_equal(["1042", nil, "50"], row.raw_values)
    assert_equal(["1,042", nil, "50"], row.formatted_values)
    assert_equal([nil, nil, "A1*2"], row.formulas)
    assert_equal(row.values, row[:values])
    assert_equal(row.raw_values, row[:raw_values])
    assert_equal(row.formatted_values, row[:formatted_values])
    assert_equal(row.formulas, row[:formulas])
  end

  test "worksheet and workbook cell and formatted_value lookups" do
    c1 = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: 20.51, format_code: "\"$\"#,##0.00")
    row = Xlsxrb::Elements::Row.new(index: 0, cells: [c1])
    ws = Xlsxrb::Elements::Worksheet.new(name: "Data", rows: [row])
    wb = Xlsxrb::Elements::Workbook.new(sheets: [ws])

    assert_equal(c1, ws.cell(0, 0))
    assert_equal(c1, ws.cell("A1"))
    assert_nil(ws.cell(1, 0))
    assert_nil(ws.cell("B2"))
    assert_equal("$20.51", ws.formatted_value(0, 0))
    assert_equal("$20.51", ws.formatted_value("A1"))
    assert_nil(ws.formatted_value(1, 0))

    assert_equal("$20.51", wb.formatted_value(0, 0, 0))
    assert_equal("$20.51", wb.formatted_value("Data", "A1"))
    assert_nil(wb.formatted_value("NonExistent", "A1"))
  end
end
