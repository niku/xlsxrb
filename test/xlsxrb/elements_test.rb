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
  cover Xlsxrb::Elements::Image
  cover Xlsxrb::Elements::Styles
  cover Xlsxrb::Elements::PivotTable
  cover Xlsxrb::Elements::TableColumn
  cover Xlsxrb::Elements::ThreadedComment
  cover Xlsxrb::Elements::Comment
  cover Xlsxrb::Elements::FilterColumn
  # --- Cell ---

  test "cell creates a valid cell" do
    cell = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "hello")
    assert(cell.valid?)
    assert_predicate(cell, :frozen?)
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

    recorded_msg = nil
    recorded_uplevel = nil
    c_unfrozen = cell.dup
    c_unfrozen.define_singleton_method(:warn) do |msg = nil, uplevel: nil|
      recorded_msg = msg
      recorded_uplevel = uplevel
    end
    assert_equal("123.45", c_unfrozen.content)
    assert_equal("[DEPRECATION] `content` is deprecated; use `value` instead.", recorded_msg)
    assert_equal(1, recorded_uplevel)
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
    assert_equal([1, 1, 10, Xlsxrb::Elements::Formula.new(expression: "A1"), 1, {}, [], nil, nil, nil, nil, []], c1.deconstruct)
    assert_equal(1, c1.deconstruct_keys(%i[row_index value])[:row_index])
    assert_equal(10, c1.deconstruct_keys(%i[row_index value])[:value])
    assert_equal({ ref: "B2", row_index: 1, column_index: 1, value: 10, formula: Xlsxrb::Elements::Formula.new(expression: "A1"), style_index: 1, format_code: nil, hyperlink: nil, comment: nil, threaded_comments: [] }, c1.deconstruct_keys(nil))
    case c1
    in { ref: "B2", value: 10 }
      assert_true(true)
    else
      flunk "pattern matching failed on Cell"
    end
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
    assert_equal("Search", cell_link.tooltip)
    assert_instance_of(Xlsxrb::Elements::Hyperlink, cell_link.hyperlink)
    assert_equal(Xlsxrb::Elements::Hyperlink.new(url: "https://google.com", tooltip: "Search"), cell_link.hyperlink)
    assert_equal({ url: "https://google.com", tooltip: "Search" }, cell_link.hyperlink)
    assert_equal({ url: "https://google.com", tooltip: "Search" }, cell_link[:hyperlink])
    assert_equal("https://google.com", cell_link[:url])
    assert_equal("Search", cell_link[:tooltip])

    cell_loc = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, hyperlink: { location: "Sheet2!A1" })
    assert_true(cell_loc.link?)
    assert_equal("Sheet2!A1", cell_loc.url)
    assert_nil(cell_loc.tooltip)

    cell_str = cell_no_link.with(hyperlink: "https://ruby-lang.org")
    assert_true(cell_str.link?)
    assert_equal("https://ruby-lang.org", cell_str.url)
    assert_nil(cell_str.tooltip)

    # First-class Elements::Hyperlink and Xlsxrb.hyperlink
    hl = Xlsxrb.hyperlink("https://example.com", display: "Example Site", tooltip: "Visit website", location: "Sheet1!A1")
    assert_instance_of(Xlsxrb::Elements::Hyperlink, hl)
    assert_equal("https://example.com", hl.url)
    assert_equal("Example Site", hl.display)
    assert_equal("Visit website", hl.tooltip)
    assert_equal("Sheet1!A1", hl.location)
    assert_equal("https://example.com", hl[:url])
    assert_equal("Visit website", hl[:tooltip])
    assert_equal("https://example.com", hl.to_s)
    assert_equal({ url: "https://example.com", display: "Example Site", tooltip: "Visit website", location: "Sheet1!A1" }, hl.to_h)

    cell_hl = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, hyperlink: hl)
    assert_true(cell_hl.link?)
    assert_equal("https://example.com", cell_hl.url)
    assert_equal("Visit website", cell_hl.tooltip)
    assert_equal("https://example.com", cell_hl[:url])
    assert_equal("Visit website", cell_hl[:tooltip])

    assert_false(Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, hyperlink: "").link?)
    assert_false(Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, hyperlink: {}).link?)
    assert_nil(Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, hyperlink: "").url)
    assert_nil(Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, hyperlink: "").tooltip)
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
    expected_comment = Xlsxrb::Elements::Comment.new(author: "Author", text: "Important Note")
    assert_equal(expected_comment, cell_comment.comment)
    assert_equal(expected_comment, cell_comment[:comment])
    assert_equal("Author", cell_comment.comment.author)
    assert_equal("Important Note", cell_comment.comment.text)

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

    wrapped_links = links.transform_values { |h| Xlsxrb::Elements::Hyperlink.new(**h) }
    assert_equal(wrapped_links, ws.hyperlinks)
    assert_equal("https://example.com", ws.hyperlink("A1").url)
    assert_equal("https://example.com", ws.hyperlink("a1").url)
    assert_equal("https://example.com", ws.hyperlink(:A1).url)
    assert_equal("https://example.com", ws.hyperlink(0, 0).url)
    assert_equal("https://example.com", ws.hyperlink(nil, 0).url)
    assert_equal("https://example.com", ws.hyperlink(0, "A").url)
    assert_equal("https://example.com", ws.hyperlink(0, :A).url)
    assert_equal("Sheet2!A1", ws.hyperlink("B2").location)
    assert_equal("Sheet2!A1", ws.hyperlink(1, 1).location)
    assert_equal("Sheet2!A1", ws.hyperlink(1, "B").location)
    assert_equal("Sheet2!A1", ws.hyperlink(1, :B).location)
    assert_nil(ws.hyperlink("C3"))
    assert_nil(ws.hyperlink(2, 2))
    assert_nil(ws.hyperlink(2, "C"))
    assert_nil(ws.hyperlink(2, :C))

    updated = ws.update_cell("C3", value: "Test", hyperlink: { url: "https://updated.com" })
    assert_equal("https://updated.com", updated.hyperlink("C3").url)
    assert_true(updated["C3"].link?)
    assert_equal("https://updated.com", updated["C3"].url)
  end

  test "worksheet comments and comment coordinate lookup" do
    comments = [{ ref: "A1", author: "Tester", text: "Hello" }, { ref: "B2", author: "Reviewer", text: "Check this" }]
    ws = Xlsxrb::Elements::Worksheet.new(name: "Sheet1", comments: comments)

    wrapped = comments.map { |c| Xlsxrb::Elements::Comment.new(**c) }
    assert_equal(wrapped, ws.comments)
    assert_equal(wrapped[0], ws.comment("A1"))
    assert_equal("Hello", ws.comment("A1").text)
    assert_equal("Hello", ws.comment("a1").text)
    assert_equal("Hello", ws.comment(:A1).text)
    assert_equal("Hello", ws.comment(0, 0).text)
    assert_equal("Hello", ws.comment(nil, 0).text)
    assert_equal("Hello", ws.comment(0, "A").text)
    assert_equal("Hello", ws.comment(0, :A).text)
    assert_equal(wrapped[1], ws.comment("B2"))
    assert_equal("Check this", ws.comment(1, 1).text)
    assert_equal("Check this", ws.comment(1, "B").text)
    assert_equal("Check this", ws.comment(1, :B).text)
    assert_nil(ws.comment("C3"))
    assert_nil(ws.comment(2, 2))
    assert_nil(ws.comment(2, "C"))
    assert_nil(ws.comment(2, :C))

    updated = ws.update_cell("C3", value: "Test", comment: { author: "Me", text: "New comment" })
    assert_equal("New comment", updated.comment("C3").text)
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
    assert_equal(Xlsxrb::Elements::Formula.new(expression: "SUM(A2:A5)"), cell.formula)
    assert_equal("SUM(A2:A5)", cell.formula_expression)
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
    row = Xlsxrb::Elements::Row.new(index: 2, cells: [c1, c2], height: 30.0, hidden: true, custom_height: true, collapsed: true, outline_level: 1)

    assert(row.valid?)
    assert_equal(c1, row[0])
    assert_nil(row[1])
    assert_equal(c2, row[2])
    assert_equal(c1, row["A"])
    assert_nil(row["B"])
    assert_equal(c2, row["C"])

    # cell_at looks up by column_index
    assert_equal(c1, row.cell_at(0))
    assert_nil(row.cell_at(1))
    assert_equal(c2, row.cell_at(2))

    # Symbol accessors & predicates
    assert_equal([c1, c2], row[:cells])
    assert_equal(2, row[:index])
    assert_in_delta(30.0, row[:height])
    assert_equal(true, row[:hidden])
    assert_equal(true, row.hidden)
    assert_equal(true, row.hidden?)
    assert_equal(true, row[:hidden?])
    assert_equal(true, row[:custom_height])
    assert_equal(true, row.custom_height)
    assert_equal(true, row.custom_height?)
    assert_equal(true, row[:custom_height?])
    assert_equal(true, row[:collapsed])
    assert_equal(true, row.collapsed)
    assert_equal(true, row.collapsed?)
    assert_equal(true, row[:collapsed?])
    assert_equal(1, row[:outline_level])
    assert_equal(1, row.outline_level)
    assert_equal({ height: 30.0, hidden: true, custom_height: true, collapsed: true, outline_level: 1 }, row[:attrs])

    expected_attrs = {
      index: 2,
      height: 30.0,
      hidden: true,
      custom_height: true,
      collapsed: true,
      outline_level: 1
    }
    assert_equal(expected_attrs, row.attributes)
    assert_equal(expected_attrs, row[:attributes])
    assert(row.attributes.frozen?)

    # Defaults and false predicates
    default_row = Xlsxrb::Elements::Row.new(index: 0)
    assert_equal(false, default_row.hidden)
    assert_equal(false, default_row.hidden?)
    assert_equal(false, default_row[:hidden?])
    assert_equal(false, default_row.custom_height)
    assert_equal(false, default_row.custom_height?)
    assert_equal(false, default_row[:custom_height?])
    assert_equal(false, default_row.collapsed)
    assert_equal(false, default_row.collapsed?)
    assert_equal(false, default_row[:collapsed?])
    assert_equal({ index: 0, height: nil, hidden: false, custom_height: false, collapsed: false, outline_level: nil }, default_row.attributes)
    assert(default_row.attributes.frozen?)

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
    assert_equal({ height: nil, hidden: false, custom_height: false, collapsed: false, outline_level: nil }, row_without_style[:attrs])

    row_with_style = Xlsxrb::Elements::Row.new(index: 1, unmapped_data: { style_index: 4 })
    assert_equal(4, row_with_style.style_index)
    assert_equal(4, row_with_style[:style_index])
    assert_equal({ height: nil, hidden: false, custom_height: false, collapsed: false, outline_level: nil, style_index: 4 }, row_with_style[:attrs])
    assert_equal({ index: 1, height: nil, hidden: false, custom_height: false, collapsed: false, outline_level: nil, style_index: 4 }, row_with_style.attributes)
    assert(row_with_style.attributes.frozen?)
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

  test "column style_index returns style from keyword or unmapped_data" do
    col1 = Xlsxrb::Elements::Column.new(index: 0)
    assert_nil(col1.style_index)

    col_kw = Xlsxrb::Elements::Column.new(index: 0, style_index: 7)
    assert_equal(7, col_kw.style_index)

    col2 = Xlsxrb::Elements::Column.new(index: 1, unmapped_data: { style_index: 2 })
    assert_equal(2, col2.style_index)

    col3 = Xlsxrb::Elements::Column.new(index: 2, unmapped_data: { style: 5 })
    assert_equal(5, col3.style_index)

    col_both = Xlsxrb::Elements::Column.new(index: 3, style_index: 9, unmapped_data: { style_index: 2 })
    assert_equal(9, col_both.style_index)
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

    # Forbidden characters: \ / ? * [ ] :
    %w[\\ / ? * [ ] :].each do |char|
      errs_char = Xlsxrb::Elements::Worksheet.validate("Sheet#{char}1", [])
      assert(errs_char.any? { |e| e.include?("cannot contain \\, /, ?, *, [, ], or :") }, "Expected error for character: #{char}")
    end

    # Apostrophe at start or end
    errs_start_quote = Xlsxrb::Elements::Worksheet.validate("'Sheet1", [])
    assert(errs_start_quote.any? { |e| e.include?("cannot start or end with an apostrophe") })
    errs_end_quote = Xlsxrb::Elements::Worksheet.validate("Sheet1'", [])
    assert(errs_end_quote.any? { |e| e.include?("cannot start or end with an apostrophe") })

    # Reserved History name
    errs_history = Xlsxrb::Elements::Worksheet.validate("History", [])
    assert(errs_history.any? { |e| e.include?("'History' is reserved") })
    errs_history_lower = Xlsxrb::Elements::Worksheet.validate("history", [])
    assert(errs_history_lower.any? { |e| e.include?("'History' is reserved") })

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
    %w[\\ / ? * [ ] :].each do |char|
      assert_false(Xlsxrb::Elements::Worksheet.valid_name?("Sheet#{char}1"))
    end
    assert_false(Xlsxrb::Elements::Worksheet.valid_name?("'Sheet1"))
    assert_false(Xlsxrb::Elements::Worksheet.valid_name?("Sheet1'"))
    assert_false(Xlsxrb::Elements::Worksheet.valid_name?("History"))
    assert_false(Xlsxrb::Elements::Worksheet.valid_name?("history"))
    assert_false(Xlsxrb::Elements::Worksheet.valid_name?("Sheet\x001"))
    assert_true(Xlsxrb::Elements::Worksheet.valid_name?("John's Sheet"))
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
    assert_equal(0..2, ws.column_range)
    assert_equal(1..3, ws.row_range)

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

    # [] bracket access with 0-based coordinate indices [row, col]
    assert_equal(c2, ws[1, 0])
    assert_equal(c2, ws[1, "A"])
    assert_equal(c1, ws[3, 1])
    assert_equal(c1, ws[3, "B"])
    assert_nil(ws[1, 1])
    assert_nil(ws[99, 0])

    # row_range and column_range
    assert_equal(1..3, ws.row_range)
    assert_equal(0..2, ws.column_range)
    assert_equal({ min_row: 1, max_row: 3, min_column: 0, max_column: 2 }, ws.dimension_bounds)
    assert_true(ws.dimension_bounds.frozen?)

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
    assert_nil(empty_ws.row_range)
    assert_nil(empty_ws.first_column)
    assert_nil(empty_ws.first_col)
    assert_nil(empty_ws.last_column)
    assert_nil(empty_ws.last_col)
    assert_nil(empty_ws.column_range)
    assert_nil(empty_ws.dimension)
    assert_equal({}, empty_ws.cells_hash)
    assert_equal([], empty_ws.cells)
    assert_nil(empty_ws["A1"])
    assert_nil(empty_ws[0, 0])
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
    assert_equal({ min_row: 0, max_row: 99, min_column: 0, max_column: 25 }, ws.dimension_bounds)
    assert_equal({ min_row: 1, max_row: 9, min_column: 1, max_column: 3 }, ws2.dimension_bounds)
    assert_equal({ min_row: 2, max_row: 2, min_column: 2, max_column: 2 }, single_ws.dimension_bounds)
    empty_ws = Xlsxrb::Elements::Worksheet.new(name: "Empty")
    assert_nil(empty_ws.dimension_bounds)

    invalid_ws1 = Xlsxrb::Elements::Worksheet.new(name: "Invalid1", dimension: "INVALID")
    assert_nil(invalid_ws1.dimension_bounds)
    invalid_ws2 = Xlsxrb::Elements::Worksheet.new(name: "Invalid2", dimension: "A1:INVALID")
    assert_nil(invalid_ws2.dimension_bounds)
    invalid_ws3 = Xlsxrb::Elements::Worksheet.new(name: "Invalid3", dimension: "INVALID:B2")
    assert_nil(invalid_ws3.dimension_bounds)
    invalid_ws4 = Xlsxrb::Elements::Worksheet.new(name: "Invalid4", dimension: "A1:B2:C3")
    assert_nil(invalid_ws4.dimension_bounds)
    invalid_ws5 = Xlsxrb::Elements::Worksheet.new(name: "Invalid5", dimension: "")
    assert_nil(invalid_ws5.dimension_bounds)

    no_dim_obj = Object.new.extend(Xlsxrb::Elements::CoordinateAccess)
    assert_nil(no_dim_obj.dimension_bounds)
  end

  test "coordinate access with custom non-enumerable host" do
    custom_class = Class.new do
      include Xlsxrb::Elements::CoordinateAccess

      attr_accessor :rows, :dimension

      def initialize(rows, dimension = nil)
        @rows = rows
        @dimension = dimension
      end

      def cells
        raise "cells should not be called by coordinate queries"
      end
    end

    c_a = Xlsxrb::Elements::Cell.new(row_index: 1, column_index: 0, value: 1)
    c_b = Xlsxrb::Elements::Cell.new(row_index: 3, column_index: 2, value: 2)
    r1 = Xlsxrb::Elements::Row.new(index: 1, cells: [c_a])
    r2 = Xlsxrb::Elements::Row.new(index: 3, cells: [c_b])
    host = custom_class.new([r2, r1], "A2:C4")

    assert_equal(r1, host.row_at(1))
    assert_nil(host.row_at(99))
    assert_equal(r1, host.first_row)
    assert_equal(r2, host.last_row)
    assert_equal(1..3, host.row_range)
    assert_equal(0..2, host.column_range)
    assert_equal(0..2, custom_class.new([r1, r2], "A2:C4").column_range)
    assert_equal({ min_row: 1, max_row: 3, min_column: 0, max_column: 2 }, host.dimension_bounds)

    # cell_value deprecated method
    assert_equal(1, host.cell_value("A2"))
    assert_nil(host.cell_value("invalid"))
    assert_nil(host.cell_value("Z99"))
    assert_nil(host.cell_value("B2"))

    # first_column / last_column / aliases
    assert_equal(1, host.first_column)
    assert_equal(1, host.first_col)
    assert_equal(3, host.last_column)
    assert_equal(3, host.last_col)

    # row_range when first_row or last_row is nil
    empty_host = custom_class.new([])
    assert_nil(empty_host.first_row)
    assert_nil(empty_host.last_row)
    assert_nil(empty_host.row_range)
    assert_nil(empty_host.first_column)
    assert_nil(empty_host.last_column)

    # mock where only first_row is nil or only last_row is nil
    half_host = custom_class.new([r1])
    half_host.define_singleton_method(:first_row) { nil }
    assert_nil(half_host.row_range)
    half_host2 = custom_class.new([r1])
    half_host2.define_singleton_method(:last_row) { nil }
    assert_nil(half_host2.row_range)

    # host with rows with empty cells
    empty_cells_host = custom_class.new([Xlsxrb::Elements::Row.new(index: 0, cells: [])])
    assert_nil(empty_cells_host.column_range)
    assert_equal(0..0, empty_cells_host.row_range)

    # row without #each to ensure cells is accessed
    row_without_each = Struct.new(:index, :cells).new(5, [Xlsxrb::Elements::Cell.new(row_index: 5, column_index: 4, value: 5)])
    host_without_each = custom_class.new([row_without_each])
    assert_equal(4..4, host_without_each.column_range)
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
    wrapped_cf = cf_rule.map { |r| Xlsxrb::Elements::ConditionalFormat.new(**r) }
    wrapped_dv = dv_rule.map { |r| Xlsxrb::Elements::DataValidation.new(**r) }
    assert_equal(wrapped_cf, ws_with_rules.conditional_formatting)
    assert_equal(wrapped_cf, ws_with_rules.conditional_formats)
    assert_equal(wrapped_dv, ws_with_rules.data_validations)

    # with
    new_cf = [{ sqref: "C1:C10", type: "containsText" }]
    wrapped_new_cf = new_cf.map { |r| Xlsxrb::Elements::ConditionalFormat.new(**r) }
    ws_updated = ws_with_rules.with(conditional_formatting: new_cf)
    assert_equal(wrapped_new_cf, ws_updated.conditional_formatting)
    assert_equal(wrapped_dv, ws_updated.data_validations)

    ws_updated_alias = ws_with_rules.with(conditional_formats: new_cf)
    assert_equal(wrapped_new_cf, ws_updated_alias.conditional_formatting)

    new_dv = [{ sqref: "D1:D10", type: "list", formula1: '"A,B,C"' }]
    wrapped_new_dv = new_dv.map { |r| Xlsxrb::Elements::DataValidation.new(**r) }
    ws_updated_dv = ws_with_rules.with(data_validations: new_dv)
    assert_equal(wrapped_cf, ws_updated_dv.conditional_formatting)
    assert_equal(wrapped_new_dv, ws_updated_dv.data_validations)

    # deconstruct_keys
    keys = ws_with_rules.deconstruct_keys(nil)
    assert_equal(wrapped_cf, keys[:conditional_formatting])
    assert_equal(wrapped_cf, keys[:conditional_formats])
    assert_equal(wrapped_dv, keys[:data_validations])

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

  test "worksheet handles merged_cells and auto_filter" do
    ws = Xlsxrb::Elements::Worksheet.new(name: "MergedSheet", merged_cells: ["A1:B2"], auto_filter: "A1:E10")
    assert_equal(["A1:B2"], ws.merged_cells)
    assert_equal("A1:E10", ws.auto_filter)

    ws_empty = Xlsxrb::Elements::Worksheet.new(name: "EmptySheet")
    assert_equal([], ws_empty.merged_cells)
    assert_nil(ws_empty.auto_filter)

    # with
    ws_updated = ws_empty.with(merged_cells: ["C1:D3"], auto_filter: "B2:F20")
    assert_equal(["C1:D3"], ws_updated.merged_cells)
    assert_equal("B2:F20", ws_updated.auto_filter)

    # deconstruct_keys
    keys = ws.deconstruct_keys(nil)
    assert_equal(["A1:B2"], keys[:merged_cells])
    assert_equal("A1:E10", keys[:auto_filter])

    # == and hash
    same_ws = Xlsxrb::Elements::Worksheet.new(name: "MergedSheet", merged_cells: ["A1:B2"], auto_filter: "A1:E10")
    diff_merge = Xlsxrb::Elements::Worksheet.new(name: "MergedSheet", merged_cells: ["A1:C3"], auto_filter: "A1:E10")
    diff_filter = Xlsxrb::Elements::Worksheet.new(name: "MergedSheet", merged_cells: ["A1:B2"], auto_filter: "A1:G10")

    assert_equal(ws, same_ws)
    assert_equal(ws.hash, same_ws.hash)
    refute_equal(ws, diff_merge)
    refute_equal(ws, diff_filter)
  end

  test "worksheet handles sheet structure, print, and view metadata" do
    tables = [{ id: 1, name: "Table1", ref: "A1:C10" }]
    shapes = [{ name: "Box", id: 2 }]
    freeze = { row: 1, col: 0 }
    page_setup = { orientation: "landscape" }
    page_margins = { left: 0.7, right: 0.7 }
    header_footer = { odd_header: "Header" }
    print_options = { grid_lines: true }
    protection = { sheet: true }
    view = { show_grid_lines: false }
    props = { tab_color: "FFFF0000" }
    fmt = { default_row_height: 24.0, default_col_width: 15.0 }
    row_brks = [10]
    col_brks = [5]

    ws = Xlsxrb::Elements::Worksheet.new(
      name: "StructSheet",
      tables: tables,
      shapes: shapes,
      freeze_pane: freeze,
      page_setup: page_setup,
      page_margins: page_margins,
      header_footer: header_footer,
      print_options: print_options,
      sheet_protection: protection,
      sheet_view: view,
      sheet_properties: props,
      sheet_format: fmt,
      row_breaks: row_brks,
      col_breaks: col_brks
    )

    assert_equal(1, ws.tables.size)
    assert_equal("Table1", ws.tables.first.name)
    assert_equal(1, ws.shapes.size)
    assert_equal("Box", ws.shapes.first.name)
    assert_equal(1, ws.freeze_pane.row)
    assert_equal("landscape", ws.page_setup.orientation)
    assert_equal(0.7, ws.page_margins.left)
    assert_equal("Header", ws.header_footer.odd_header)
    assert_equal(true, ws.print_options.grid_lines)
    assert_equal(true, ws.sheet_protection.sheet)
    assert_equal(false, ws.sheet_view.show_grid_lines)
    assert_equal("FFFF0000", ws.sheet_properties.tab_color)
    assert_equal(24.0, ws.sheet_format.default_row_height)
    assert_equal([{ id: 10, max: 16_383, man: 1 }], ws.row_breaks.map(&:to_h))
    assert_equal([{ id: 5, max: 1_048_575, man: 1 }], ws.column_breaks.map(&:to_h))

    # Defaults on empty worksheet
    ws_empty = Xlsxrb::Elements::Worksheet.new(name: "EmptyStruct")
    assert_equal([], ws_empty.tables)
    assert_equal([], ws_empty.shapes)
    assert_nil(ws_empty.freeze_pane)
    assert_true(ws_empty.page_setup.empty?)
    assert_nil(ws_empty.page_margins)
    assert_true(ws_empty.header_footer.empty?)
    assert_true(ws_empty.print_options.empty?)
    assert_nil(ws_empty.sheet_protection)
    assert_true(ws_empty.sheet_view.empty?)
    assert_true(ws_empty.sheet_properties.empty?)
    assert_true(ws_empty.sheet_format.empty?)
    assert_equal([], ws_empty.row_breaks)
    assert_equal([], ws_empty.col_breaks)

    # with
    ws_updated = ws_empty.with(freeze_pane: { row: 2, col: 1 }, sheet_properties: { tab_color: "00FF00" }, sheet_format: { default_row_height: 20 })
    assert_equal(2, ws_updated.freeze_pane.row)
    assert_equal(1, ws_updated.freeze_pane.column)
    assert_equal("00FF00", ws_updated.sheet_properties.tab_color)
    assert_equal(20, ws_updated.sheet_format.default_row_height)

    # deconstruct_keys
    keys = ws.deconstruct_keys(nil)
    assert_equal(ws.tables, keys[:tables])
    assert_equal(ws.freeze_pane, keys[:freeze_pane])
    assert_equal(ws.page_setup, keys[:page_setup])
    assert_equal(ws.sheet_format, keys[:sheet_format])

    # == and hash
    same_ws = Xlsxrb::Elements::Worksheet.new(
      name: "StructSheet",
      tables: tables, shapes: shapes, freeze_pane: freeze,
      page_setup: page_setup, page_margins: page_margins,
      header_footer: header_footer, print_options: print_options,
      sheet_protection: protection, sheet_view: view,
      sheet_properties: props, sheet_format: fmt, row_breaks: row_brks, col_breaks: col_brks
    )
    diff_ws = ws.with(freeze_pane: nil)
    diff_ws_fmt = ws.with(sheet_format: {})

    assert_equal(ws, same_ws)
    assert_equal(ws.hash, same_ws.hash)
    refute_equal(ws, diff_ws)
    refute_equal(ws, diff_ws_fmt)
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
    assert_equal("Summary!$B$10", wb.defined_name("TotalRevenue").value)
    assert_nil(wb.defined_name("NonExistent"))

    # Sheet-specific lookup
    assert_equal("Details!$A$1", wb.defined_name("LocalRate", sheet: "Details").value)
    assert_equal("Details!$A$1", wb.defined_name("LocalRate", sheet: 1).value)
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

  test "workbook document properties and protection accessors" do
    ws = Xlsxrb::Elements::Worksheet.new(name: "Sheet1")

    wb_empty = Xlsxrb::Elements::Workbook.new(sheets: [ws])
    assert_true(wb_empty.core_properties.empty?)
    assert_true(wb_empty.app_properties.empty?)
    assert_equal({}, wb_empty.custom_properties)
    assert_nil(wb_empty.workbook_protection)

    core_props = { title: "Report", creator: "Alice" }
    app_props = { application: "xlsxrb" }
    custom_props = [{ name: "Environment", type: :lpwstr, value: "Production" }]
    wb_prot = { lock_structure: true }

    wb_direct = Xlsxrb::Elements::Workbook.new(
      sheets: [ws],
      unmapped_data: {
        core_properties: core_props,
        app_properties: app_props,
        custom_properties: custom_props,
        workbook_protection: wb_prot
      }
    )
    assert_equal("Report", wb_direct.core_properties.title)
    assert_equal("Alice", wb_direct.core_properties.creator)
    assert_equal("xlsxrb", wb_direct.app_properties.app_name)
    assert_equal("Production", wb_direct.custom_properties["Environment"].value)
    assert_equal(true, wb_direct.workbook_protection.lock_structure)

    wb_facade = Xlsxrb::Elements::Workbook.new(
      sheets: [ws],
      unmapped_data: {
        facade: {
          core_properties: core_props,
          app_properties: app_props,
          custom_properties: custom_props,
          workbook_protection: wb_prot
        }
      }
    )
    assert_equal("Report", wb_facade.core_properties.title)
    assert_equal("Alice", wb_facade.core_properties.creator)
    assert_equal("xlsxrb", wb_facade.app_properties.app_name)
    assert_equal("Production", wb_facade.custom_properties["Environment"].value)
    assert_equal(true, wb_facade.workbook_protection.lock_structure)
  end

  test "workbook save delegates to Xlsxrb.write_workbook with default arguments" do
    ws = Xlsxrb::Elements::Worksheet.new(name: "Sheet1")
    wb = Xlsxrb::Elements::Workbook.new(sheets: [ws])
    called = nil
    orig_method = Xlsxrb.method(:write)
    Xlsxrb.define_singleton_method(:write) do |target, workbook, shared_strings: true, compression_level: nil, **options|
      called = { target: target, workbook: workbook, shared_strings: shared_strings, compression_level: compression_level, options: options }
      :saved
    end
    begin
      ret = wb.save("dummy.xlsx")
      assert_equal(:saved, ret)
      assert_equal("dummy.xlsx", called[:target])
      assert_same(wb, called[:workbook])
      assert_equal(true, called[:shared_strings])
      assert_nil(called[:compression_level])
      assert_equal({}, called[:options])
    ensure
      Xlsxrb.define_singleton_method(:write, orig_method)
    end
  end

  test "workbook save forwards custom keyword arguments and options" do
    ws = Xlsxrb::Elements::Worksheet.new(name: "Sheet1")
    wb = Xlsxrb::Elements::Workbook.new(sheets: [ws])
    called = nil
    orig_method = Xlsxrb.method(:write)
    Xlsxrb.define_singleton_method(:write) do |target, workbook, shared_strings: true, compression_level: nil, **options|
      called = { target: target, workbook: workbook, shared_strings: shared_strings, compression_level: compression_level, options: options }
      :saved_custom
    end
    begin
      ret = wb.save("out.xlsx", shared_strings: false, compression_level: 4, password: "secret")
      assert_equal(:saved_custom, ret)
      assert_equal("out.xlsx", called[:target])
      assert_same(wb, called[:workbook])
      assert_equal(false, called[:shared_strings])
      assert_equal(4, called[:compression_level])
      assert_equal({ password: "secret" }, called[:options])
    ensure
      Xlsxrb.define_singleton_method(:write, orig_method)
    end
  end

  test "workbook to_data delegates to Xlsxrb.write with default arguments" do
    ws = Xlsxrb::Elements::Worksheet.new(name: "Sheet1")
    wb = Xlsxrb::Elements::Workbook.new(sheets: [ws])
    called = nil
    orig_method = Xlsxrb.method(:write)
    Xlsxrb.define_singleton_method(:write) do |workbook, shared_strings: true, compression_level: nil, **options|
      called = { workbook: workbook, shared_strings: shared_strings, compression_level: compression_level, options: options }
      "binary_data"
    end
    begin
      ret = wb.to_data
      assert_equal("binary_data", ret)
      assert_same(wb, called[:workbook])
      assert_equal(true, called[:shared_strings])
      assert_nil(called[:compression_level])
      assert_equal({}, called[:options])
    ensure
      Xlsxrb.define_singleton_method(:write, orig_method)
    end
  end

  test "workbook to_data forwards custom keyword arguments and options" do
    ws = Xlsxrb::Elements::Worksheet.new(name: "Sheet1")
    wb = Xlsxrb::Elements::Workbook.new(sheets: [ws])
    called = nil
    orig_method = Xlsxrb.method(:write)
    Xlsxrb.define_singleton_method(:write) do |workbook, shared_strings: true, compression_level: nil, **options|
      called = { workbook: workbook, shared_strings: shared_strings, compression_level: compression_level, options: options }
      "binary_custom"
    end
    begin
      ret = wb.to_data(shared_strings: false, compression_level: 9, password: "pw")
      assert_equal("binary_custom", ret)
      assert_same(wb, called[:workbook])
      assert_equal(false, called[:shared_strings])
      assert_equal(9, called[:compression_level])
      assert_equal({ password: "pw" }, called[:options])
    ensure
      Xlsxrb.define_singleton_method(:write, orig_method)
    end
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

  test "Xlsxrb.formula supports type and ref" do
    f = Xlsxrb.formula("SUM(A1:A10)", type: :shared, ref: "A1:A10")
    assert_equal(:shared, f.type)
    assert_equal("A1:A10", f.ref)
  end

  test "Xlsxrb.array_formula creates an array Formula" do
    f = Xlsxrb.array_formula("SUM(B1:B10*C1:C10)", "A1:C1", cached_value: 0)
    assert_instance_of(Xlsxrb::Elements::Formula, f)
    assert_equal("SUM(B1:B10*C1:C10)", f.expression)
    assert_equal(:array, f.type)
    assert_equal("A1:C1", f.ref)
    assert_equal(0, f.cached_value)
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

  test "row values and worksheet each_row_values with type_cast: true" do
    c1 = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: 44_927, format_code: "yyyy-mm-dd")
    c2 = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 1, value: "hello")
    row = Xlsxrb::Elements::Row.new(index: 0, cells: [c1, c2])

    assert_equal([44_927, "hello"], row.values)
    assert_equal([44_927, "hello"], row.values(type_cast: false))
    assert_equal([Date.new(2023, 1, 1), "hello"], row.values(type_cast: true))

    ws = Xlsxrb::Elements::Worksheet.new(name: "Data", rows: [row])
    assert_kind_of(Enumerator, ws.each_row_values)
    assert_equal([[44_927, "hello"]], ws.each_row_values.to_a)
    assert_equal([[Date.new(2023, 1, 1), "hello"]], ws.each_row_values(type_cast: true).to_a)

    collected = []
    ws.each_row_values(type_cast: true) { |v| collected << v }
    assert_equal([[Date.new(2023, 1, 1), "hello"]], collected)
  end

  test "row empty? and worksheet trim_empty_rows" do
    empty_row1 = Xlsxrb::Elements::Row.new(index: 0, cells: [])
    assert_true(empty_row1.empty?)

    c_nil = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: nil)
    empty_row2 = Xlsxrb::Elements::Row.new(index: 0, cells: [c_nil])
    assert_true(empty_row2.empty?)

    c_empty_str = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "")
    empty_row3 = Xlsxrb::Elements::Row.new(index: 0, cells: [c_empty_str])
    assert_true(empty_row3.empty?)

    c_zero = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: 0)
    assert_false(Xlsxrb::Elements::Row.new(index: 0, cells: [c_zero]).empty?)

    c_false = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: false)
    assert_false(Xlsxrb::Elements::Row.new(index: 0, cells: [c_false]).empty?)

    c_val1 = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "R1")
    r1 = Xlsxrb::Elements::Row.new(index: 0, cells: [c_val1])

    c_val3 = Xlsxrb::Elements::Cell.new(row_index: 2, column_index: 0, value: "R3")
    r3 = Xlsxrb::Elements::Row.new(index: 2, cells: [c_val3])

    r2 = empty_row1
    r4 = empty_row2
    r5 = empty_row3

    ws = Xlsxrb::Elements::Worksheet.new(name: "Test", rows: [r1, r2, r3, r4, r5])
    assert_false(ws.trim_empty_rows?)
    assert_equal(5, ws.each_row.to_a.size)
    assert_equal(3, ws.each_row(trim_empty_rows: true).to_a.size)
    assert_equal([["R1"], [], ["R3"]], ws.each_row_values(trim_empty_rows: true).to_a)

    ws_trimmed = ws.with(trim_empty_rows: true)
    assert_true(ws_trimmed.trim_empty_rows?)
    assert_equal(3, ws_trimmed.each_row.to_a.size)
    assert_equal(5, ws_trimmed.each_row(trim_empty_rows: false).to_a.size)
  end

  test "image element properties and methods" do
    data = "PNG_SAMPLE_DATA".b
    img = Xlsxrb::Elements::Image.new(
      filename: "logo.png",
      name: "Company Logo",
      data: data,
      from_col: 1,
      from_row: 2,
      to_col: 4,
      to_row: 6,
      cx: 150_000,
      cy: 200_000
    )

    assert_equal("logo.png", img.filename)
    assert_equal("Company Logo", img.name)
    assert_equal("image/png", img.content_type)
    assert_equal(data, img.data)
    assert_equal(1, img.from_col)
    assert_equal(2, img.from_row)
    assert_equal(4, img.to_col)
    assert_equal(6, img.to_row)
    assert_equal(150_000, img.cx)
    assert_equal(200_000, img.cy)
    assert_equal("B3", img.ref)
    assert_equal("B3", img.cell)
    assert_equal("E7", img.to_cell_ref)
    assert_true(img.two_cell_anchor?)
    assert_false(img.one_cell_anchor?)
    assert_equal(15, img.bytesize)

    # 1-cell anchor
    img_1cell = Xlsxrb::Elements::Image.new(
      filename: "icon.jpg",
      from_col: 0,
      from_row: 0
    )
    assert_equal("image/jpeg", img_1cell.content_type)
    assert_equal("A1", img_1cell.ref)
    assert_nil(img_1cell.to_cell_ref)
    assert_true(img_1cell.one_cell_anchor?)
    assert_false(img_1cell.two_cell_anchor?)
    assert_equal(0, img_1cell.bytesize)

    # write_to file
    temp = Tempfile.new(["xlsxrb-image-test", ".png"])
    temp.close
    begin
      bytes_written = img.write_to(temp.path)
      assert_equal(15, bytes_written)
      assert_equal(data, File.binread(temp.path))

      # write_to with no data raises Error
      assert_raise(Xlsxrb::Error) { img_1cell.write_to(temp.path) }
    ensure
      temp.unlink
    end
  end

  test "image extraction from streaming and in-memory worksheets and workbook" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-img-reader", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    png_bytes = "\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x02\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDAT\x08\xd7c\xf8\xcf\xc0\x00\x00\x00\x02\x00\x01\xe2!\xbc3\x00\x00\x00\x00IEND\xaeB`\x82".b

    writer = Xlsxrb::Ooxml::Writer.new
    writer.set_cell("A1", "img test")
    writer.insert_image(
      png_bytes,
      ext: "png",
      from_col: 1,
      from_row: 2,
      to_col: 6,
      to_row: 12,
      name: "RoundTrip Pic"
    )
    writer.write(xlsx_path)

    wb = Xlsxrb.read(xlsx_path)
    sheet = wb.sheets.first

    # 1. StreamSheet access
    imgs = sheet.images
    assert_equal(1, imgs.size)
    img = imgs.first
    assert_equal("RoundTrip Pic", img.name)
    assert_equal(png_bytes, img.data)
    assert_equal("B3", img.ref)
    assert_equal("G13", img.to_cell_ref)
    assert_equal("image/png", img.content_type)
    assert_equal("image1.png", img.filename)

    # images_at coordinate lookups
    assert_equal([img], sheet.images_at("B3"))
    assert_equal([img], sheet.images_at("b3"))
    assert_equal([img], sheet.images_at(:b3))
    assert_equal([img], sheet.images_at(2, 1))
    assert_equal([], sheet.images_at("A1"))

    # with_images method
    assert_same(sheet, sheet.with_images)
    yielded = nil
    sheet.with_images { |s| yielded = s }
    assert_same(sheet, yielded)

    # 2. Workbook level access
    wb_imgs = wb.images
    assert_equal(1, wb_imgs.size)
    assert_equal(img, wb_imgs.first)
    assert_equal([img], wb.images_at(0, "B3"))
    assert_equal([img], wb.images_at("Sheet1", "B3"))
    assert_equal([], wb.images_at("NonExistent", "B3"))

    # 3. Elements::Worksheet access after load
    ws = sheet.load
    assert_equal(1, ws.images.size)
    assert_equal([img], ws.images_at("B3"))
    assert_equal([], ws.images_at("A1"))
    assert_same(ws, ws.with_images)
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end

  test "worksheet print_area and print_titles initialization and accessors" do
    ws = Xlsxrb::Elements::Worksheet.new(
      name: "PrintTest",
      print_area: "A1:D10",
      print_titles: { rows: "1:2", cols: "A:B" }
    )
    assert_equal("A1:D10", ws.print_area)
    assert_equal("1:2", ws.print_titles.rows)
    assert_equal("A:B", ws.print_titles.columns)

    ws_empty = Xlsxrb::Elements::Worksheet.new(name: "EmptyPrint")
    assert_nil(ws_empty.print_area)
    assert_nil(ws_empty.print_titles)
  end

  test "worksheet sparkline_groups initialization, with, and deconstruct_keys" do
    sgs = [{ type: "line", sparklines: [{ data_ref: "A1:E1", location_ref: "F1" }] }]
    ws = Xlsxrb::Elements::Worksheet.new(name: "SparkTest", sparkline_groups: sgs, print_area: "A1:F10")
    wrapped_sgs = sgs.map do |sg|
      sls = sg[:sparklines].map { |s| Xlsxrb::Elements::Sparkline.new(**s) }
      Xlsxrb::Elements::SparklineGroup.new(**sg, sparklines: sls)
    end
    assert_equal(wrapped_sgs, ws.sparkline_groups)
    assert_true(ws.sparkline_groups.frozen?)

    # default
    ws_empty = Xlsxrb::Elements::Worksheet.new(name: "EmptySpark")
    assert_equal([], ws_empty.sparkline_groups)

    # with
    new_sgs = [{ type: "column", sparklines: [{ data_ref: "A2:E2", location_ref: "F2" }] }]
    wrapped_new_sgs = new_sgs.map do |sg|
      sls = sg[:sparklines].map { |s| Xlsxrb::Elements::Sparkline.new(**s) }
      Xlsxrb::Elements::SparklineGroup.new(**sg, sparklines: sls)
    end
    ws2 = ws.with(sparkline_groups: new_sgs)
    assert_equal(wrapped_new_sgs, ws2.sparkline_groups)
    assert_equal("A1:F10", ws2.print_area)

    # deconstruct_keys
    keys = ws.deconstruct_keys(nil)
    assert_equal(wrapped_sgs, keys[:sparkline_groups])
    assert_equal("A1:F10", keys[:print_area])
  end

  test "cell threaded_comments accessors, fast_create, equality, with, deconstruct, and index access" do
    tc1 = { text: "Hello", author: "Alice", id: "{11111111-1111-1111-1111-111111111111}" }
    tc2 = { text: "World", author: "Bob", parent_id: "{11111111-1111-1111-1111-111111111111}" }
    cell = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: 42, threaded_comments: [tc1, tc2])

    assert_true(cell.threaded_comments?)
    assert_equal([tc1, tc2], cell.threaded_comments)
    assert_equal(tc1, cell.threaded_comment)
    assert_equal([tc1, tc2], cell[:threaded_comments])
    assert_equal(tc1, cell[:threaded_comment])

    empty_cell = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: 42)
    assert_false(empty_cell.threaded_comments?)
    assert_equal([], empty_cell.threaded_comments)
    assert_nil(empty_cell.threaded_comment)
    assert_equal([], empty_cell[:threaded_comments])
    assert_nil(empty_cell[:threaded_comment])

    fast_cell = Xlsxrb::Elements::Cell.fast_create(0, 0, 42, nil, nil)
    assert_false(fast_cell.threaded_comments?)
    assert_equal([], fast_cell.threaded_comments)

    cell_clone = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: 42, threaded_comments: [tc1, tc2])
    assert_equal(cell, cell_clone)
    assert_equal(cell.hash, cell_clone.hash)
    refute_equal(cell, empty_cell)

    with_cell = empty_cell.with(threaded_comments: [tc1])
    assert_equal([tc1], with_cell.threaded_comments)
    assert_equal(tc1, with_cell.threaded_comment)

    assert_equal([0, 0, 42, nil, nil, {}, [], nil, nil, nil, nil, [tc1, tc2]], cell.deconstruct)
    keys = cell.deconstruct_keys([:threaded_comments])
    assert_equal([tc1, tc2], keys[:threaded_comments])
  end

  test "worksheet threaded_comments accessors, ref lookup, with, equality, and hash" do
    tc1 = { ref: "B2", text: "Note 1", author: "Alice" }
    tc2 = { cell: "B2", text: "Reply 1", author: "Bob" }
    tc3 = { ref: "C3", text: "Other", author: "Charlie" }
    ws = Xlsxrb::Elements::Worksheet.new(name: "Sheet1", threaded_comments: [tc1, tc2, tc3])

    wrapped = [tc1, tc2, tc3].map { |tc| Xlsxrb::Elements::ThreadedComment.new(**tc) }
    assert_equal(wrapped, ws.threaded_comments)
    assert_equal(wrapped[0], ws.threaded_comment("B2"))
    assert_equal(wrapped[0], ws.threaded_comment(1, 1)) # row 1, col 1 is B2
    assert_equal([wrapped[0], wrapped[1]], ws.threaded_comments_at("B2"))
    assert_equal([wrapped[0], wrapped[1]], ws.threaded_comments_at(1, 1))
    assert_equal(wrapped[2], ws.threaded_comment("C3"))
    assert_equal([wrapped[2]], ws.threaded_comments_at("C3"))
    assert_nil(ws.threaded_comment("A1"))
    assert_equal([], ws.threaded_comments_at("A1"))

    by_ref = ws.threaded_comments_by_ref
    assert_equal([wrapped[0], wrapped[1]], by_ref["B2"])
    assert_equal([wrapped[2]], by_ref["C3"])

    ws_copy = Xlsxrb::Elements::Worksheet.new(name: "Sheet1", threaded_comments: [tc1, tc2, tc3])
    assert_equal(ws, ws_copy)
    assert_equal(ws.hash, ws_copy.hash)

    ws_empty = Xlsxrb::Elements::Worksheet.new(name: "Sheet1")
    refute_equal(ws, ws_empty)

    ws_with = ws_empty.with(threaded_comments: [tc3])
    assert_equal([wrapped[2]], ws_with.threaded_comments)

    keys = ws.deconstruct_keys(nil)
    assert_equal(wrapped, keys[:threaded_comments])
  end

  def capture_stderr
    old_stderr = $stderr
    old_verbose = $VERBOSE
    $VERBOSE = false
    $stderr = StringIO.new
    yield
    $stderr.string
  ensure
    $stderr = old_stderr
    $VERBOSE = old_verbose
  end

  # --- PivotTable ---

  test "pivot_table initialization with normalized fields, attributes, and options catch-all" do
    # Defaults and col_fields normalization
    pt1 = Xlsxrb::Elements::PivotTable.new(name: "Pivot1", ref: "A1:D10")
    assert_equal("Pivot1", pt1.name)
    assert_equal("A1:D10", pt1.ref)
    assert_nil(pt1.source_ref)
    assert_nil(pt1.row_fields)
    assert_equal([], pt1.column_fields)
    assert_nil(pt1.data_fields)
    assert_nil(pt1.page_fields)
    assert_nil(pt1.data_caption)
    assert_nil(pt1.page_wrap)
    assert_nil(pt1.column_grand_totals)
    assert_nil(pt1.row_grand_totals)
    assert_nil(pt1.show_row_headers)
    assert_nil(pt1.show_column_headers)
    assert_nil(pt1.show_row_stripes)
    assert_nil(pt1.show_column_stripes)
    assert_nil(pt1.cache)
    assert_nil(pt1.options)

    # location fallback when ref is nil
    pt_loc = Xlsxrb::Elements::PivotTable.new(location: "C3:E15")
    assert_equal("C3:E15", pt_loc.ref)

    # Full attributes, location precedence over nil, col_fields keyword normalization, and unknown kwargs options
    pt2 = Xlsxrb::Elements::PivotTable.new(
      name: "Pivot2",
      ref: "B2:E20",
      location: nil,
      source_ref: "Source!A1:Z100",
      row_fields: ["Region"],
      col_fields: ["Year"],
      data_fields: [{ name: "Sales", fld: 2, subtotal: "sum" }],
      page_fields: ["Department"],
      data_caption: "Data",
      page_wrap: 2,
      column_grand_totals: true,
      row_grand_totals: false,
      show_row_headers: true,
      show_column_headers: false,
      show_row_stripes: true,
      show_column_stripes: false,
      cache: { cache_id: 1, source_ref: "Source!A1:Z100" },
      options: { pre_opt: 42 },
      custom_flag: true
    )
    assert_equal("Pivot2", pt2.name)
    assert_equal("B2:E20", pt2.ref)
    assert_equal("Source!A1:Z100", pt2.source_ref)
    assert_equal(["Region"], pt2.row_fields)
    assert_equal(["Year"], pt2.column_fields)
    assert_equal([{ name: "Sales", fld: 2, subtotal: "sum" }], pt2.data_fields)
    assert_equal(["Department"], pt2.page_fields)
    assert_equal("Data", pt2.data_caption)
    assert_equal(2, pt2.page_wrap)
    assert_equal(true, pt2.column_grand_totals)
    assert_equal(false, pt2.row_grand_totals)
    assert_equal(true, pt2.show_row_headers)
    assert_equal(false, pt2.show_column_headers)
    assert_equal(true, pt2.show_row_stripes)
    assert_equal(false, pt2.show_column_stripes)
    assert_equal({ cache_id: 1, source_ref: "Source!A1:Z100" }, pt2.cache)
    assert_equal({ pre_opt: 42, custom_flag: true }, pt2.options)
    assert_true(pt2.options.frozen?)

    # column_fields takes precedence over col_fields
    pt3 = Xlsxrb::Elements::PivotTable.new(column_fields: ["Month"], col_fields: ["Year"])
    assert_equal(["Month"], pt3.column_fields)

    # options unrolls into top-level Hash in to_h
    h = pt2.to_h
    assert_equal(true, h[:custom_flag])
    assert_equal(42, h[:pre_opt])
    assert_nil(h[:options])
  end

  # --- TableColumn ---

  test "table_column initialization and calculated_column_formula alias" do
    tc1 = Xlsxrb::Elements::TableColumn.new(
      id: 1,
      name: "Total",
      totals_row_function: "sum",
      totals_row_label: "Sum",
      totals_row_formula: "SUM([Total])",
      data_dxf_id: 5,
      calculated_column_formula: "[@[Price]]*[@[Qty]]"
    )
    assert_equal(1, tc1.id)
    assert_equal("Total", tc1.name)
    assert_equal("sum", tc1.totals_row_function)
    assert_equal("Sum", tc1.totals_row_label)
    assert_equal("SUM([Total])", tc1.totals_row_formula)
    assert_equal(5, tc1.data_dxf_id)
    assert_equal("[@[Price]]*[@[Qty]]", tc1.calculated_column_formula)

    # data_formula keyword compatibility
    tc2 = Xlsxrb::Elements::TableColumn.new(id: 2, name: "Tax", data_formula: "[@[Total]]*0.1")
    assert_equal("[@[Total]]*0.1", tc2.calculated_column_formula)

    # calculated_column_formula takes precedence when both provided
    tc3 = Xlsxrb::Elements::TableColumn.new(id: 3, name: "Disc", calculated_column_formula: "F1", data_formula: "F2")
    assert_equal("F1", tc3.calculated_column_formula)

    # defaults
    tc4 = Xlsxrb::Elements::TableColumn.new(id: 4, name: "Plain")
    assert_nil(tc4.calculated_column_formula)
  end

  # --- ThreadedComment ---

  test "threaded_comment normalization for done, author_id, and created_at aliases" do
    # ref and cell normalization
    tc_ref = Xlsxrb::Elements::ThreadedComment.new(ref: "A1", cell: nil)
    assert_equal("A1", tc_ref.ref)

    tc_cell = Xlsxrb::Elements::ThreadedComment.new(ref: nil, cell: "B2")
    assert_equal("B2", tc_cell.ref)

    # resolved and done normalization
    tc1 = Xlsxrb::Elements::ThreadedComment.new(text: "Hello", resolved: true, done: false)
    assert_equal(true, tc1.resolved)

    tc2 = Xlsxrb::Elements::ThreadedComment.new(text: "Hello", resolved: false, done: true)
    assert_equal(false, tc2.resolved)

    tc3 = Xlsxrb::Elements::ThreadedComment.new(text: "Hello", done: true)
    assert_equal(true, tc3.resolved)

    tc4 = Xlsxrb::Elements::ThreadedComment.new(text: "Hello", done: false)
    assert_equal(false, tc4.resolved)

    tc5 = Xlsxrb::Elements::ThreadedComment.new(text: "Hello")
    assert_nil(tc5.resolved)

    # author_id and person_id normalization
    tc6 = Xlsxrb::Elements::ThreadedComment.new(author_id: "auth1", person_id: nil)
    assert_equal("auth1", tc6.author_id)

    tc7 = Xlsxrb::Elements::ThreadedComment.new(author_id: nil, person_id: "pers1")
    assert_equal("pers1", tc7.author_id)

    tc8 = Xlsxrb::Elements::ThreadedComment.new(author_id: "auth1", person_id: "pers1")
    assert_equal("auth1", tc8.author_id)

    # created_at, timestamp, and dt normalization
    tc9 = Xlsxrb::Elements::ThreadedComment.new(created_at: "2026-01-01", timestamp: nil, dt: nil)
    assert_equal("2026-01-01", tc9.created_at)

    tc10 = Xlsxrb::Elements::ThreadedComment.new(created_at: nil, timestamp: "2026-01-02", dt: nil)
    assert_equal("2026-01-02", tc10.created_at)

    tc11 = Xlsxrb::Elements::ThreadedComment.new(created_at: nil, timestamp: nil, dt: "2026-01-03")
    assert_equal("2026-01-03", tc11.created_at)

    tc12 = Xlsxrb::Elements::ThreadedComment.new(created_at: "2026-01-01", timestamp: "2026-01-02", dt: "2026-01-03")
    assert_equal("2026-01-01", tc12.created_at)

    tc13 = Xlsxrb::Elements::ThreadedComment.new(created_at: nil, timestamp: "2026-01-02", dt: "2026-01-03")
    assert_equal("2026-01-02", tc13.created_at)

    # other members
    tc14 = Xlsxrb::Elements::ThreadedComment.new(id: "tc_id", text: "Comm", author: "User", parent_id: "pid")
    assert_equal("tc_id", tc14.id)
    assert_equal("Comm", tc14.text)
    assert_equal("User", tc14.author)
    assert_equal("pid", tc14.parent_id)
  end

  # --- Comment ---

  test "comment initialization with visible attribute" do
    c_default = Xlsxrb::Elements::Comment.new(text: "Note")
    assert_equal(false, c_default.visible)
    assert_nil(c_default.ref)

    c_visible = Xlsxrb::Elements::Comment.new(
      ref: "C5",
      text: "Note",
      author: "Alice",
      r_id: "rId1",
      shape_id: 1025,
      author_id: 2,
      visible: true
    )
    assert_equal("C5", c_visible.ref)
    assert_equal("C5", c_visible.cell)
    assert_equal("Note", c_visible.text)
    assert_equal("Alice", c_visible.author)
    assert_equal("rId1", c_visible.r_id)
    assert_equal(1025, c_visible.shape_id)
    assert_equal(2, c_visible.author_id)
    assert_equal(true, c_visible.visible)

    c_hidden = Xlsxrb::Elements::Comment.new(cell: "D6", visible: false)
    assert_equal("D6", c_hidden.ref)
    assert_equal(false, c_hidden.visible)
  end

  # --- FilterColumn ---

  test "filter_column initialization with hidden, show_button, and hidden_button normalization" do
    fc_default = Xlsxrb::Elements::FilterColumn.new(col_id: 0)
    assert_equal(0, fc_default.col_id)
    assert_equal(false, fc_default.hidden)
    assert_equal(true, fc_default.show_button)

    fc_all = Xlsxrb::Elements::FilterColumn.new(
      col_id: 1,
      filters: ["A"],
      custom_filters: [{ op: "gt", val: "10" }],
      dynamic_filter: { type: "aboveAverage" },
      top10: { val: 5 },
      color_filter: { dxf_id: 0 },
      icon_filter: { icon_id: 1 },
      hidden: true,
      show_button: true
    )
    assert_equal(1, fc_all.col_id)
    assert_equal(["A"], fc_all.filters)
    assert_equal([{ op: "gt", val: "10" }], fc_all.custom_filters)
    assert_equal({ type: "aboveAverage" }, fc_all.dynamic_filter)
    assert_equal({ val: 5 }, fc_all.top10)
    assert_equal({ dxf_id: 0 }, fc_all.color_filter)
    assert_equal({ icon_id: 1 }, fc_all.icon_filter)
    assert_equal(true, fc_all.hidden)
    assert_equal(true, fc_all.show_button)

    fc_hidden_btn = Xlsxrb::Elements::FilterColumn.new(col_id: 2, hidden_button: true)
    assert_equal(true, fc_hidden_btn.hidden)

    fc_hide_button = Xlsxrb::Elements::FilterColumn.new(col_id: 3, show_button: false)
    assert_equal(false, fc_hide_button.show_button)

    fc_show_button = Xlsxrb::Elements::FilterColumn.new(col_id: 4, show_button: true)
    assert_equal(true, fc_show_button.show_button)
  end

  # --- Worksheet and Workbook repeat_rows and repeat_columns ---

  test "worksheet and workbook builders repeat_rows and repeat_columns without deprecation warnings" do
    wb = Xlsxrb::WorkbookBuilder.new
    sheet_builder = nil
    wb.sheet("Data") { |s| sheet_builder = s }

    warnings = capture_stderr do
      sheet_builder.repeat_rows(0, 1)
      sheet_builder.repeat_columns(0, 1)
    end
    assert_empty(warnings)
    assert_instance_of(Xlsxrb::Elements::PrintTitles, sheet_builder.print_titles)
    assert_equal("1:2", sheet_builder.print_titles.rows)
    assert_equal("A:B", sheet_builder.print_titles.columns)

    wb_warnings = capture_stderr do
      wb.repeat_rows(0, 1)
      wb.repeat_columns(0, 1)
    end
    assert_empty(wb_warnings)
  end

  # --- Worksheet#full_calc_on_load? ---

  test "worksheet full_calc_on_load? evaluates calculation and sheet properties" do
    ws1 = Xlsxrb::Elements::Worksheet.new(name: "S1", sheet_calc_pr: { full_calc_on_load: true }, sheet_properties: { full_calc_on_load: false })
    assert_true(ws1.full_calc_on_load?)

    ws2 = Xlsxrb::Elements::Worksheet.new(name: "S2", sheet_calc_pr: { full_calc_on_load: false }, sheet_properties: { full_calc_on_load: true })
    assert_true(ws2.full_calc_on_load?)

    ws3 = Xlsxrb::Elements::Worksheet.new(name: "S3", sheet_calc_pr: { full_calc_on_load: false }, sheet_properties: { full_calc_on_load: false })
    assert_false(ws3.full_calc_on_load?)

    ws4 = Xlsxrb::Elements::Worksheet.new(name: "S4")
    assert_false(ws4.full_calc_on_load?)

    ws5 = Xlsxrb::Elements::Worksheet.new(name: "S5", sheet_calc_pr: Xlsxrb::Elements::SheetCalcProperties.new(full_calc_on_load: "truthy_str"), sheet_properties: { full_calc_on_load: false })
    assert_false(ws5.full_calc_on_load?)

    ws6 = Xlsxrb::Elements::Worksheet.new(
      name: "S6",
      sheet_calc_pr: Xlsxrb::Elements::SheetCalcProperties.new(full_calc_on_load: false),
      sheet_properties: Xlsxrb::Elements::SheetProperties.new(full_calc_on_load: "truthy_str")
    )
    assert_false(ws6.full_calc_on_load?)
  end

  # --- Styles#font_size ---

  test "styles font_size resolution handles numeric and non-numeric values" do
    styles = Xlsxrb::Elements::Styles.new(
      fonts: [
        { name: "Calibri", size: 11 },
        { name: "Arial", size: 10.5 },
        { name: "Consolas", size: nil },
        { name: "Courier", size: "invalid" }
      ]
    )
    assert_equal(11, styles.font_size(0))
    assert_equal(10.5, styles.font_size(1))
    assert_nil(styles.font_size(2))
    assert_nil(styles.font_size(3))
    assert_nil(styles.font_size(99))
  end

  # --- Cell and Worksheet Value Object Consistency in Xlsxrb.build ---

  test "Xlsxrb.build attaches Elements Value Objects for hyperlink, comment, and threaded_comment" do
    wb = Xlsxrb.build do |b|
      b.sheet "Report" do |s|
        s.row ["Cell A1", "Cell B1"]
        s.hyperlink "A1", "https://example.com", tooltip: "Link tip", display: "Link display"
        s.comment "A1", "Comment note", author: "Auditor"
        s.threaded_comment "A1", "Thread message", author: "Reviewer"
      end
    end

    sheet = wb.sheet("Report")
    cell_a1 = sheet["A1"]

    assert_instance_of(Xlsxrb::Elements::Hyperlink, cell_a1.hyperlink)
    assert_instance_of(Xlsxrb::Elements::Hyperlink, sheet.hyperlink("A1"))
    assert_equal(sheet.hyperlink("A1"), cell_a1.hyperlink)
    assert_equal("https://example.com", cell_a1.hyperlink.url)
    assert_equal("Link tip", cell_a1.hyperlink.tooltip)
    assert_equal("Link display", cell_a1.hyperlink.display)

    assert_instance_of(Xlsxrb::Elements::Comment, cell_a1.comment)
    assert_instance_of(Xlsxrb::Elements::Comment, sheet.comment("A1"))
    assert_equal(sheet.comment("A1"), cell_a1.comment)
    assert_equal("Comment note", cell_a1.comment.text)
    assert_equal("Auditor", cell_a1.comment.author)

    assert_instance_of(Xlsxrb::Elements::ThreadedComment, cell_a1.threaded_comment)
    assert_instance_of(Xlsxrb::Elements::ThreadedComment, sheet.threaded_comment("A1"))
    assert_equal(sheet.threaded_comment("A1"), cell_a1.threaded_comment)
    assert_equal("Thread message", cell_a1.threaded_comment.text)
    assert_equal("Reviewer", cell_a1.threaded_comment.author)
    assert_equal([cell_a1.threaded_comment], cell_a1.threaded_comments)

    cell_b1 = sheet["B1"]
    assert_nil(cell_b1.hyperlink)
    assert_nil(cell_b1.comment)
    assert_nil(cell_b1.threaded_comment)
    assert_equal([], cell_b1.threaded_comments)
  end

  # --- Boolean Predicate Methods on Metadata Records, Row, and Column ---

  test "elements metadata value object boolean predicate methods" do
    # Font
    f1 = Xlsxrb::Elements::Font.new(bold: true, italic: true, strike: true, outline: true, shadow: true, condense: true, extend: true)
    assert_true(f1.bold?)
    assert_true(f1.italic?)
    assert_true(f1.strike?)
    assert_true(f1.outline?)
    assert_true(f1.shadow?)
    assert_true(f1.condense?)
    assert_true(f1.extend?)

    f0 = Xlsxrb::Elements::Font.new(bold: false, italic: false, strike: false, outline: false, shadow: false, condense: false, extend: false)
    assert_false(f0.bold?)
    assert_false(f0.italic?)
    assert_false(f0.strike?)
    assert_false(f0.outline?)
    assert_false(f0.shadow?)
    assert_false(f0.condense?)
    assert_false(f0.extend?)

    f_nil = Xlsxrb::Elements::Font.new
    assert_false(f_nil.bold?)
    assert_false(f_nil.italic?)
    assert_false(f_nil.strike?)
    assert_false(f_nil.outline?)
    assert_false(f_nil.shadow?)
    assert_false(f_nil.condense?)
    assert_false(f_nil.extend?)

    # Alignment
    a1 = Xlsxrb::Elements::Alignment.new(wrap_text: true, shrink_to_fit: true, justify_last_line: true)
    assert_true(a1.wrap_text?)
    assert_true(a1.shrink_to_fit?)
    assert_true(a1.justify_last_line?)

    a0 = Xlsxrb::Elements::Alignment.new(wrap_text: false, shrink_to_fit: false, justify_last_line: false)
    assert_false(a0.wrap_text?)
    assert_false(a0.shrink_to_fit?)
    assert_false(a0.justify_last_line?)

    # Protection
    p1 = Xlsxrb::Elements::Protection.new(locked: true, hidden: true)
    assert_true(p1.locked?)
    assert_true(p1.hidden?)

    p0 = Xlsxrb::Elements::Protection.new(locked: false, hidden: false)
    assert_false(p0.locked?)
    assert_false(p0.hidden?)

    # CellXf
    xf1 = Xlsxrb::Elements::CellXf.new(
      quote_prefix: true, pivot_button: true,
      apply_number_format: true, apply_font: true, apply_fill: true,
      apply_border: true, apply_alignment: true, apply_protection: true
    )
    assert_true(xf1.quote_prefix?)
    assert_true(xf1.pivot_button?)
    assert_true(xf1.apply_number_format?)
    assert_true(xf1.apply_font?)
    assert_true(xf1.apply_fill?)
    assert_true(xf1.apply_border?)
    assert_true(xf1.apply_alignment?)
    assert_true(xf1.apply_protection?)

    xf0 = Xlsxrb::Elements::CellXf.new(
      quote_prefix: false, pivot_button: false,
      apply_number_format: false, apply_font: false, apply_fill: false,
      apply_border: false, apply_alignment: false, apply_protection: false
    )
    assert_false(xf0.quote_prefix?)
    assert_false(xf0.pivot_button?)
    assert_false(xf0.apply_number_format?)
    assert_false(xf0.apply_font?)
    assert_false(xf0.apply_fill?)
    assert_false(xf0.apply_border?)
    assert_false(xf0.apply_alignment?)
    assert_false(xf0.apply_protection?)

    # PageSetup
    ps1 = Xlsxrb::Elements::PageSetup.new(black_and_white: true, draft: true)
    assert_true(ps1.black_and_white?)
    assert_true(ps1.draft?)

    ps0 = Xlsxrb::Elements::PageSetup.new(black_and_white: false, draft: false)
    assert_false(ps0.black_and_white?)
    assert_false(ps0.draft?)

    # HeaderFooter
    hf1 = Xlsxrb::Elements::HeaderFooter.new(different_first: true, different_odd_even: true, scale_with_doc: true, align_with_margins: true)
    assert_true(hf1.different_first?)
    assert_true(hf1.different_odd_even?)
    assert_true(hf1.scale_with_doc?)
    assert_true(hf1.align_with_margins?)

    hf0 = Xlsxrb::Elements::HeaderFooter.new(different_first: false, different_odd_even: false, scale_with_doc: false, align_with_margins: false)
    assert_false(hf0.different_first?)
    assert_false(hf0.different_odd_even?)
    assert_false(hf0.scale_with_doc?)
    assert_false(hf0.align_with_margins?)

    # PrintOptions
    po1 = Xlsxrb::Elements::PrintOptions.new(grid_lines: true, grid_lines_set: true, headings: true, horizontal_centered: true, vertical_centered: true)
    assert_true(po1.grid_lines?)
    assert_true(po1.grid_lines_set?)
    assert_true(po1.headings?)
    assert_true(po1.horizontal_centered?)
    assert_true(po1.vertical_centered?)

    po0 = Xlsxrb::Elements::PrintOptions.new(grid_lines: false, grid_lines_set: false, headings: false, horizontal_centered: false, vertical_centered: false)
    assert_false(po0.grid_lines?)
    assert_false(po0.grid_lines_set?)
    assert_false(po0.headings?)
    assert_false(po0.horizontal_centered?)
    assert_false(po0.vertical_centered?)

    # SheetProperties
    spr1 = Xlsxrb::Elements::SheetProperties.new(fit_to_page: true, apply_styles: true, summary_below: true, summary_right: true, show_outline_symbols: true)
    assert_true(spr1.fit_to_page?)
    assert_true(spr1.apply_styles?)
    assert_true(spr1.summary_below?)
    assert_true(spr1.summary_right?)
    assert_true(spr1.show_outline_symbols?)

    spr0 = Xlsxrb::Elements::SheetProperties.new(fit_to_page: false, apply_styles: false, summary_below: false, summary_right: false, show_outline_symbols: false)
    assert_false(spr0.fit_to_page?)
    assert_false(spr0.apply_styles?)
    assert_false(spr0.summary_below?)
    assert_false(spr0.summary_right?)
    assert_false(spr0.show_outline_symbols?)

    # SheetFormat
    sf1 = Xlsxrb::Elements::SheetFormat.new(custom_height: true, zero_height: true, thick_top: true, thick_bottom: true)
    assert_true(sf1.custom_height?)
    assert_true(sf1.zero_height?)
    assert_true(sf1.thick_top?)
    assert_true(sf1.thick_bottom?)

    sf0 = Xlsxrb::Elements::SheetFormat.new(custom_height: false, zero_height: false, thick_top: false, thick_bottom: false)
    assert_false(sf0.custom_height?)
    assert_false(sf0.zero_height?)
    assert_false(sf0.thick_top?)
    assert_false(sf0.thick_bottom?)

    # SheetView
    sv1 = Xlsxrb::Elements::SheetView.new(
      window_protection: true, show_formulas: true, show_grid_lines: true,
      show_row_col_headers: true, show_zeros: true, right_to_left: true,
      tab_selected: true, show_ruler: true, show_outline_symbols: true,
      default_grid_color: true, show_white_space: true
    )
    assert_true(sv1.window_protection?)
    assert_true(sv1.show_formulas?)
    assert_true(sv1.show_grid_lines?)
    assert_true(sv1.show_row_col_headers?)
    assert_true(sv1.show_zeros?)
    assert_true(sv1.right_to_left?)
    assert_true(sv1.tab_selected?)
    assert_true(sv1.show_ruler?)
    assert_true(sv1.show_outline_symbols?)
    assert_true(sv1.default_grid_color?)
    assert_true(sv1.show_white_space?)

    sv0 = Xlsxrb::Elements::SheetView.new(
      window_protection: false, show_formulas: false, show_grid_lines: false,
      show_row_col_headers: false, show_zeros: false, right_to_left: false,
      tab_selected: false, show_ruler: false, show_outline_symbols: false,
      default_grid_color: false, show_white_space: false
    )
    assert_false(sv0.window_protection?)
    assert_false(sv0.show_formulas?)
    assert_false(sv0.show_grid_lines?)
    assert_false(sv0.show_row_col_headers?)
    assert_false(sv0.show_zeros?)
    assert_false(sv0.right_to_left?)
    assert_false(sv0.tab_selected?)
    assert_false(sv0.show_ruler?)
    assert_false(sv0.show_outline_symbols?)
    assert_false(sv0.default_grid_color?)
    assert_false(sv0.show_white_space?)

    # SheetCalcProperties
    scp1 = Xlsxrb::Elements::SheetCalcProperties.new(full_calc_on_load: true)
    assert_true(scp1.full_calc_on_load?)
    scp0 = Xlsxrb::Elements::SheetCalcProperties.new(full_calc_on_load: false)
    assert_false(scp0.full_calc_on_load?)

    # PageBreak
    pb1 = Xlsxrb::Elements::PageBreak.new(id: 1, manual: true)
    assert_true(pb1.manual?)
    pb0 = Xlsxrb::Elements::PageBreak.new(id: 1, manual: false)
    assert_false(pb0.manual?)

    # DefinedName
    dn1 = Xlsxrb::Elements::DefinedName.new(name: "Test", formula: "A1", hidden: true)
    assert_true(dn1.hidden?)
    dn0 = Xlsxrb::Elements::DefinedName.new(name: "Test", formula: "A1", hidden: false)
    assert_false(dn0.hidden?)

    # WorkbookProperties
    wbp1 = Xlsxrb::Elements::WorkbookProperties.new(date1904: true, filter_privacy: true, show_sheet_tabs: true, full_calc_on_load: true, refresh_all: true)
    assert_true(wbp1.date1904?)
    assert_true(wbp1.filter_privacy?)
    assert_true(wbp1.show_sheet_tabs?)
    assert_true(wbp1.full_calc_on_load?)
    assert_true(wbp1.refresh_all?)
    wbp0 = Xlsxrb::Elements::WorkbookProperties.new(date1904: false, filter_privacy: false, show_sheet_tabs: false, full_calc_on_load: false, refresh_all: false)
    assert_false(wbp0.date1904?)
    assert_false(wbp0.filter_privacy?)
    assert_false(wbp0.show_sheet_tabs?)
    assert_false(wbp0.full_calc_on_load?)
    assert_false(wbp0.refresh_all?)

    # WorkbookProtection
    wpr1 = Xlsxrb::Elements::WorkbookProtection.new(lock_structure: true, lock_windows: true, lock_revision: true)
    assert_true(wpr1.lock_structure?)
    assert_true(wpr1.lock_windows?)
    assert_true(wpr1.lock_revision?)
    wpr0 = Xlsxrb::Elements::WorkbookProtection.new(lock_structure: false, lock_windows: false, lock_revision: false)
    assert_false(wpr0.lock_structure?)
    assert_false(wpr0.lock_windows?)
    assert_false(wpr0.lock_revision?)

    # SheetProtection
    shp1 = Xlsxrb::Elements::SheetProtection.new(
      sheet: true, objects: true, scenarios: true, format_cells: true,
      format_columns: true, format_rows: true, insert_columns: true,
      insert_rows: true, insert_hyperlinks: true, delete_columns: true,
      delete_rows: true, select_locked_cells: true, sort: true,
      auto_filter: true, pivot_tables: true, select_unlocked_cells: true
    )
    assert_true(shp1.sheet?)
    assert_true(shp1.objects?)
    assert_true(shp1.scenarios?)
    assert_true(shp1.format_cells?)
    assert_true(shp1.format_columns?)
    assert_true(shp1.format_rows?)
    assert_true(shp1.insert_columns?)
    assert_true(shp1.insert_rows?)
    assert_true(shp1.insert_hyperlinks?)
    assert_true(shp1.delete_columns?)
    assert_true(shp1.delete_rows?)
    assert_true(shp1.select_locked_cells?)
    assert_true(shp1.sort?)
    assert_true(shp1.auto_filter?)
    assert_true(shp1.pivot_tables?)
    assert_true(shp1.select_unlocked_cells?)

    shp0 = Xlsxrb::Elements::SheetProtection.new(
      sheet: false, objects: false, scenarios: false, format_cells: false,
      format_columns: false, format_rows: false, insert_columns: false,
      insert_rows: false, insert_hyperlinks: false, delete_columns: false,
      delete_rows: false, select_locked_cells: false, sort: false,
      auto_filter: false, pivot_tables: false, select_unlocked_cells: false
    )
    assert_false(shp0.sheet?)
    assert_false(shp0.objects?)
    assert_false(shp0.scenarios?)
    assert_false(shp0.format_cells?)
    assert_false(shp0.format_columns?)
    assert_false(shp0.format_rows?)
    assert_false(shp0.insert_columns?)
    assert_false(shp0.insert_rows?)
    assert_false(shp0.insert_hyperlinks?)
    assert_false(shp0.delete_columns?)
    assert_false(shp0.delete_rows?)
    assert_false(shp0.select_locked_cells?)
    assert_false(shp0.sort?)
    assert_false(shp0.auto_filter?)
    assert_false(shp0.pivot_tables?)
    assert_false(shp0.select_unlocked_cells?)

    # FilterColumn
    fc1 = Xlsxrb::Elements::FilterColumn.new(col_id: 0, hidden: true, show_button: true)
    assert_true(fc1.hidden?)
    assert_true(fc1.show_button?)
    fc0 = Xlsxrb::Elements::FilterColumn.new(col_id: 0, hidden: false, show_button: false)
    assert_false(fc0.hidden?)
    assert_false(fc0.show_button?)

    # Comment
    c_vis = Xlsxrb::Elements::Comment.new(text: "Hi", visible: true)
    assert_true(c_vis.visible?)
    c_invis = Xlsxrb::Elements::Comment.new(text: "Hi", visible: false)
    assert_false(c_invis.visible?)

    # ThreadedComment
    tc_res = Xlsxrb::Elements::ThreadedComment.new(text: "Note", resolved: true, done: true)
    assert_true(tc_res.resolved?)
    tc_unres = Xlsxrb::Elements::ThreadedComment.new(text: "Note", resolved: false, done: false)
    assert_false(tc_unres.resolved?)

    # DataValidation
    dv1 = Xlsxrb::Elements::DataValidation.new(sqref: "A1", allow_blank: true, show_input_message: true, show_error_message: true, show_drop_down: true)
    assert_true(dv1.allow_blank?)
    assert_true(dv1.show_input_message?)
    assert_true(dv1.show_error_message?)
    assert_true(dv1.show_drop_down?)

    dv0 = Xlsxrb::Elements::DataValidation.new(sqref: "A1", allow_blank: false, show_input_message: false, show_error_message: false, show_drop_down: false)
    assert_false(dv0.allow_blank?)
    assert_false(dv0.show_input_message?)
    assert_false(dv0.show_error_message?)
    assert_false(dv0.show_drop_down?)

    # Table
    tbl1 = Xlsxrb::Elements::Table.new(id: 1, name: "T1", display_name: "T1", ref: "A1:B2", totals_row_shown: true, header_row_shown: true)
    assert_true(tbl1.totals_row_shown?)
    assert_true(tbl1.header_row_shown?)

    tbl0 = Xlsxrb::Elements::Table.new(id: 1, name: "T1", display_name: "T1", ref: "A1:B2", totals_row_count: 0, header_row_count: 0)
    assert_false(tbl0.totals_row_shown?)
    assert_false(tbl0.header_row_shown?)

    # SparklineGroup
    sg1 = Xlsxrb::Elements::SparklineGroup.new(markers: true, high: true, low: true, first: true, last: true, negative: true)
    assert_true(sg1.markers?)
    assert_true(sg1.high?)
    assert_true(sg1.low?)
    assert_true(sg1.first?)
    assert_true(sg1.last?)
    assert_true(sg1.negative?)

    sg0 = Xlsxrb::Elements::SparklineGroup.new(markers: false, high: false, low: false, first: false, last: false, negative: false)
    assert_false(sg0.markers?)
    assert_false(sg0.high?)
    assert_false(sg0.low?)
    assert_false(sg0.first?)
    assert_false(sg0.last?)
    assert_false(sg0.negative?)

    # PivotTable
    pt1 = Xlsxrb::Elements::PivotTable.new(
      name: "P1", cache_id: 1, data_caption: "Sum",
      column_grand_totals: true, row_grand_totals: true,
      show_row_headers: true, show_column_headers: true,
      show_row_stripes: true, show_column_stripes: true
    )
    assert_true(pt1.column_grand_totals?)
    assert_true(pt1.row_grand_totals?)
    assert_true(pt1.show_row_headers?)
    assert_true(pt1.show_column_headers?)
    assert_true(pt1.show_row_stripes?)
    assert_true(pt1.show_column_stripes?)

    pt0 = Xlsxrb::Elements::PivotTable.new(
      name: "P1", cache_id: 1, data_caption: "Sum",
      column_grand_totals: false, row_grand_totals: false,
      show_row_headers: false, show_column_headers: false,
      show_row_stripes: false, show_column_stripes: false
    )
    assert_false(pt0.column_grand_totals?)
    assert_false(pt0.row_grand_totals?)
    assert_false(pt0.show_row_headers?)
    assert_false(pt0.show_column_headers?)
    assert_false(pt0.show_row_stripes?)
    assert_false(pt0.show_column_stripes?)

    # IgnoredError
    ie1 = Xlsxrb::Elements::IgnoredError.new(
      sqref: "A1:B2", eval_error: true, two_digit_text_year: true,
      number_stored_as_text: true, formula: true, formula_range: true,
      unlocked_formula: true, empty_cell_reference: true,
      list_data_validation: true, calculated_column: true
    )
    assert_true(ie1.eval_error?)
    assert_true(ie1.two_digit_text_year?)
    assert_true(ie1.number_stored_as_text?)
    assert_true(ie1.formula?)
    assert_true(ie1.formula_range?)
    assert_true(ie1.unlocked_formula?)
    assert_true(ie1.empty_cell_reference?)
    assert_true(ie1.list_data_validation?)
    assert_true(ie1.calculated_column?)

    ie0 = Xlsxrb::Elements::IgnoredError.new(
      sqref: "A1:B2", eval_error: false, two_digit_text_year: false,
      number_stored_as_text: false, formula: false, formula_range: false,
      unlocked_formula: false, empty_cell_reference: false,
      list_data_validation: false, calculated_column: false
    )
    assert_false(ie0.eval_error?)
    assert_false(ie0.two_digit_text_year?)
    assert_false(ie0.number_stored_as_text?)
    assert_false(ie0.formula?)
    assert_false(ie0.formula_range?)
    assert_false(ie0.unlocked_formula?)
    assert_false(ie0.empty_cell_reference?)
    assert_false(ie0.list_data_validation?)
    assert_false(ie0.calculated_column?)
  end

  test "row, column, and stream_row boolean predicates" do
    row = Xlsxrb::Elements::Row.new(index: 1, custom_height: true)
    assert_true(row.custom_height?)
    row_def = Xlsxrb::Elements::Row.new(index: 1)
    assert_false(row_def.custom_height?)

    col = Xlsxrb::Elements::Column.new(index: 0, custom_width: true, best_fit: true, collapsed: true)
    assert_true(col.custom_width?)
    assert_true(col.best_fit?)
    assert_true(col.collapsed?)

    col_def = Xlsxrb::Elements::Column.new(index: 0)
    assert_false(col_def.custom_width?)
    assert_false(col_def.best_fit?)
    assert_false(col_def.collapsed?)
  end

  test "elements workbook enumerable support and blockless each" do
    ws1 = Xlsxrb::Elements::Worksheet.new(name: "A")
    ws2 = Xlsxrb::Elements::Worksheet.new(name: "B")
    wb = Xlsxrb::Elements::Workbook.new(sheets: [ws1, ws2])

    assert_true(Xlsxrb::Elements::Workbook.include?(Enumerable))

    enum = wb.each
    assert_instance_of(Enumerator, enum)
    assert_equal(wb.sheets, enum.to_a)
    assert_equal([ws1, ws2], enum.to_a)

    assert_equal(%w[A B], wb.map(&:name))
    assert_equal(ws1, wb.find { |s| s.name == "A" })
    assert_equal(1, wb.count { |s| s.name == "B" })
    assert_true(wb.any? { |s| s.name == "A" })
    assert_false(wb.none? { |s| s.name == "A" })
  end

  test "pattern matching deconstruct and deconstruct_keys on cell, row, and column" do
    # Cell
    hl = Xlsxrb::Elements::Hyperlink.new(url: "https://example.com", tooltip: "tip")
    cm = Xlsxrb::Elements::Comment.new(text: "comment")
    tc = Xlsxrb::Elements::ThreadedComment.new(text: "thread")
    full_cell = Xlsxrb::Elements::Cell.new(
      row_index: 2,
      column_index: 1,
      value: "hello",
      formula: "=A1+1",
      style_index: 4,
      format_code: "yyyy-mm-dd",
      hyperlink: hl,
      comment: cm,
      threaded_comments: [tc]
    )
    expected_cell_keys = {
      ref: "B3",
      row_index: 2,
      column_index: 1,
      value: "hello",
      formula: Xlsxrb::Elements::Formula.new(expression: "A1+1"),
      style_index: 4,
      format_code: "yyyy-mm-dd",
      hyperlink: hl,
      comment: cm,
      threaded_comments: [tc]
    }
    assert_equal(expected_cell_keys, full_cell.deconstruct_keys(nil))
    case full_cell
    in { ref: "B3", value: "hello", formula: Xlsxrb::Elements::Formula, format_code: "yyyy-mm-dd" }
      assert_true(true)
    else
      flunk("Cell pattern match failed")
    end

    # Row
    c1 = Xlsxrb::Elements::Cell.new(row_index: 1, column_index: 0, value: "hello")
    c2 = Xlsxrb::Elements::Cell.new(row_index: 1, column_index: 1, value: 99)
    row = Xlsxrb::Elements::Row.new(
      index: 1,
      cells: [c1, c2],
      height: 25.0,
      hidden: true,
      custom_height: true,
      collapsed: true,
      outline_level: 2
    )

    assert_equal(["hello", 99], row.deconstruct)
    case row
    in ["hello", 99]
      assert_true(true)
    else
      flunk("Row array pattern match failed")
    end

    expected_row_keys = {
      index: 1,
      cells: [c1, c2],
      values: ["hello", 99],
      height: 25.0,
      hidden: true,
      custom_height: true,
      collapsed: true,
      outline_level: 2,
      style_index: nil
    }
    assert_equal(expected_row_keys, row.deconstruct_keys(nil))
    case row
    in { index: 1, values: ["hello", 99], custom_height: true, outline_level: 2 }
      assert_true(true)
    else
      flunk("Row hash pattern match failed")
    end

    row_false = Xlsxrb::Elements::Row.new(
      index: 0,
      cells: [],
      height: nil,
      hidden: false,
      custom_height: false,
      collapsed: false,
      outline_level: nil
    )
    assert_equal(
      { index: 0, cells: [], values: [], height: nil, hidden: false, custom_height: false, collapsed: false, outline_level: nil, style_index: nil },
      row_false.deconstruct_keys(nil)
    )

    # Column
    col = Xlsxrb::Elements::Column.new(
      index: 3,
      width: 18.0,
      hidden: true,
      custom_width: true,
      best_fit: true,
      collapsed: true,
      outline_level: 1,
      style_index: 5
    )
    expected_col_keys = {
      index: 3,
      width: 18.0,
      hidden: true,
      custom_width: true,
      best_fit: true,
      collapsed: true,
      outline_level: 1,
      style_index: 5
    }
    assert_equal(expected_col_keys, col.deconstruct_keys(nil))
    case col
    in { index: 3, width: 18.0, custom_width: true, best_fit: true, collapsed: true }
      assert_true(true)
    else
      flunk("Column hash pattern match failed")
    end

    col_false = Xlsxrb::Elements::Column.new(
      index: 0,
      width: nil,
      hidden: false,
      custom_width: false,
      best_fit: false,
      collapsed: false,
      outline_level: nil,
      style_index: nil
    )
    assert_equal(
      { index: 0, width: nil, hidden: false, custom_width: false, best_fit: false, collapsed: false, outline_level: nil, style_index: nil },
      col_false.deconstruct_keys(nil)
    )
  end
end

class ElementsDeprecationTest < Test::Unit::TestCase
  test "deprecated elements methods emit appropriate deprecation warnings" do
    # RichTextRun#[]
    rtr = Xlsxrb::Elements::RichTextRun.new(text: "Hello", font: { bold: true })
    err = capture_stderr { assert_equal("Hello", rtr[:text]) }
    assert_match(/\[Xlsxrb DEPRECATION\] RichTextRun#\[\] is deprecated; use #text or #font instead\./, err)
    err = capture_stderr { assert_equal({ bold: true }, rtr[:font]) }
    assert_match(/\[Xlsxrb DEPRECATION\] RichTextRun#\[\] is deprecated; use #text or #font instead\./, err)

    # Hyperlink#== (Hash)
    hl = Xlsxrb::Elements::Hyperlink.new(url: "https://example.com")
    err = capture_stderr { assert_equal(hl, { url: "https://example.com" }) }
    assert_match(/\[Xlsxrb DEPRECATION\] Comparing Hyperlink with Hash is deprecated; compare with Hyperlink instead\./, err)

    # CustomProperty#== (Hash)
    cp = Xlsxrb::Elements::CustomProperty.new(name: "Author", value: "Alice")
    err = capture_stderr { assert_equal(cp, { name: "Author", value: "Alice" }) }
    assert_match(/\[Xlsxrb DEPRECATION\] Comparing CustomProperty with Hash is deprecated; compare with CustomProperty instead\./, err)

    # Styles#== (Hash)
    styles = Xlsxrb::Elements::Styles.new
    err = capture_stderr { assert_equal(styles, {}) }
    assert_match(/\[Xlsxrb DEPRECATION\] Comparing Styles with Hash is deprecated; compare with Styles instead\./, err)

    # PrintTitles#cols
    pt = Xlsxrb::Elements::PrintTitles.new(columns: "$A:$B")
    err = capture_stderr { assert_equal("$A:$B", pt.cols) }
    assert_match(/\[Xlsxrb DEPRECATION\] PrintTitles#cols is deprecated; use #columns instead\./, err)

    # FreezePane#col
    fp = Xlsxrb::Elements::FreezePane.new(column: 2)
    err = capture_stderr { assert_equal(2, fp.col) }
    assert_match(/\[Xlsxrb DEPRECATION\] FreezePane#col is deprecated; use #column instead\./, err)

    # WorkbookProperties#first_tab
    wp = Xlsxrb::Elements::WorkbookProperties.new(first_sheet: 1)
    err = capture_stderr { assert_equal(1, wp.first_tab) }
    assert_match(/\[Xlsxrb DEPRECATION\] WorkbookProperties#first_tab is deprecated; use #first_sheet instead\./, err)

    # Table#show_row_stripes
    tbl = Xlsxrb::Elements::Table.new(name: "T1", ref: "A1:B2", style_info: { show_row_stripes: true })
    assert_true(tbl.show_row_stripes)

    # SheetProperties#outline_pr
    sp = Xlsxrb::Elements::SheetProperties.new(outline_properties: { summary_below: true })
    err = capture_stderr { assert_equal({ summary_below: true }, sp.outline_pr) }
    assert_match(/\[Xlsxrb DEPRECATION\] SheetProperties#outline_pr is deprecated; use #outline_properties instead\./, err)

    # BackgroundImage#rel_id, #ext, #filename
    bg = Xlsxrb::Elements::BackgroundImage.new(r_id: "rId1", extension: "png", path: "bg.png")
    err = capture_stderr { assert_equal("rId1", bg.rel_id) }
    assert_match(/\[Xlsxrb DEPRECATION\] BackgroundImage#rel_id is deprecated; use #r_id instead\./, err)
    err = capture_stderr { assert_equal("png", bg.ext) }
    assert_match(/\[Xlsxrb DEPRECATION\] BackgroundImage#ext is deprecated; use #extension instead\./, err)
    err = capture_stderr { assert_equal("bg.png", bg.filename) }
    assert_match(/\[Xlsxrb DEPRECATION\] BackgroundImage#filename is deprecated; use #path instead\./, err)

    # PivotTable#location
    pvt = Xlsxrb::Elements::PivotTable.new(ref: "A1:C10")
    err = capture_stderr { assert_equal("A1:C10", pvt.location) }
    assert_match(/\[Xlsxrb DEPRECATION\] PivotTable#location is deprecated; use #ref instead\./, err)

    # SheetFormat: default_col_width, outline_level_col, base_col_width
    sf = Xlsxrb::Elements::SheetFormat.new(default_column_width: 10, outline_level_column: 1, base_column_width: 8)
    err = capture_stderr { assert_equal(10, sf.default_col_width) }
    assert_match(/\[Xlsxrb DEPRECATION\] SheetFormat#default_col_width is deprecated; use #default_column_width instead\./, err)
    err = capture_stderr { assert_equal(1, sf.outline_level_col) }
    assert_match(/\[Xlsxrb DEPRECATION\] SheetFormat#outline_level_col is deprecated; use #outline_level_column instead\./, err)
    err = capture_stderr { assert_equal(8, sf.base_col_width) }
    assert_match(/\[Xlsxrb DEPRECATION\] SheetFormat#base_col_width is deprecated; use #base_column_width instead\./, err)

    # Chart: rid, chart_type, from_col, to_col
    chart = Xlsxrb::Elements::Chart.new(r_id: "rId2", type: "bar", from_column: 1, to_column: 5)
    err = capture_stderr { assert_equal("rId2", chart.rid) }
    assert_match(/\[Xlsxrb DEPRECATION\] Chart#rid is deprecated; use #r_id instead\./, err)
    err = capture_stderr { assert_equal("bar", chart.chart_type) }
    assert_match(/\[Xlsxrb DEPRECATION\] Chart#chart_type is deprecated; use #type instead\./, err)
    err = capture_stderr { assert_equal(1, chart.from_col) }
    assert_match(/\[Xlsxrb DEPRECATION\] Chart#from_col is deprecated; use #from_column instead\./, err)
    err = capture_stderr { assert_equal(5, chart.to_col) }
    assert_match(/\[Xlsxrb DEPRECATION\] Chart#to_col is deprecated; use #to_column instead\./, err)

    # Comment: cell
    comm = Xlsxrb::Elements::Comment.new(ref: "A1", text: "Note")
    err = capture_stderr { assert_equal("A1", comm.cell) }
    assert_match(/\[Xlsxrb DEPRECATION\] Comment#cell is deprecated; use #ref instead\./, err)

    # ThreadedComment: cell, done, done?
    tc = Xlsxrb::Elements::ThreadedComment.new(ref: "B2", resolved: true)
    err = capture_stderr { assert_equal("B2", tc.cell) }
    assert_match(/\[Xlsxrb DEPRECATION\] ThreadedComment#cell is deprecated; use #ref instead\./, err)
    err = capture_stderr { assert_equal(true, tc.done) }
    assert_match(/\[Xlsxrb DEPRECATION\] ThreadedComment#done is deprecated; use #resolved instead\./, err)
    err = capture_stderr { assert_true(tc.done?) }
    assert_match(/\[Xlsxrb DEPRECATION\] ThreadedComment#done\? is deprecated; use #resolved\? instead\./, err)

    # DataValidation: show_dropdown, show_dropdown?
    dv = Xlsxrb::Elements::DataValidation.new(sqref: "C1", show_drop_down: true)
    err = capture_stderr { assert_equal(true, dv.show_dropdown) }
    assert_match(/\[Xlsxrb DEPRECATION\] DataValidation#show_dropdown is deprecated; use #show_drop_down instead\./, err)
    err = capture_stderr { assert_true(dv.show_dropdown?) }
    assert_match(/\[Xlsxrb DEPRECATION\] DataValidation#show_dropdown\? is deprecated; use #show_drop_down\? instead\./, err)

    # ConditionalFormatRule & ConditionalFormat: formulas
    cfr = Xlsxrb::Elements::ConditionalFormatRule.new(type: "expression", formula: "A1>0")
    err = capture_stderr { assert_equal("A1>0", cfr.formulas) }
    assert_match(/\[Xlsxrb DEPRECATION\] ConditionalFormatRule#formulas is deprecated; use #formula instead\./, err)
    cf = Xlsxrb::Elements::ConditionalFormat.new(sqref: "A1:A10", formula: "A1>0")
    err = capture_stderr { assert_equal("A1>0", cf.formulas) }
    assert_match(/\[Xlsxrb DEPRECATION\] ConditionalFormat#formulas is deprecated; use #formula instead\./, err)

    # TableColumn: data_formula
    tcol = Xlsxrb::Elements::TableColumn.new(id: 1, name: "Col1", calculated_column_formula: "F1")
    err = capture_stderr { assert_equal("F1", tcol.data_formula) }
    assert_match(/\[Xlsxrb DEPRECATION\] TableColumn#data_formula is deprecated; use #calculated_column_formula instead\./, err)

    # SparklineGroup: display_empty, max, min
    sg = Xlsxrb::Elements::SparklineGroup.new(display_empty_cells_as: "zero", manual_max: 100.0, manual_min: 0.0)
    err = capture_stderr { assert_equal("zero", sg.display_empty) }
    assert_match(/\[Xlsxrb DEPRECATION\] SparklineGroup#display_empty is deprecated; use #display_empty_cells_as instead\./, err)
    err = capture_stderr { assert_equal(100.0, sg.max) }
    assert_match(/\[Xlsxrb DEPRECATION\] SparklineGroup#max is deprecated; use #manual_max instead\./, err)
    err = capture_stderr { assert_equal(0.0, sg.min) }
    assert_match(/\[Xlsxrb DEPRECATION\] SparklineGroup#min is deprecated; use #manual_min instead\./, err)

    # HeaderFooterImage: ext, filename, rid
    hfi = Xlsxrb::Elements::HeaderFooterImage.new(r_id: "rId3", path: "logo.png", extension: "png")
    err = capture_stderr { assert_equal("png", hfi.ext) }
    assert_match(/\[Xlsxrb DEPRECATION\] HeaderFooterImage#ext is deprecated; use #extension instead\./, err)
    err = capture_stderr { assert_equal("logo.png", hfi.filename) }
    assert_match(/\[Xlsxrb DEPRECATION\] HeaderFooterImage#filename is deprecated; use #path instead\./, err)
    err = capture_stderr { assert_equal("rId3", hfi.rid) }
    assert_match(/\[Xlsxrb DEPRECATION\] HeaderFooterImage#rid is deprecated; use #r_id instead\./, err)

    # Font: sz
    fnt = Xlsxrb::Elements::Font.new(size: 12)
    err = capture_stderr { assert_equal(12, fnt.sz) }
    assert_match(/\[Xlsxrb DEPRECATION\] Font#sz is deprecated; use #size instead\./, err)

    # Fill: pattern
    fill = Xlsxrb::Elements::Fill.new(pattern_type: "solid")
    err = capture_stderr { assert_equal("solid", fill.pattern) }
    assert_match(/\[Xlsxrb DEPRECATION\] Fill#pattern is deprecated; use #pattern_type instead\./, err)

    # Formula predicates without ?
    formula = Xlsxrb::Elements::Formula.new(expression: "SUM(A1:A2)", aca: true, bx: false, dt2d: true, dtr: false)
    err = capture_stderr { assert_equal(true, formula.always_calculate_array) }
    assert_match(/\[Xlsxrb DEPRECATION\] Formula#always_calculate_array is deprecated; use #always_calculate_array\? instead\./, err)
    err = capture_stderr { assert_equal(false, formula.assign_to_name) }
    assert_match(/\[Xlsxrb DEPRECATION\] Formula#assign_to_name is deprecated; use #assign_to_name\? instead\./, err)
    err = capture_stderr { assert_equal(true, formula.data_table_2d) }
    assert_match(/\[Xlsxrb DEPRECATION\] Formula#data_table_2d is deprecated; use #data_table_2d\? instead\./, err)
    err = capture_stderr { assert_equal(false, formula.data_table_row) }
    assert_match(/\[Xlsxrb DEPRECATION\] Formula#data_table_row is deprecated; use #data_table_row\? instead\./, err)

    # Column#letter
    col = Xlsxrb::Elements::Column.new(index: 0)
    err = capture_stderr { assert_equal("A", col.letter) }
    assert_match(/\[Xlsxrb DEPRECATION\] Column#letter is deprecated; use #column_letter instead\./, err)

    # Image#cell_ref
    img = Xlsxrb::Elements::Image.new(filename: "pic.png", from_col: 0, from_row: 0)
    err = capture_stderr { assert_equal("A1", img.cell_ref) }
    assert_match(/\[Xlsxrb DEPRECATION\] Image#cell_ref is deprecated; use #ref instead\./, err)

    # Styles#format_code
    st = Xlsxrb::Elements::Styles.new(cell_xfs: [{ num_fmt_id: 14 }], num_fmts: { 14 => "yyyy-mm-dd" })
    err = capture_stderr { assert_equal("yyyy-mm-dd", st.format_code(0)) }
    assert_match(/\[Xlsxrb DEPRECATION\] Styles#format_code is deprecated; use #number_format instead\./, err)

    # Worksheet#initialize with conditional_formatting: and col_breaks:
    err = capture_stderr { Xlsxrb::Elements::Worksheet.new(name: "S1", conditional_formatting: [cf]) }
    assert_match(/\[Xlsxrb DEPRECATION\] Passing conditional_formatting: to Worksheet\.new is deprecated; use conditional_formats: instead\./, err)
    err = capture_stderr { Xlsxrb::Elements::Worksheet.new(name: "S2", col_breaks: [1, 2]) }
    assert_match(/\[Xlsxrb DEPRECATION\] Passing col_breaks: to Worksheet\.new is deprecated; use column_breaks: instead\./, err)
  end

  test "ADR 0004 & 0006 metadata predicates and normalization" do
    # AppProperties predicates
    app_t = Xlsxrb::Elements::AppProperties.new(scale_crop: true, links_up_to_date: true, shared_doc: true, hyperlinks_changed: true)
    assert_true(app_t.scale_crop?)
    assert_true(app_t.links_up_to_date?)
    assert_true(app_t.shared_doc?)
    assert_true(app_t.hyperlinks_changed?)

    app_f = Xlsxrb::Elements::AppProperties.new(scale_crop: false, links_up_to_date: false, shared_doc: false, hyperlinks_changed: false)
    assert_false(app_f.scale_crop?)
    assert_false(app_f.links_up_to_date?)
    assert_false(app_f.shared_doc?)
    assert_false(app_f.hyperlinks_changed?)

    # PageSetup predicates
    ps_t = Xlsxrb::Elements::PageSetup.new(use_first_page_number: true, use_printer_defaults: true, horizontal_centered: true, vertical_centered: true)
    assert_true(ps_t.use_first_page_number?)
    assert_true(ps_t.use_printer_defaults?)
    assert_true(ps_t.horizontal_centered?)
    assert_true(ps_t.vertical_centered?)

    ps_f = Xlsxrb::Elements::PageSetup.new
    assert_false(ps_f.use_first_page_number?)
    assert_false(ps_f.use_printer_defaults?)
    assert_false(ps_f.horizontal_centered?)
    assert_false(ps_f.vertical_centered?)

    # SheetProperties predicates
    sp_t = Xlsxrb::Elements::SheetProperties.new(
      filter_mode: true, enable_format_conditions_calculation: true, full_calc_on_load: true,
      tab_color_auto: true, auto_page_breaks: true, sync_horizontal: true, sync_vertical: true,
      transition_evaluation: true, transition_entry: true, published: true
    )
    assert_true(sp_t.filter_mode?)
    assert_true(sp_t.enable_format_conditions_calculation?)
    assert_true(sp_t.full_calc_on_load?)
    assert_true(sp_t.tab_color_auto?)
    assert_true(sp_t.auto_page_breaks?)
    assert_true(sp_t.sync_horizontal?)
    assert_true(sp_t.sync_vertical?)
    assert_true(sp_t.transition_evaluation?)
    assert_true(sp_t.transition_entry?)
    assert_true(sp_t.published?)

    sp_f = Xlsxrb::Elements::SheetProperties.new
    assert_false(sp_f.filter_mode?)
    assert_false(sp_f.enable_format_conditions_calculation?)
    assert_false(sp_f.full_calc_on_load?)
    assert_false(sp_f.tab_color_auto?)
    assert_false(sp_f.auto_page_breaks?)
    assert_false(sp_f.sync_horizontal?)
    assert_false(sp_f.sync_vertical?)
    assert_false(sp_f.transition_evaluation?)
    assert_false(sp_f.transition_entry?)
    assert_false(sp_f.published?)

    # SortState predicates
    ss_t = Xlsxrb::Elements::SortState.new(case_sensitive: true)
    assert_true(ss_t.case_sensitive?)
    ss_f = Xlsxrb::Elements::SortState.new
    assert_false(ss_f.case_sensitive?)

    # ConditionalFormatRule predicates
    cfr_t = Xlsxrb::Elements::ConditionalFormatRule.new(above_average: true, equal_average: true, bottom: true, percent: true, stop_if_true: true)
    assert_true(cfr_t.above_average?)
    assert_true(cfr_t.equal_average?)
    assert_true(cfr_t.bottom?)
    assert_true(cfr_t.percent?)
    assert_true(cfr_t.stop_if_true?)

    cfr_f = Xlsxrb::Elements::ConditionalFormatRule.new
    assert_false(cfr_f.above_average?)
    assert_false(cfr_f.equal_average?)
    assert_false(cfr_f.bottom?)
    assert_false(cfr_f.percent?)
    assert_false(cfr_f.stop_if_true?)

    # ConditionalFormat predicates
    cf_t = Xlsxrb::Elements::ConditionalFormat.new(above_average: true, equal_average: true, bottom: true, percent: true, stop_if_true: true, pivot: true)
    assert_true(cf_t.above_average?)
    assert_true(cf_t.equal_average?)
    assert_true(cf_t.bottom?)
    assert_true(cf_t.percent?)
    assert_true(cf_t.stop_if_true?)
    assert_true(cf_t.pivot?)

    cf_f = Xlsxrb::Elements::ConditionalFormat.new
    assert_false(cf_f.above_average?)
    assert_false(cf_f.equal_average?)
    assert_false(cf_f.bottom?)
    assert_false(cf_f.percent?)
    assert_false(cf_f.stop_if_true?)
    assert_false(cf_f.pivot?)

    # Border predicates
    b_t = Xlsxrb::Elements::Border.new(diagonal_up: true, diagonal_down: true)
    assert_true(b_t.diagonal_up?)
    assert_true(b_t.diagonal_down?)

    b_f = Xlsxrb::Elements::Border.new
    assert_false(b_f.diagonal_up?)
    assert_false(b_f.diagonal_down?)

    # TableStyleInfo predicates
    tsi_t = Xlsxrb::Elements::TableStyleInfo.new(show_first_column: true, show_last_column: true, show_row_stripes: true, show_column_stripes: true)
    assert_true(tsi_t.show_first_column?)
    assert_true(tsi_t.show_last_column?)
    assert_true(tsi_t.show_row_stripes?)
    assert_true(tsi_t.show_column_stripes?)

    tsi_f = Xlsxrb::Elements::TableStyleInfo.new
    assert_false(tsi_f.show_first_column?)
    assert_false(tsi_f.show_last_column?)
    assert_false(tsi_f.show_row_stripes?)
    assert_false(tsi_f.show_column_stripes?)

    # Table style predicates
    tbl_t = Xlsxrb::Elements::Table.new(id: 1, name: "T1", ref: "A1:B2", style_info: tsi_t)
    assert_true(tbl_t.show_first_column?)
    assert_true(tbl_t.show_last_column?)
    assert_true(tbl_t.show_row_stripes?)
    assert_true(tbl_t.show_column_stripes?)

    tbl_f = Xlsxrb::Elements::Table.new(id: 1, name: "T1", ref: "A1:B2", style_info: tsi_f)
    assert_false(tbl_f.show_first_column?)
    assert_false(tbl_f.show_last_column?)
    assert_false(tbl_f.show_row_stripes?)
    assert_false(tbl_f.show_column_stripes?)

    tbl_nil = Xlsxrb::Elements::Table.new(id: 1, name: "T1", ref: "A1:B2")
    assert_false(tbl_nil.show_first_column?)
    assert_false(tbl_nil.show_last_column?)
    assert_false(tbl_nil.show_row_stripes?)
    assert_false(tbl_nil.show_column_stripes?)
  end
end
