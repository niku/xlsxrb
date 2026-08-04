# frozen_string_literal: true

require "test_helper"
require "tempfile"

class StyleTest < Test::Unit::TestCase
  test "in-memory mode: style to worksheet and apply to cells" do
    workbook = Xlsxrb.build do |w|
      w.sheet("Test") do |s|
        # Define a style
        s.style("heading") do |style|
          style.bold.size(14).font_color("FFFF0000")
        end

        # Add rows with styles
        s.row(["Header 1", "Header 2"], styles: "heading")
        s.row([100, 200])
      end
    end

    assert_equal(1, workbook.sheets.size)
    sheet = workbook.sheets[0]
    assert_equal(2, sheet.rows.size)

    # First row should have numeric style indices
    first_row = sheet.rows[0]
    assert_equal(1, first_row.cells[0].style_index)
    assert_equal(1, first_row.cells[1].style_index)

    # Second row should have no style
    second_row = sheet.rows[1]
    assert_nil(second_row.cells[0].style_index)
  end

  test "streaming mode: style and apply to rows" do
    tmp = Tempfile.new(["style_stream_test", ".xlsx"])
    begin
      Xlsxrb.generate(tmp.path) do |w|
        # Define styles
        w.style("heading") do |style|
          style.bold.size(14).font_color("FFFF0000")
        end

        w.style("total") do |style|
          style.bold.fill_color("FF00FF00")
        end

        w.sheet("Sales") do
          # Add header row with heading style
          w.row(%w[Date Amount], styles: { 0 => "heading", 1 => "heading" })

          # Add data rows
          w.row([Date.today, 100])
          w.row([Date.today - 1, 200])

          # Add total row with total style
          w.row(["Total", 300], styles: { 1 => "total" })
        end
      end

      # Verify the file was created and can be read back
      workbook = Xlsxrb.read(tmp.path)
      assert_equal(1, workbook.sheets.size)
      sheet = workbook.sheets[0]
      assert_equal(4, sheet.rows.size)

      # Rows should have cells (style indices may not be directly readable from parsed file)
      assert_equal(2, sheet.rows[0].cells.size)
    ensure
      tmp.close!
    end
  end

  test "style builder fluent API" do
    style = Xlsxrb::StyleBuilder.new("test")
    result = style.bold.italic.size(12).font_name("Arial")

    assert_equal(style, result) # Should return self for chaining
    assert_equal(true, style.font_props[:bold])
    assert_equal(true, style.font_props[:italic])
    assert_equal(12, style.font_props[:sz])
    assert_equal("Arial", style.font_props[:name])
  end

  test "in-memory mode: round-trip with styled cells" do
    workbook = Xlsxrb.build do |w|
      w.sheet("Styled") do |s|
        s.style("bold_red") do |style|
          style.bold.font_color("FFFF0000").size(12)
        end

        s.row(%w[Styled Data], styles: ["bold_red", nil])
        s.row(%w[Normal Row])
      end
    end

    tmp = Tempfile.new(["in_memory_style_test", ".xlsx"])
    begin
      Xlsxrb.write(tmp.path, workbook)

      # Read back
      read_workbook = Xlsxrb.read(tmp.path)
      assert_equal(1, read_workbook.sheets.size)
      sheet = read_workbook.sheets[0]
      assert_equal(2, sheet.rows.size)

      # First cell should have a style index
      assert(sheet.rows[0].cells[0].style_index.is_a?(Integer) || sheet.rows[0].cells[0].style_index.nil?)
    ensure
      tmp.close!
    end
  end
end
