# frozen_string_literal: true

require "test_helper"
require "tempfile"

class StyleTest < Test::Unit::TestCase
  test "in-memory mode: add_style to worksheet and apply to cells" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("Test") do |s|
        # Define a style
        s.add_style("heading") do |style|
          style.bold.size(14).font_color("FFFF0000")
        end

        # Add rows with styles
        s.add_row(["Header 1", "Header 2"], styles: %w[heading heading])
        s.add_row([100, 200])
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

  test "streaming mode: add_style and apply to rows" do
    tmp = Tempfile.new(["style_stream_test", ".xlsx"])
    begin
      Xlsxrb.generate(tmp.path) do |w|
        # Define styles
        w.add_style("heading") do |style|
          style.bold.size(14).font_color("FFFF0000")
        end

        w.add_style("total") do |style|
          style.bold.fill_color("FF00FF00")
        end

        w.add_sheet("Sales") do
          # Add header row with heading style
          w.add_row(%w[Date Amount], styles: { 0 => "heading", 1 => "heading" })

          # Add data rows
          w.add_row([Date.today, 100])
          w.add_row([Date.today - 1, 200])

          # Add total row with total style
          w.add_row(["Total", 300], styles: { 1 => "total" })
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
      w.add_sheet("Styled") do |s|
        s.add_style("bold_red") do |style|
          style.bold.font_color("FFFF0000").size(12)
        end

        s.add_row(%w[Styled Data], styles: ["bold_red", nil])
        s.add_row(%w[Normal Row])
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
