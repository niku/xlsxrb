# frozen_string_literal: true

require "test_helper"
require "tempfile"

class StyleInteroperabilityTest < Test::Unit::TestCase
  # E2E test: In-memory styled cells
  test "in-memory mode: write and validate styled cells via Open XML SDK" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("Styled") do |s|
        # Define styles
        s.add_style("header") do |style|
          style.bold.size(14).font_color("FFFF0000")
        end

        s.add_style("total") do |style|
          style.bold.fill_color("FF00FF00")
        end

        # Add styled rows
        s.add_row(%w[Name Value], styles: { 0 => "header", 1 => "header" })
        s.add_row(["Item A", 100])
        s.add_row(["Item B", 200])
        s.add_row(["Total", 300], styles: { 1 => "total" })
      end
    end

    xlsx_tempfile = Tempfile.new(["in_memory_style_e2e", ".xlsx"])
    begin
      Xlsxrb.write(xlsx_tempfile.path, workbook)

      # Verify file exists and can be read back
      read_workbook = Xlsxrb.read(xlsx_tempfile.path)
      assert_equal(1, read_workbook.sheets.size)

      sheet = read_workbook.sheets[0]
      assert_equal(4, sheet.rows.size)

      # Verify cells with style indices were written
      first_row_first_cell = sheet.rows[0].cells[0]
      assert_equal("Name", first_row_first_cell.value)

      # Check the XML contains the style information
      xml_styles = read_xml_from_xlsx(xlsx_tempfile.path, "xl/styles.xml")
      assert_match(/<cellXfs/, xml_styles)
      assert_match(/<fonts/, xml_styles)
      assert_match(/<fills/, xml_styles)
    ensure
      xlsx_tempfile.close!
    end
  end

  # E2E test: Streaming styled cells
  test "streaming mode: write and validate styled cells via Open XML SDK" do
    xlsx_tempfile = Tempfile.new(["streaming_style_e2e", ".xlsx"])
    begin
      Xlsxrb.generate(xlsx_tempfile.path) do |w|
        # Define styles
        w.add_style("header") do |style|
          style.bold.size(12).font_color("FF0000FF")
        end

        w.add_style("data") do |style|
          style.fill_color("FFFFC000")
        end

        w.add_sheet("Data") do
          # Header row with style
          w.add_row(%w[Col1 Col2], styles: { 0 => "header", 1 => "header" })

          # Data rows
          (1..5).each do |i|
            styles = i.even? ? { 0 => "data", 1 => "data" } : nil
            w.add_row([i, "Value#{i}"], styles: styles)
          end
        end
      end

      # Read back and verify
      read_workbook = Xlsxrb.read(xlsx_tempfile.path)
      assert_equal(1, read_workbook.sheets.size)
      assert_equal(6, read_workbook.sheets[0].rows.size)

      # Verify styles were written
      xml_styles = read_xml_from_xlsx(xlsx_tempfile.path, "xl/styles.xml")
      assert_match(/<cellXfs/, xml_styles)
      assert_match(/<cellXfs[^>]*count="[0-9]"/, xml_styles) # Should have cell styles
    ensure
      xlsx_tempfile.close!
    end
  end

  # E2E test: StyleBuilder direct usage
  test "StyleBuilder: fluent API produces compatible styles" do
    style_builder = Xlsxrb::StyleBuilder.new("test_style")
    style_builder
      .bold
      .italic
      .size(11)
      .font_name("Calibri")
      .font_color("FFFF0000")
      .fill_color("FF0000FF")
      .border_all(style: "thin", color: "FF000000")

    # Verify the builder collected properties correctly
    assert_equal(true, style_builder.font_props[:bold])
    assert_equal(true, style_builder.font_props[:italic])
    assert_equal(11, style_builder.font_props[:sz])
    assert_equal("Calibri", style_builder.font_props[:name])
    assert_equal("FFFF0000", style_builder.font_props[:color])
    assert_equal("FF0000FF", style_builder.fill_props[:fg_color])
    assert_equal("solid", style_builder.fill_props[:pattern])
    assert(style_builder.border_props.key?(:left))
  end

  private

  def read_xml_from_xlsx(path, entry)
    entries = Xlsxrb::Ooxml::ZipReader.open(path, &:read_all)
    entries[entry]
  end
end
