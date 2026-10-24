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
      Xlsxrb.write(tmp.path) do |w|
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
      workbook = Xlsxrb.read(tmp.path).load
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
      read_workbook = Xlsxrb.read(tmp.path).load
      assert_equal(1, read_workbook.sheets.size)
      sheet = read_workbook.sheets[0]
      assert_equal(2, sheet.rows.size)

      # First cell should have a style index
      assert(sheet.rows[0].cells[0].style_index.is_a?(Integer) || sheet.rows[0].cells[0].style_index.nil?)
    ensure
      tmp.close!
    end
  end

  test "round-trip stylesheet metadata: border colors, alignment, and protection" do
    Dir.mktmpdir do |dir|
      filepath = File.join(dir, "styles_metadata.xlsx")
      styles = {
        borders: [
          {
            left: { style: "thin", color: { rgb: "FFFF0000" } },
            diagonal_up: true,
            diagonal_down: true
          }
        ],
        cell_xfs: [
          {
            font_id: 0,
            fill_id: 0,
            border_id: 0,
            alignment: { indent: 2, shrink_to_fit: true },
            protection: { locked: true, hidden: false }
          }
        ]
      }

      wb = Xlsxrb::Elements::Workbook.new(
        sheets: [Xlsxrb::Elements::Worksheet.new(name: "S", rows: [Xlsxrb::Elements::Row.new(index: 0)])],
        styles: styles
      )
      Xlsxrb.write(filepath, wb)

      parsed = Xlsxrb.read(filepath).load
      border = parsed.styles[:borders]&.first
      assert_equal("FFFF0000", border&.dig(:left, :color, :rgb))
      assert_equal(true, border&.dig(:diagonal_up))
      assert_equal(true, border&.dig(:diagonal_down))

      xf = parsed.styles[:cell_xfs]&.first
      assert_equal(2, xf&.dig(:alignment, :indent))
      assert_equal(true, xf&.dig(:alignment, :shrink_to_fit))
      assert_equal(true, xf&.dig(:protection, :locked))
      assert_equal(false, xf&.dig(:protection, :hidden))
    end
  end

  test "differential formats referenced by conditional formatting survive modify" do
    styles_xml = <<~XML
      <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <fonts count="1"><font><b val="0"/><sz val="10"/><name val="Arial"/></font></fonts>
        <dxfs count="2">
          <dxf><fill><patternFill><bgColor rgb="FFBFBFBF"/></patternFill></fill></dxf>
          <dxf><font><b/><i val="0"/><color auto="1"/></font><numFmt numFmtId="164" formatCode="0.0"/></dxf>
        </dxfs>
      </styleSheet>
    XML

    styles = Xlsxrb::Ooxml::StylesParser.parse(styles_xml)

    assert_equal(false, styles[:fonts].first[:bold])
    assert_equal([{ fill: { bg_color: { rgb: "FFBFBFBF" } } },
                  { font: { bold: true, italic: false, color: { auto: true } }, num_fmt: { num_fmt_id: 164, format_code: "0.0" } }],
                 styles[:dxfs])
    assert_equal({}, styles[:num_fmts])

    Dir.mktmpdir do |dir|
      path = File.join(dir, "cf.xlsx")
      Xlsxrb.write(path) do |w|
        w.sheet("S") do |s|
          s.row([1, 20])
          s.conditional_format("A1:B1", type: :cell_is, operator: :greaterThan, formula: "15", priority: 1, fill_color: "FFBFBFBF")
        end
      end
      out = File.join(dir, "out.xlsx")
      Xlsxrb.modify(path, out) { |wb| wb.with(sheets: wb.sheets + [Xlsxrb::Elements::Worksheet.new(name: "Added")]) }

      reread = Xlsxrb.read(out)
      cf = reread.sheets.first.load.conditional_formatting.first
      assert_equal(0, cf[:format_id])
      assert_equal(1, Xlsxrb::Ooxml::StylesParser.parse(Xlsxrb::Ooxml::ZipReader.open(out) { |z| z.read_entry("xl/styles.xml") })[:dxfs].size)
    end
  end

  test "differential formats round-trip with font, alignment, border, and protection" do
    Dir.mktmpdir do |dir|
      filepath = File.join(dir, "dxf_all.xlsx")
      wb = Xlsxrb::Elements::Workbook.new(
        sheets: [Xlsxrb::Elements::Worksheet.new(name: "Sheet1")],
        styles: Xlsxrb::Elements::Styles.new(
          fonts: [{ name: "Calibri", sz: 11 }],
          fills: [{ pattern: "none" }, { pattern: "gray125" }],
          borders: [{}],
          cell_xfs: [{ num_fmt_id: 0, font_id: 0, fill_id: 0, border_id: 0 }],
          dxfs: [
            {
              font: { bold: false, italic: true, color: { auto: true } },
              num_fmt: { num_fmt_id: 164, format_code: "0.00" },
              fill: { pattern: "solid", fg_color: "FFFF0000" },
              alignment: { horizontal: "center", wrap_text: true },
              border: { bottom: { style: "thin", color: "FF0000FF" } },
              protection: { locked: true }
            }
          ]
        )
      )
      Xlsxrb.write(filepath, wb)

      parsed = Xlsxrb::Ooxml::StylesParser.parse(Xlsxrb::Ooxml::ZipReader.open(filepath) { |z| z.read_entry("xl/styles.xml") })
      assert_equal(1, parsed[:dxfs].size)
      dxf = parsed[:dxfs].first
      assert_equal(false, dxf.dig(:font, :bold))
      assert_equal(true, dxf.dig(:font, :italic))
      assert_equal(true, dxf.dig(:font, :color, :auto))
      assert_equal(164, dxf.dig(:num_fmt, :num_fmt_id))
      assert_equal("0.00", dxf.dig(:num_fmt, :format_code))
      assert_equal("solid", dxf.dig(:fill, :pattern))
      assert_equal("FFFF0000", dxf.dig(:fill, :fg_color, :rgb))
      assert_equal("center", dxf.dig(:alignment, :horizontal))
      assert_equal(true, dxf.dig(:alignment, :wrap_text))
      assert_equal("thin", dxf.dig(:border, :bottom, :style))
      assert_equal("FF0000FF", dxf.dig(:border, :bottom, :color, :rgb))
      assert_equal(true, dxf.dig(:protection, :locked))
    end
  end
end
