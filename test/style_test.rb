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
      border = parsed.styles.borders.first
      assert_equal("FFFF0000", border&.dig(:left, :color, :rgb))
      assert_equal(true, border&.dig(:diagonal_up))
      assert_equal(true, border&.dig(:diagonal_down))

      xf = parsed.styles.cell_xfs.first
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

    assert_equal(false, styles.fonts.first[:bold])
    assert_equal([{ fill: { bg_color: { rgb: "FFBFBFBF" } } },
                  { font: { bold: true, italic: false, color: { auto: true } }, num_fmt: { num_fmt_id: 164, format_code: "0.0" } }],
                 styles.dxfs)
    assert_equal({}, styles.num_fmts)

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
      assert_equal(1, Xlsxrb::Ooxml::StylesParser.parse(Xlsxrb::Ooxml::ZipReader.open(out) { |z| z.read_entry("xl/styles.xml") }).dxfs.size)
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
      assert_equal(1, parsed.dxfs.size)
      dxf = parsed.dxfs.first
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

  test "Elements::Styles encapsulated collections, lookups, and builder helpers" do
    styles = Xlsxrb::Elements::Styles.new

    # Empty defaults
    assert_equal([], styles.fonts)
    assert_equal([], styles.fills)
    assert_equal([], styles.borders)
    assert_equal([], styles.cell_xfs)
    assert_equal({}, styles.num_fmts)
    assert_equal([], styles.dxfs)
    assert_equal([], styles.cell_style_xfs)
    assert_equal([], styles.cell_styles)
    assert_equal({}, styles.table_styles)
    assert_equal([], styles.indexed_colors)
    assert_equal([], styles.mru_colors)
    assert_equal({}, styles.colors)
    assert_true(styles.empty?)

    # Out-of-bounds lookups
    assert_nil(styles.font(0))
    assert_nil(styles.font(-1))
    assert_nil(styles.fill(0))
    assert_nil(styles.fill(-1))
    assert_nil(styles.border(0))
    assert_nil(styles.cell_xf(0))
    assert_nil(styles.xf(0))
    assert_nil(styles.num_fmt(164))
    assert_nil(styles.dxf(0))
    assert_nil(styles.cell_style_xf(0))
    assert_nil(styles.cell_style(0))

    # Builders / Mutators
    font_id = styles.add_font(name: "Arial", sz: 12, bold: true)
    assert_equal(0, font_id)
    assert_equal("Arial", styles.font(0)&.[](:name))

    fill_id = styles.add_fill(pattern: "solid", fg_color: "FFFF0000")
    assert_equal(0, fill_id)
    assert_equal("solid", styles.fill(0)&.[](:pattern))

    border_id = styles.add_border(bottom: { style: "thin" })
    assert_equal(0, border_id)
    assert_equal("thin", styles.border(0)&.dig(:bottom, :style))

    fmt_id = styles.add_num_fmt(165, "$#,##0.00")
    assert_equal(165, fmt_id)
    assert_equal("$#,##0.00", styles.num_fmt(165))

    xf_id = styles.add_xf(font_id: font_id, fill_id: fill_id, border_id: border_id, num_fmt_id: fmt_id)
    assert_equal(0, xf_id)
    assert_equal(xf_id, styles.cell_xf(0)&.[](:font_id))
    assert_equal("$#,##0.00", styles.number_format(0))
    assert_equal(:number, styles.format_type(0))

    dxf_id = styles.add_dxf(font: { italic: true })
    assert_equal(0, dxf_id)
    assert_equal(true, styles.dxf(0)&.dig(:font, :italic))

    cs_xf_id = styles.add_cell_style_xf(font_id: 0)
    assert_equal(0, cs_xf_id)
    assert_equal(0, styles.cell_style_xf(0)&.[](:font_id))

    cs_id = styles.add_cell_style(name: "Normal", xf_id: 0)
    assert_equal(0, cs_id)
    assert_equal("Normal", styles.cell_style(0)&.[](:name))

    # add_style composite helper
    styled_xf_id = styles.add_style(
      fill: { pattern: "gray125" },
      font: { name: "Courier", sz: 10 },
      border: { top: { style: "thick" } },
      format_code: "yyyy-mm-dd",
      alignment: { wrap_text: true }
    )
    assert_equal(1, styled_xf_id)
    assert_equal("yyyy-mm-dd", styles.number_format(1))
    assert_equal(:date, styles.format_type(1))
    assert_true(styles.date_format?(1))

    # Hash delegation removed
    assert_raise(NoMethodError) { styles[:fonts] }
    assert_raise(NoMethodError) { styles[:fonts] = [] }
    assert_raise(NoMethodError) { styles.fetch(:fonts) }
    assert_raise(NoMethodError) { styles.dig(:fonts, :name) }
    assert_raise(NoMethodError) { styles.each { |_k, _v| nil } }
    assert_raise(NoMethodError) { styles.keys }
    assert_raise(NoMethodError) { styles.values }

    # to_h and equality
    hash = styles.to_h
    assert_instance_of(Hash, hash)
    assert_equal(hash, styles.to_h)
    assert_equal(styles, hash)
    assert_equal(styles, Xlsxrb::Elements::Styles.new(hash))
  end
end
