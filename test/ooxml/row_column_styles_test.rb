# frozen_string_literal: true

require "test_helper"

class RowColumnStylesTest < Test::Unit::TestCase
  def test_row_and_column_styles_roundtrip
    Dir.mktmpdir do |dir|
      filepath = File.join(dir, "test.xlsx")
      # Write a workbook with row styles, column styles, and border color
      styles = {
        fonts: [{ name: "Calibri", sz: 11 }, { name: "Arial", sz: 14, bold: true }],
        fills: [{ pattern: "none" }, { pattern: "gray125" }, { pattern: "solid", fg_color: "FFFFFF00" }],
        borders: [{ left: { style: "thin", color: "FFFF0000" } }],
        cell_xfs: [
          { font_id: 0, fill_id: 0, border_id: 0 },
          { font_id: 1, fill_id: 2, border_id: 0 }
        ]
      }

      wb = Xlsxrb::Elements::Workbook.new(
        sheets: [
          Xlsxrb::Elements::Worksheet.new(
            name: "StyledSheet",
            rows: [
              Xlsxrb::Elements::Row.new(
                index: 0,
                cells: [Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "Hello")],
                unmapped_data: { style_index: 1 }
              )
            ],
            columns: [
              Xlsxrb::Elements::Column.new(
                index: 0,
                width: 20.0,
                unmapped_data: { style_index: 1 }
              )
            ]
          )
        ],
        styles: styles
      )

      Xlsxrb.write(filepath, wb)

      # Read back and verify styles, row style_index, and column style_index
      parsed_wb = Xlsxrb.read(filepath).load
      parsed_ws = parsed_wb.sheets.first
      refute_nil parsed_ws

      # Verify row style_index
      assert_equal(1, parsed_ws.rows.first.style_index)

      # Verify column style_index
      assert_equal(1, parsed_ws.columns.first.style_index)

      # Verify border color parsed
      parsed_styles = parsed_wb.styles
      refute_nil parsed_styles
      border = parsed_styles.borders.first
      assert_equal("FFFF0000", border&.dig(:left, :color, :rgb))
    end
  end

  def test_workbook_writer_num_fmts_hash_and_array_support
    Dir.mktmpdir do |dir|
      # Test with Hash format (as returned by Reader)
      filepath_hash = File.join(dir, "hash_styles.xlsx")
      wb_hash = Xlsxrb::Elements::Workbook.new(
        sheets: [Xlsxrb::Elements::Worksheet.new(name: "S", rows: [Xlsxrb::Elements::Row.new(index: 0)])],
        styles: { num_fmts: { 164 => "#,##0.00" } }
      )
      Xlsxrb.write(filepath_hash, wb_hash)

      entries = Xlsxrb::Ooxml::ZipReader.open(filepath_hash, &:read_all)
      styles_xml = entries["xl/styles.xml"]
      assert_match(%r{<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0\.00"/></numFmts>}, styles_xml)

      parsed = Xlsxrb.read(filepath_hash).load
      assert_equal("#,##0.00", parsed.styles.dig(:num_fmts, 164))

      # Test with Array format
      filepath_arr = File.join(dir, "arr_styles.xlsx")
      wb_arr = Xlsxrb::Elements::Workbook.new(
        sheets: [Xlsxrb::Elements::Worksheet.new(name: "S", rows: [Xlsxrb::Elements::Row.new(index: 0)])],
        styles: { num_fmts: [{ num_fmt_id: 165, format_code: "$#,##0" }] }
      )
      Xlsxrb.write(filepath_arr, wb_arr)

      entries_arr = Xlsxrb::Ooxml::ZipReader.open(filepath_arr, &:read_all)
      styles_xml_arr = entries_arr["xl/styles.xml"]
      assert_match(%r{<numFmts count="1"><numFmt numFmtId="165" formatCode="\$#,##0"/></numFmts>}, styles_xml_arr)
    end
  end

  def test_cascade_style_inheritance_cell_row_column
    # Test WorksheetWriter cascades: Cell Style > Row Style > Column Style
    io = StringIO.new
    writer = Xlsxrb::Ooxml::WorksheetWriter.new(io)
    writer.start(columns: [
                   { index: 0, style_index: 10 }, # col A has style 10
                   { index: 1, style_index: 20 }, # col B has style 20
                   { index: 2, style_index: 30 }  # col C has style 30
                 ])
    # Row 0: Row style 100
    # Cell A1: unstyled -> gets row style 100
    # Cell B1: explicit style 5 -> keeps style 5
    # Cell C1: unstyled -> gets row style 100
    writer.write_row(0, [
                       Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "A1"),
                       Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 1, value: "B1", style_index: 5),
                       Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 2, value: "C1")
                     ], attrs: { style_index: 100 })

    # Row 1: No row style (via write_row_values)
    # Col A (idx 0): unstyled cell -> gets col style 10
    # Col B (idx 1): explicit cell style 7 -> keeps style 7
    # Col C (idx 2): unstyled cell -> gets col style 30
    writer.write_row_values(1, %w[A2 B2 C2], styles: [nil, 7, nil])
    writer.finish

    xml = io.string
    assert_match(%r{<c r="A1" s="100" t="inlineStr"><is><t>A1</t></is></c>}, xml)
    assert_match(%r{<c r="B1" s="5" t="inlineStr"><is><t>B1</t></is></c>}, xml)
    assert_match(%r{<c r="C1" s="100" t="inlineStr"><is><t>C1</t></is></c>}, xml)

    assert_match(%r{<c r="A2" s="10" t="inlineStr"><is><t>A2</t></is></c>}, xml)
    assert_match(%r{<c r="B2" s="7" t="inlineStr"><is><t>B2</t></is></c>}, xml)
    assert_match(%r{<c r="C2" s="30" t="inlineStr"><is><t>C2</t></is></c>}, xml)
  end

  def test_stream_writer_and_worksheet_builder_column_style
    # Test StreamWriter with column style
    sw_io = StringIO.new
    Xlsxrb.write(sw_io) do |w|
      w.style(:currency, number_format: "$#,##0.00")
      w.sheet("Sales") do |s|
        s.column("B", width: 15.0, style: :currency)
        s.row(["Item", 1234.5])
      end
    end

    parsed = Xlsxrb.read(StringIO.new(sw_io.string)).load
    sheet = parsed.sheets.first
    currency_num_fmt_id = parsed.styles.num_fmts.key("$#,##0.00")
    currency_id = parsed.styles.cell_xfs.find_index { |xf| xf[:num_fmt_id] == currency_num_fmt_id }
    assert_not_nil currency_id
    assert_equal currency_id, sheet["B1"].style_index

    # Test WorksheetBuilder with column style
    wb_built = Xlsxrb.build do |wb|
      wb.sheet("DOM") do |ws|
        ws.style(:highlight, font: { bold: true })
        ws.column(1, style: :highlight)
        ws.row(%w[Normal Highlighted])
      end
    end

    dom_sheet = wb_built.sheets.first
    assert_equal 1, dom_sheet.columns.first.style_index
    assert_equal 1, dom_sheet["B1"].style_index

    wb_io = StringIO.new
    Xlsxrb.write(wb_io, wb_built)
    parsed_dom = Xlsxrb.read(StringIO.new(wb_io.string)).load
    assert_equal 1, parsed_dom.sheets.first.columns.first.style_index
    assert_equal 1, parsed_dom.sheets.first["B1"].style_index
  end

  def test_unresolved_style_name_does_not_emit_invalid_string_attribute
    io = StringIO.new
    writer = Xlsxrb::Ooxml::WorksheetWriter.new(io)
    writer.start
    # "unknown" string and :nonexistent symbol are not in style_map and not Integer
    writer.write_row_values(0, %w[A1 B1 C1], styles: ["unknown", :nonexistent, 5], style_map: { "valid" => 1 })
    writer.finish

    xml = io.string
    # A1 ("unknown") and B1 (:nonexistent) must not have s="..."
    assert_no_match(/<c r="A1"[^>]* s="unknown"/, xml)
    assert_no_match(/<c r="B1"[^>]* s="nonexistent"/, xml)
    # C1 (raw integer 5) must have s="5"
    assert_match(/<c r="C1"[^>]* s="5"/, xml)
  end

  def test_column_and_row_style_inheritance_dom_and_streaming
    # 1. DOM builder
    wb = Xlsxrb.build do |b|
      b.sheet("Inheritance") do |ws|
        ws.style(:currency, number_format: "$#,##0.00")
        ws.style(:row_hl, font: { bold: true })
        ws.style(:cell_override, font: { color: "FF0000" })

        # Col 0 has named style :currency, Col 1 has inline style { italic: true }
        ws.column(0, style: :currency)
        ws.column(1, style: { italic: true })

        # Row 0: no row style. Cell A1 unstyled -> inherits Col 0. Cell B1 unstyled -> inherits Col 1.
        ws.row([100, "ItalicText"])

        # Row 1: row style :row_hl.
        # Cell A2 unstyled -> inherits Row 1 (:row_hl) over Col 0.
        # Cell B2 explicit :cell_override -> overrides Row 1 and Col 1.
        ws.row([200, "OverrideText"], styles: [nil, :cell_override], style: :row_hl)

        # Row 2: inline row style { fill_color: "FFFF00" }
        ws.row([300, "RowInline"], style: { fill_color: "FFFF00" })
      end
    end

    dom_sheet = wb.sheets.first
    currency_id = dom_sheet.columns[0].style_index
    col1_inline_id = dom_sheet.columns[1].style_index
    assert_kind_of Integer, currency_id
    assert_kind_of Integer, col1_inline_id

    # Cell A1 inherits currency_id and format_code
    assert_equal currency_id, dom_sheet["A1"].style_index
    assert_equal "$#,##0.00", dom_sheet["A1"].format_code

    # Cell B1 inherits col1_inline_id
    assert_equal col1_inline_id, dom_sheet["B1"].style_index

    # Row 1 has resolved row style
    row1_style_id = dom_sheet.rows[1].style_index
    assert_kind_of Integer, row1_style_id

    # Cell A2 inherits row style over column style
    assert_equal row1_style_id, dom_sheet["A2"].style_index

    # Cell B2 keeps explicit cell style
    refute_equal row1_style_id, dom_sheet["B2"].style_index
    refute_equal col1_inline_id, dom_sheet["B2"].style_index

    # Serialize and check XML for row and col style attributes
    io = StringIO.new
    Xlsxrb.write(io, wb)
    entries = Xlsxrb::Ooxml::ZipReader.open(StringIO.new(io.string), &:read_all)
    ws_xml = entries["xl/worksheets/sheet1.xml"]

    # Verify <col style="..."> uses integer IDs
    assert_match(/<col[^>]*style="#{currency_id}"/, ws_xml)
    assert_match(/<col[^>]*style="#{col1_inline_id}"/, ws_xml)

    # Verify <row s="..." customFormat="1"> is emitted on Row 2
    assert_match(/<row r="2"[^>]*s="#{row1_style_id}" customFormat="1"/, ws_xml)

    # Roundtrip load
    parsed = Xlsxrb.read(StringIO.new(io.string)).load
    parsed_sheet = parsed.sheets.first
    assert_equal currency_id, parsed_sheet["A1"].style_index
    assert_equal col1_inline_id, parsed_sheet["B1"].style_index
    assert_equal row1_style_id, parsed_sheet["A2"].style_index

    # 2. Streaming builder
    sw_io = StringIO.new
    Xlsxrb.write(sw_io) do |w|
      w.style(:header, font: { bold: true })
      w.sheet("Stream") do |s|
        s.column(0, width: 20.0, style: { italic: true })
        s.row(%w[ColInherited Normal])
        s.row(%w[RowInherited RowInherited2], style: :header)
      end
    end

    stream_entries = Xlsxrb::Ooxml::ZipReader.open(StringIO.new(sw_io.string), &:read_all)
    stream_ws_xml = stream_entries["xl/worksheets/sheet1.xml"]
    assert_match(/<row r="2"[^>]*s="\d+" customFormat="1"/, stream_ws_xml)

    parsed_stream = Xlsxrb.read(StringIO.new(sw_io.string)).load
    ps = parsed_stream.sheets.first
    assert_kind_of Integer, ps.columns.first.style_index
    assert_equal ps.columns.first.style_index, ps["A1"].style_index
    assert_equal ps.rows[1].style_index, ps["A2"].style_index
  end
end
