# frozen_string_literal: true

require "test_helper"

class StreamSheetTest < Test::Unit::TestCase
  cover Xlsxrb::StreamSheet

  test "stream_sheet each_row, each_cell, and load into Elements::Worksheet" do
    xml = <<~XML
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
          <row r="1">
            <c r="A1"><v>100</v></c>
            <c r="B1" t="s"><v>0</v></c>
          </row>
          <row r="2">
            <c r="A2"><v>200</v></c>
          </row>
        </sheetData>
      </worksheet>
    XML

    sheet = Xlsxrb::StreamSheet.new("DataSheet", xml.b, ["Hello"])

    assert_equal("DataSheet", sheet.name)

    # Enumerators
    assert_instance_of(Enumerator, sheet.each_row)
    assert_instance_of(Enumerator, sheet.each_cell)

    # each_row enumeration
    rows = sheet.each_row.to_a
    assert_equal(2, rows.size)
    assert_equal(0, rows[0].index)
    assert_equal(1, rows[1].index)

    # each_cell enumeration across rows
    cells = sheet.each_cell.to_a
    assert_equal(3, cells.size)
    assert_equal(100, cells[0].value)
    assert_equal("Hello", cells[1].value)
    assert_equal(200, cells[2].value)

    # Enumerable each delegation
    assert_equal(2, sheet.count)
    assert_equal([0, 1], sheet.map(&:index))

    # load into in-memory Worksheet
    ws = sheet.load
    assert_instance_of(Xlsxrb::Elements::Worksheet, ws)
    assert_equal("DataSheet", ws.name)
    assert_equal(100, ws["A1"].value)
    assert_equal("Hello", ws["B1"].value)
    assert_equal(200, ws["A2"].value)

    # to_worksheet alias
    assert_equal(ws["A1"].value, sheet.to_worksheet["A1"].value)
  end

  test "stream_sheet feature accessors: merged_cells, auto_filter, data_validations, conditional_formats" do
    xml = <<~XML
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData/>
        <autoFilter ref="A1:E100"/>
        <mergeCells count="2">
          <mergeCell ref="A1:B2"/>
          <mergeCell ref="C3:D4"/>
        </mergeCells>
        <conditionalFormatting sqref="B2:B20">
          <cfRule type="cellIs" operator="greaterThan">
            <formula>50</formula>
          </cfRule>
        </conditionalFormatting>
        <dataValidations count="1">
          <dataValidation type="list" sqref="C2:C50" allowBlank="1">
            <formula1>"Option1,Option2"</formula1>
          </dataValidation>
        </dataValidations>
      </worksheet>
    XML

    sheet = Xlsxrb::StreamSheet.new("Features", xml.b, [])

    assert_equal(["A1:B2", "C3:D4"], sheet.merged_cells)
    assert_equal("A1:E100", sheet.auto_filter)

    assert_equal(1, sheet.data_validations.size)
    assert_equal("C2:C50", sheet.data_validations.first[:sqref])
    assert_equal("list", sheet.data_validations.first[:type])
    assert_equal(true, sheet.data_validations.first[:allow_blank])

    assert_equal(1, sheet.conditional_formats.size)
    assert_equal("B2:B20", sheet.conditional_formats.first[:sqref])
    assert_equal("cellIs", sheet.conditional_formats.first[:type])
    assert_equal("greaterThan", sheet.conditional_formats.first[:operator])
    assert_equal(["50"], sheet.conditional_formats.first[:formulas])
  end

  test "stream_sheet visibility state and propagation to in-memory Worksheet on load" do
    xml = <<~XML
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData/>
      </worksheet>
    XML

    default_sheet = Xlsxrb::StreamSheet.new("Visible", xml.b, [])
    assert_equal(:visible, default_sheet.state)
    assert_true(default_sheet.visible?)
    assert_false(default_sheet.hidden?)
    assert_equal(:visible, default_sheet.load.state)
    assert_true(default_sheet.load.visible?)

    hidden_sheet = Xlsxrb::StreamSheet.new("Hidden", xml.b, [], state: :hidden)
    assert_equal(:hidden, hidden_sheet.state)
    assert_false(hidden_sheet.visible?)
    assert_true(hidden_sheet.hidden?)
    assert_equal(:hidden, hidden_sheet.load.state)
    assert_true(hidden_sheet.load.hidden?)

    very_hidden_sheet = Xlsxrb::StreamSheet.new("VeryHidden", xml.b, [], state: :very_hidden)
    assert_equal(:very_hidden, very_hidden_sheet.state)
    assert_false(very_hidden_sheet.visible?)
    assert_true(very_hidden_sheet.hidden?)
    assert_equal(:very_hidden, very_hidden_sheet.load.state)
    assert_true(very_hidden_sheet.load.hidden?)
  end

  test "stream_sheet hyperlinks and hyperlink lookup" do
    xml = <<~XML
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
          <row r="1">
            <c r="A1"><v>Link</v></c>
          </row>
        </sheetData>
        <hyperlinks>
          <hyperlink ref="A1" location="Sheet2!A1" display="Go to Sheet2" tooltip="Click"/>
          <hyperlink ref="B2" location="Sheet1!C3"/>
        </hyperlinks>
      </worksheet>
    XML

    sheet = Xlsxrb::StreamSheet.new("Links", xml.b, [])
    assert_equal(2, sheet.hyperlinks.size)
    assert_equal({ location: "Sheet2!A1", display: "Go to Sheet2", tooltip: "Click" }, sheet.hyperlink("A1"))
    assert_equal({ location: "Sheet1!C3" }, sheet.hyperlink("B2"))
    assert_nil(sheet.hyperlink("C3"))

    ws = sheet.load
    assert_equal(2, ws.hyperlinks.size)
    assert_equal({ location: "Sheet2!A1", display: "Go to Sheet2", tooltip: "Click" }, ws.hyperlink("A1"))
    assert_true(ws["A1"].link?)
    assert_true(ws["A1"].hyperlink?)
    assert_equal("Sheet2!A1", ws["A1"].url)
    assert_equal({ location: "Sheet2!A1", display: "Go to Sheet2", tooltip: "Click" }, ws["A1"].hyperlink)
  end

  test "stream_sheet comments without zip_reader returns empty and nil lookup" do
    sheet = Xlsxrb::StreamSheet.new("NoComments", "<worksheet><sheetData/></worksheet>".b, [])
    assert_equal([], sheet.comments)
    assert_nil(sheet.comment("A1"))
    assert_equal({}, sheet.comments_by_ref)

    ws = sheet.load
    assert_equal([], ws.comments)
    assert_nil(ws.comment("A1"))
  end

  test "stream_sheet cell and formatted_value lookups and raw_values on StreamRow" do
    xml = <<~XML
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
          <row r="1">
            <c r="A1" s="1"><v>1042</v></c>
            <c r="B1"><f>A1*2</f><v>2084</v></c>
          </row>
        </sheetData>
      </worksheet>
    XML

    styles = {
      cell_xfs: [
        { num_fmt_id: 0 },
        { num_fmt_id: 4 } # #,##0.00
      ]
    }

    sheet = Xlsxrb::StreamSheet.new("Data", xml.b, [], styles, zip_reader: nil)
    assert_equal(styles, sheet.styles)

    cell_a1 = sheet.cell("A1")
    assert_not_nil(cell_a1)
    assert_equal("1042", cell_a1.raw_value)
    assert_equal("#,##0.00", cell_a1.format_code)
    assert_equal("1,042.00", cell_a1.formatted_value)

    assert_equal("1,042.00", sheet.formatted_value("A1"))
    assert_equal("1,042.00", sheet.formatted_value(0, 0))
    assert_equal("2084", sheet.formatted_value("B1"))
    assert_nil(sheet.formatted_value("C1"))

    row = sheet.each_row.first
    assert_equal([1042, 2084], row.values)
    assert_equal(%w[1042 2084], row.raw_values)
    assert_equal(["1,042.00", "2084"], row.formatted_values)
    assert_equal([nil, "A1*2"], row.formulas)
    assert_equal(row.raw_values, row[:raw_values])
    assert_equal(row.formatted_values, row[:formatted_values])
    assert_equal(row.formulas, row[:formulas])

    # load into in-memory Worksheet
    ws = sheet.load
    assert_equal(styles, ws.styles)
    assert_equal("1,042.00", ws["A1"].formatted_value)
    assert_equal("1,042.00", ws.formatted_value("A1"))
    assert_equal(%w[1042 2084], ws.rows.first.raw_values)
    assert_equal(["1,042.00", "2084"], ws.rows.first.formatted_values)
  end

  test "stream_sheet supports date1904 system flag" do
    xml = <<~XML
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
          <row r="1">
            <c r="A1"><v>0</v></c>
          </row>
        </sheetData>
      </worksheet>
    XML

    sheet1900 = Xlsxrb::StreamSheet.new("Sheet1900", xml.b, [], date1904: false)
    assert_equal(false, sheet1900.date1904?)
    row1900 = sheet1900.each_row.first
    assert_equal(false, row1900.date1904?)
    assert_equal(false, row1900[:date1904])
    assert_equal(false, row1900.cells.first.date1904?)
    assert_equal(Date.new(1899, 12, 31), row1900.cells.first.to_date)

    sheet1904 = Xlsxrb::StreamSheet.new("Sheet1904", xml.b, [], date1904: true)
    assert_equal(true, sheet1904.date1904?)
    row1904 = sheet1904.each_row.first
    assert_equal(true, row1904.date1904?)
    assert_equal(true, row1904[:date1904])
    assert_equal(true, row1904.cells.first.date1904?)
    assert_equal(Date.new(1904, 1, 1), row1904.cells.first.to_date)

    ws1904 = sheet1904.load
    assert_equal(true, ws1904.date1904?)
    assert_equal(Date.new(1904, 1, 1), ws1904["A1"].to_date)
  end

  test "stream_sheet exposes dimension and 1-based bounds" do
    xml = <<~XML
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <dimension ref="A1:Z50000"/>
        <sheetData>
          <row r="1"><c r="A1"><v>1</v></c></row>
        </sheetData>
      </worksheet>
    XML

    sheet = Xlsxrb::StreamSheet.new("Sheet1", xml.b, [])
    assert_equal("A1:Z50000", sheet.dimension)
    assert_equal(1, sheet.first_row)
    assert_equal(1, sheet.first_column)
    assert_equal(1, sheet.first_col)
    assert_equal(50_000, sheet.last_row)
    assert_equal(26, sheet.last_column)
    assert_equal(26, sheet.last_col)

    ws = sheet.load
    assert_equal("A1:Z50000", ws.dimension)
  end

  test "stream_sheet dimension bounds with non-A1 range and single cell" do
    xml_range = <<~XML
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <dimension ref="B2:D10"/>
        <sheetData/>
      </worksheet>
    XML
    sheet_range = Xlsxrb::StreamSheet.new("RangeSheet", xml_range.b, [])
    assert_equal("B2:D10", sheet_range.dimension)
    assert_equal(2, sheet_range.first_row)
    assert_equal(2, sheet_range.first_col)
    assert_equal(10, sheet_range.last_row)
    assert_equal(4, sheet_range.last_col)

    xml_single = <<~XML
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <dimension ref="C5"/>
        <sheetData/>
      </worksheet>
    XML
    sheet_single = Xlsxrb::StreamSheet.new("SingleSheet", xml_single.b, [])
    assert_equal("C5", sheet_single.dimension)
    assert_equal(5, sheet_single.first_row)
    assert_equal(3, sheet_single.first_col)
    assert_equal(5, sheet_single.last_row)
    assert_equal(3, sheet_single.last_col)
  end

  test "stream_sheet dimension handles missing and malformed tags" do
    xml_missing = <<~XML
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData/>
      </worksheet>
    XML
    sheet_missing = Xlsxrb::StreamSheet.new("Missing", xml_missing.b, [])
    assert_nil(sheet_missing.dimension)
    assert_nil(sheet_missing.first_row)
    assert_nil(sheet_missing.first_col)
    assert_nil(sheet_missing.last_row)
    assert_nil(sheet_missing.last_col)

    xml_malformed = <<~XML
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <dimension ref="NOT_A_VALID_REF"/>
        <sheetData/>
      </worksheet>
    XML
    sheet_malformed = Xlsxrb::StreamSheet.new("Malformed", xml_malformed.b, [])
    assert_equal("NOT_A_VALID_REF", sheet_malformed.dimension)
    assert_nil(sheet_malformed.first_row)
    assert_nil(sheet_malformed.last_row)

    # Explicit dimension keyword argument
    sheet_explicit = Xlsxrb::StreamSheet.new("Explicit", xml_missing.b, [], dimension: "X1:Y2")
    assert_equal("X1:Y2", sheet_explicit.dimension)
    assert_equal(1, sheet_explicit.first_row)
    assert_equal(24, sheet_explicit.first_col)
    assert_equal(2, sheet_explicit.last_row)
    assert_equal(25, sheet_explicit.last_col)
  end

  test "stream_sheet lazily extracts dimension from chunk supplier without reading later chunks" do
    chunks = [
      "<worksheet><dimension ref=\"A1:F30\"/><sheetData>",
      "<row r=\"1\"><c r=\"A1\"><v>val</v></c></row>",
      "</sheetData></worksheet>"
    ]
    chunks_read = 0
    supplier = lambda do |&blk|
      chunks.each do |c|
        chunks_read += 1
        blk.call(c)
      end
    end

    sheet = Xlsxrb::StreamSheet.new("Supplier", supplier, [])
    assert_equal("A1:F30", sheet.dimension)
    assert_equal(1, sheet.first_row)
    assert_equal(30, sheet.last_row)
    assert_equal(6, sheet.last_col)
    # Stopped on first chunk where dimension and sheetData appeared
    assert_equal(1, chunks_read)
  end

  test "stream_sheet each_row_values yields row value arrays directly without allocating Cell objects" do
    xml = <<~XML
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
          <row r="1">
            <c r="A1"><v>10</v></c>
            <c r="B1" t="s"><v>0</v></c>
          </row>
          <row r="2">
            <c r="A2" s="1"><v>44927</v></c>
            <c r="C2"><v>30</v></c>
          </row>
        </sheetData>
      </worksheet>
    XML

    styles = Xlsxrb::Elements::Styles.new({
                                            cell_xfs: [
                                              { num_fmt_id: 0 },
                                              { num_fmt_id: 14 }
                                            ]
                                          })

    sheet = Xlsxrb::StreamSheet.new("Data", xml.b, ["hello"], styles)

    enum = sheet.each_row_values
    assert_kind_of(Enumerator, enum)
    assert_equal([[10, "hello"], [44_927, nil, 30]], enum.to_a)

    collected = []
    sheet.each_row_values { |row_vals| collected << row_vals }
    assert_equal([[10, "hello"], [44_927, nil, 30]], collected)

    typecast_collected = []
    sheet.each_row_values(type_cast: true) { |row_vals| typecast_collected << row_vals }
    assert_equal([[10, "hello"], [Date.new(2023, 1, 1), nil, 30]], typecast_collected)
  end
end
