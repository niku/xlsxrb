# frozen_string_literal: true

require "test_helper"
require "stringio"

class OoxmlTest < Test::Unit::TestCase
  # --- ZipReader ---

  test "zip_reader reads entries from a real XLSX created by ZipWriter" do
    io = StringIO.new
    Xlsxrb::Ooxml::ZipWriter.open(io) do |w|
      w.add_entry("hello.txt", "Hello, World!")
      w.add_entry("nested/file.xml", "<root/>")
    end

    io.rewind
    reader = Xlsxrb::Ooxml::ZipReader.new(io)
    entries = reader.read_all

    assert_equal(2, entries.size)
    assert_equal("Hello, World!", entries["hello.txt"])
    assert_equal("<root/>", entries["nested/file.xml"])
  end

  test "zip_reader read_entry returns nil for missing entry" do
    io = StringIO.new
    Xlsxrb::Ooxml::ZipWriter.open(io) do |w|
      w.add_entry("a.txt", "data")
    end
    io.rewind

    reader = Xlsxrb::Ooxml::ZipReader.new(io)
    assert_nil(reader.read_entry("nonexistent.txt"))
  end

  test "zip_reader each_entry yields all entries" do
    io = StringIO.new
    Xlsxrb::Ooxml::ZipWriter.open(io) do |w|
      w.add_entry("one.txt", "1")
      w.add_entry("two.txt", "2")
    end
    io.rewind

    reader = Xlsxrb::Ooxml::ZipReader.new(io)
    names = []
    reader.each_entry { |name, _data| names << name }
    assert_equal(%w[one.txt two.txt], names.sort)
  end

  # --- ZipWriter ---

  test "zip_writer creates valid ZIP with entries" do
    io = StringIO.new
    Xlsxrb::Ooxml::ZipWriter.open(io) do |w|
      w.add_entry("test.txt", "content")
    end

    data = io.string
    assert_equal([0x50, 0x4B, 0x03, 0x04], data.bytes[0..3])
  end

  test "zip_writer streaming entry write" do
    io = StringIO.new
    Xlsxrb::Ooxml::ZipWriter.open(io) do |w|
      w.start_entry("stream.txt")
      w.write_data("Hello ")
      w.write_data("World!")
      w.finish_entry
    end

    io.rewind
    reader = Xlsxrb::Ooxml::ZipReader.new(io)
    assert_equal("Hello World!", reader.read_entry("stream.txt"))
  end

  # --- XmlBuilder ---

  test "xml_builder builds XML with tags and attributes" do
    io = StringIO.new
    b = Xlsxrb::Ooxml::XmlBuilder.new(io)
    b.declaration
    b.tag("root", { xmlns: "http://example.com" }) do |_|
      b.tag("child", { id: "1" }) { |_| b.text("hello") }
      b.empty_tag("empty", { flag: "true" })
    end

    xml = io.string
    assert_include(xml, '<?xml version="1.0"')
    assert_include(xml, '<root xmlns="http://example.com">')
    assert_include(xml, '<child id="1">hello</child>')
    assert_include(xml, '<empty flag="true"/>')
    assert_include(xml, "</root>")
  end

  test "xml_builder escapes special characters" do
    io = StringIO.new
    b = Xlsxrb::Ooxml::XmlBuilder.new(io)
    b.tag("t") { |_| b.text("a < b & c > d") }

    assert_equal("<t>a &lt; b &amp; c &gt; d</t>", io.string)
  end

  test "xml_builder write_unmapped restores unknown elements" do
    io = StringIO.new
    b = Xlsxrb::Ooxml::XmlBuilder.new(io)
    node = { tag: "custom", attrs: { "val" => "42" }, children: [], text: "data" }
    b.write_unmapped(node)

    assert_equal('<custom val="42">data</custom>', io.string)
  end

  # --- SharedStringsParser ---

  test "shared_strings_parser parses shared strings" do
    xml = <<~XML
      <?xml version="1.0" encoding="UTF-8"?>
      <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
        <si><t>Hello</t></si>
        <si><t>World</t></si>
        <si><t>Test</t></si>
      </sst>
    XML

    strings = Xlsxrb::Ooxml::SharedStringsParser.parse(xml)
    assert_equal(%w[Hello World Test], strings)
  end

  test "shared_strings_parser returns empty array for nil input" do
    assert_equal([], Xlsxrb::Ooxml::SharedStringsParser.parse(nil))
  end

  test "shared_strings_parser handles rich text" do
    xml = <<~XML
      <?xml version="1.0" encoding="UTF-8"?>
      <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
        <si><r><t>Hello </t></r><r><t>World</t></r></si>
      </sst>
    XML

    strings = Xlsxrb::Ooxml::SharedStringsParser.parse(xml)
    assert_equal(["Hello World"], strings)
  end

  # --- StylesParser ---

  test "styles_parser parses number formats" do
    xml = <<~XML
      <?xml version="1.0" encoding="UTF-8"?>
      <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <numFmts count="1">
          <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
        </numFmts>
      </styleSheet>
    XML

    result = Xlsxrb::Ooxml::StylesParser.parse(xml)
    assert_equal({ 164 => "yyyy-mm-dd" }, result[:num_fmts])
  end

  test "styles_parser parses cellXfs" do
    xml = <<~XML
      <?xml version="1.0" encoding="UTF-8"?>
      <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <cellXfs count="1">
          <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        </cellXfs>
      </styleSheet>
    XML

    result = Xlsxrb::Ooxml::StylesParser.parse(xml)
    assert_equal(1, result[:cell_xfs].size)
    assert_equal(0, result[:cell_xfs][0][:num_fmt_id])
  end

  test "styles_parser returns empty hash for nil input" do
    assert_equal({}, Xlsxrb::Ooxml::StylesParser.parse(nil))
  end

  # --- WorkbookParser ---

  test "workbook_parser parses sheet list" do
    xml = <<~XML
      <?xml version="1.0" encoding="UTF-8"?>
      <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheets>
          <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
          <sheet name="Data" sheetId="2" r:id="rId2"/>
        </sheets>
      </workbook>
    XML

    sheets = Xlsxrb::Ooxml::WorkbookParser.parse(xml)
    assert_equal(2, sheets.size)
    assert_equal("Sheet1", sheets[0][:name])
    assert_equal("rId1", sheets[0][:r_id])
    assert_equal("Data", sheets[1][:name])
  end

  # --- RelationshipsParser ---

  test "relationships_parser parses relationships" do
    xml = <<~XML
      <?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
      </Relationships>
    XML

    rels = Xlsxrb::Ooxml::RelationshipsParser.parse(xml)
    assert_equal("worksheets/sheet1.xml", rels["rId1"])
    assert_equal("styles.xml", rels["rId2"])
  end

  # --- WorksheetParser ---

  test "worksheet_parser parses rows and cells" do
    xml = <<~XML
      <?xml version="1.0" encoding="UTF-8"?>
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
          <row r="1">
            <c r="A1" t="s"><v>0</v></c>
            <c r="B1"><v>42</v></c>
          </row>
          <row r="2">
            <c r="A2" t="b"><v>1</v></c>
          </row>
        </sheetData>
      </worksheet>
    XML

    shared_strings = ["Hello"]
    rows = Xlsxrb::Ooxml::WorksheetParser.parse(xml, shared_strings: shared_strings)

    assert_equal(2, rows.size)
    assert_equal(0, rows[0][:index])
    assert_equal("Hello", rows[0][:cells][0][:value])
    assert_equal(42, rows[0][:cells][1][:value])
    assert_equal(true, rows[1][:cells][0][:value])
  end

  test "worksheet_parser streaming each_row yields rows" do
    xml = <<~XML
      <?xml version="1.0" encoding="UTF-8"?>
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
          <row r="1"><c r="A1"><v>1</v></c></row>
          <row r="2"><c r="A2"><v>2</v></c></row>
          <row r="3"><c r="A3"><v>3</v></c></row>
        </sheetData>
      </worksheet>
    XML

    collected = []
    Xlsxrb::Ooxml::WorksheetParser.each_row(xml, shared_strings: []) do |row|
      collected << row[:index]
    end

    assert_equal([0, 1, 2], collected)
  end

  test "worksheet_parser parses columns" do
    xml = <<~XML
      <?xml version="1.0" encoding="UTF-8"?>
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <cols>
          <col min="1" max="1" width="20.5" customWidth="1"/>
          <col min="2" max="3" width="10.0" hidden="1"/>
        </cols>
        <sheetData/>
      </worksheet>
    XML

    columns = Xlsxrb::Ooxml::WorksheetParser.parse_columns(xml)
    assert_equal(2, columns.size)
    assert_equal(1, columns[0][:min])
    assert_in_delta(20.5, columns[0][:width])
    assert_equal(true, columns[1][:hidden])
  end

  test "worksheet_parser parses formula cells" do
    xml = <<~XML
      <?xml version="1.0" encoding="UTF-8"?>
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
          <row r="1">
            <c r="A1"><f>SUM(B1:B10)</f><v>55</v></c>
          </row>
        </sheetData>
      </worksheet>
    XML

    rows = Xlsxrb::Ooxml::WorksheetParser.parse(xml, shared_strings: [])
    cell = rows[0][:cells][0]
    assert_equal("SUM(B1:B10)", cell[:formula])
    assert_equal(55, cell[:value])
  end

  # --- WorksheetWriter ---

  test "worksheet_writer generates valid worksheet XML" do
    io = StringIO.new
    writer = Xlsxrb::Ooxml::WorksheetWriter.new(io)
    writer.start
    writer.write_row(0, [
                       { ref: "A1", value: 0, type: "s" },
                       { ref: "B1", value: 42 }
                     ])
    writer.finish

    xml = io.string
    assert_include(xml, "<worksheet")
    assert_include(xml, "<sheetData>")
    assert_include(xml, '<row r="1">')
    assert_include(xml, '<c r="A1" t="s">')
    assert_include(xml, "<v>0</v>")
    assert_include(xml, '<c r="B1">')
    assert_include(xml, "<v>42</v>")
    assert_include(xml, "</sheetData>")
    assert_include(xml, "</worksheet>")
  end

  test "worksheet_writer write_row_values serializes styled shared-string cells" do
    io = StringIO.new
    writer = Xlsxrb::Ooxml::WorksheetWriter.new(io)
    sst = []
    sst_index = {}

    writer.start
    writer.write_row_values(0, ["name", 10, nil], styles: { 0 => "header", 2 => "header" }, style_map: { "header" => 3 }, sst: sst, sst_index: sst_index)
    writer.finish

    xml = io.string
    assert_include(xml, '<row r="1">')
    assert_include(xml, '<c r="A1" s="3" t="s"><v>0</v></c>')
    assert_include(xml, '<c r="B1"><v>10</v></c>')
    assert_include(xml, '<c r="C1" s="3"/>')
    assert_equal(["name"], sst)
    assert_equal(0, sst_index["name"])
  end
end
