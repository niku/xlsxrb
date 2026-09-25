# frozen_string_literal: true

require "test_helper"
require "stringio"

class OoxmlTest < Test::Unit::TestCase
  cover Xlsxrb::Ooxml::XmlBuilder
  cover Xlsxrb::Ooxml::XmlParser::BaseListener
  cover Xlsxrb::Ooxml::SharedStringsParser
  cover Xlsxrb::Ooxml::StylesParser
  cover Xlsxrb::Ooxml::WorkbookParser
  cover Xlsxrb::Ooxml::RelationshipsParser
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

  test "zip_reader each_entry_chunk streams entry uncompressed chunks" do
    io = StringIO.new
    content = "A" * 150_000
    Xlsxrb::Ooxml::ZipWriter.open(io) do |w|
      w.add_entry("large.txt", content)
    end
    io.rewind

    reader = Xlsxrb::Ooxml::ZipReader.new(io)
    assert_equal(true, reader.entry?("large.txt"))
    assert_equal(false, reader.entry?("missing.txt"))
    assert_equal(["large.txt"], reader.entry_names)

    chunks = []
    reader.each_entry_chunk("large.txt", chunk_size: 32_768) do |chunk|
      chunks << chunk
    end
    assert_operator chunks.size, :>, 1
    assert_equal(content, chunks.join)

    # Test open_entry_io
    stream = reader.open_entry_io("large.txt")
    part1 = stream.read(5000)
    part2 = stream.read(10_000)
    rest = stream.read
    stream.close
    assert_equal("A" * 5000, part1)
    assert_equal("A" * 10_000, part2)
    assert_equal("A" * (150_000 - 15_000), rest)
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

  test "xml_builder escapes special characters in text and attributes" do
    io = StringIO.new
    b = Xlsxrb::Ooxml::XmlBuilder.new(io)
    attrs = { nil_before: nil, quote: 'He said "Hello" & smiled', nil_middle: nil, valid: "yes", nil_after: nil }
    b.tag("t", attrs) do |_|
      b.text("Tom & Jerry's <Great \"Show\">")
    end

    expected = '<t quote="He said &quot;Hello&quot; &amp; smiled" valid="yes">Tom &amp; Jerry&apos;s &lt;Great &quot;Show&quot;&gt;</t>'
    assert_equal(expected, io.string)
    assert_equal(expected, b.to_s)
  end

  test "XmlBuilder.escape correctly escapes XML special characters" do
    assert_equal("Tom &amp; Jerry", Xlsxrb::Ooxml::XmlBuilder.escape("Tom & Jerry"))
    assert_equal("&lt;tag&gt;", Xlsxrb::Ooxml::XmlBuilder.escape("<tag>"))
    assert_equal("&quot;quote&quot; and &apos;single&apos;", Xlsxrb::Ooxml::XmlBuilder.escape('"quote" and \'single\''))
    plain = "no_special_chars"
    assert_same(plain, Xlsxrb::Ooxml::XmlBuilder.escape(plain))
    assert_equal("123", Xlsxrb::Ooxml::XmlBuilder.escape(123))
    empty = ""
    assert_same(empty, Xlsxrb::Ooxml::XmlBuilder.escape(empty))
  end

  test "XmlBuilder.unescape correctly unescapes XML entities" do
    assert_equal("Tom & Jerry", Xlsxrb::Ooxml::XmlBuilder.unescape("Tom &amp; Jerry"))
    assert_equal("<tag>", Xlsxrb::Ooxml::XmlBuilder.unescape("&lt;tag&gt;"))
    assert_equal('"quote" and \'single\'', Xlsxrb::Ooxml::XmlBuilder.unescape("&quot;quote&quot; and &apos;single&apos;"))
    plain = "no_special_chars"
    assert_same(plain, Xlsxrb::Ooxml::XmlBuilder.unescape(plain))
    assert_equal("123", Xlsxrb::Ooxml::XmlBuilder.unescape(123))
    empty = ""
    assert_same(empty, Xlsxrb::Ooxml::XmlBuilder.unescape(empty))
    assert_equal("&unknown;", Xlsxrb::Ooxml::XmlBuilder.unescape("&unknown;"))

    # WorksheetParser.decode_xml_entities delegation
    assert_equal("Tom & Jerry", Xlsxrb::Ooxml::WorksheetParser.decode_xml_entities("Tom &amp; Jerry"))
  end

  test "xml_builder method chaining with declaration and tags" do
    io = StringIO.new
    b = Xlsxrb::Ooxml::XmlBuilder.new(io)
    ret = b.declaration.open_tag("item", { id: 1 }).text("chained").close_tag("item")
    assert_same(b, ret)
    assert_equal(%(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<item id="1">chained</item>), io.string)
  end

  test "xml_builder tag without block emits empty_tag and handles raw" do
    io = StringIO.new
    b = Xlsxrb::Ooxml::XmlBuilder.new(io)
    b.tag("col", { min: 1, max: 2 })
    b.raw("<raw>unescaped & content</raw>")

    assert_equal('<col min="1" max="2"/><raw>unescaped & content</raw>', io.string)
  end

  test "xml_builder write_unmapped restores unknown elements" do
    io = StringIO.new
    b = Xlsxrb::Ooxml::XmlBuilder.new(io)
    node = { tag: "custom", attrs: { "val" => "42" }, children: [], text: "data" }
    b.write_unmapped(node)

    assert_equal('<custom val="42">data</custom>', io.string)

    # Empty text and empty children produces empty tag
    io_empty = StringIO.new
    b_empty = Xlsxrb::Ooxml::XmlBuilder.new(io_empty)
    b_empty.write_unmapped({ tag: "empty_node", attrs: { "a" => "b" }, children: [], text: "" })
    assert_equal('<empty_node a="b"/>', io_empty.string)

    # Non-hash or missing tag does nothing
    io_noop = StringIO.new
    b_noop = Xlsxrb::Ooxml::XmlBuilder.new(io_noop)
    b_noop.write_unmapped(nil)
    b_noop.write_unmapped("invalid")
    b_noop.write_unmapped({})
    assert_equal("", io_noop.string)
  end

  # --- XmlParser::BaseListener ---

  test "base_listener captures unknown elements as nested unmapped_data trees" do
    listener_class = Class.new(Xlsxrb::Ooxml::XmlParser::BaseListener) do
      def recognized_tag?(_uri, localname, _qname)
        %w[root known].include?(localname)
      end
    end

    xml = <<~XML
      <root>
        <known id="1">Known Text</known>
        <extLst version="1.0">
          <ext uri="{1234}">
            <customFeature enabled="true">Feature Value</customFeature>
          </ext>
        </extLst>
      </root>
    XML

    listener = listener_class.new
    Xlsxrb::Ooxml::XmlParser.parse(xml, listener)

    unmapped = listener.unmapped_data
    assert_operator unmapped.size, :>=, 1
    ext_node = unmapped.find { |n| n[:tag] == "extLst" }
    assert_not_nil ext_node
    assert_equal({ "version" => "1.0" }, ext_node[:attrs])
    assert_equal(1, ext_node[:children].size)

    child = ext_node[:children].first
    assert_equal("ext", child[:tag])
    assert_equal({ "uri" => "{1234}" }, child[:attrs])
    assert_equal(1, child[:children].size)

    grandchild = child[:children].first
    assert_equal("customFeature", grandchild[:tag])
    assert_equal("Feature Value", grandchild[:text])
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

  test "shared_strings_parser decodes XML entities and handles self-closing tags" do
    xml = <<~XML
      <sst>
        <si><t>A &amp; B &lt; C &gt; &quot;D&quot; &apos;E&apos;</t></si>
        <si/>
        <si><t/></si>
        <si></si>
      </sst>
    XML

    strings = Xlsxrb::Ooxml::SharedStringsParser.parse(xml)
    assert_equal(["A & B < C > \"D\" 'E'", "", "", ""], strings)
  end

  test "shared_strings_parser handles self-closing sst and empty inputs" do
    assert_equal([], Xlsxrb::Ooxml::SharedStringsParser.parse("<sst/>"))
    assert_equal([], Xlsxrb::Ooxml::SharedStringsParser.parse("<sst></sst>"))
    assert_equal([], Xlsxrb::Ooxml::SharedStringsParser.parse(""))
    assert_equal([], Xlsxrb::Ooxml::SharedStringsParser.parse(nil))
  end

  test "shared_strings_parser handles rich text with empty t tags and entities" do
    xml = <<~XML
      <sst>
        <si>
          <r><t>Prefix: </t></r>
          <r><t/></r>
          <r><t>Middle &amp; &lt;special&gt; </t></r>
          <r><t>Suffix</t></r>
        </si>
      </sst>
    XML

    strings = Xlsxrb::Ooxml::SharedStringsParser.parse(xml)
    assert_equal(["Prefix: Middle & <special> Suffix"], strings)
  end

  test "shared_strings_parser each_event enumerates events with custom part_name" do
    xml = "<sst><si><t>Item 0</t></si><si><t>Item 1</t></si></sst>"
    events = Xlsxrb::Ooxml::SharedStringsParser.each_event(xml, part_name: "custom/sst.xml").to_a

    assert_equal(2, events.size)
    assert_equal(:sst_item, events[0].type)
    assert_equal(["Item 0"], events[0].args)
    assert_equal({ part: "custom/sst.xml", index: 0 }, events[0].source)
    assert_equal(["Item 1"], events[1].args)
    assert_equal({ part: "custom/sst.xml", index: 1 }, events[1].source)

    # Nil/empty each_event
    called = false
    assert_nil(Xlsxrb::Ooxml::SharedStringsParser.each_event(nil) { |_| called = true })
    assert_false(called)
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

  test "styles_parser parses fonts, fills, borders, alignments, and cellStyleXfs" do
    xml = <<~XML
      <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <fonts count="1">
          <font>
            <b/>
            <i/>
            <u val="double"/>
            <sz val="14.5"/>
            <color rgb="FFFF0000"/>
            <name val="Calibri"/>
          </font>
        </fonts>
        <fills count="1">
          <fill>
            <patternFill patternType="solid">
              <fgColor rgb="FF00FF00"/>
              <bgColor indexed="64"/>
            </patternFill>
          </fill>
        </fills>
        <borders count="1">
          <border>
            <left style="thin"/>
            <right style="thick"/>
            <top style="dashed"/>
            <bottom style="double"/>
            <diagonal style="medium"/>
          </border>
        </borders>
        <cellStyleXfs count="1">
          <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        </cellStyleXfs>
        <cellXfs count="1">
          <xf numFmtId="14" fontId="0" fillId="0" borderId="0" applyNumberFormat="1" applyFont="1">
            <alignment horizontal="center" vertical="top" wrapText="1" textRotation="90"/>
          </xf>
        </cellXfs>
      </styleSheet>
    XML

    result = Xlsxrb::Ooxml::StylesParser.parse(xml)
    assert_equal(1, result[:fonts].size)
    assert_equal(true, result[:fonts][0][:bold])
    assert_equal(true, result[:fonts][0][:italic])
    assert_equal("double", result[:fonts][0][:underline])
    assert_equal(14.5, result[:fonts][0][:sz])
    assert_equal({ rgb: "FFFF0000" }, result[:fonts][0][:color])
    assert_equal("Calibri", result[:fonts][0][:name])

    assert_equal(1, result[:fills].size)
    assert_equal("solid", result[:fills][0][:pattern])
    assert_equal({ rgb: "FF00FF00" }, result[:fills][0][:fg_color])
    assert_equal({ indexed: 64 }, result[:fills][0][:bg_color])

    assert_equal(1, result[:borders].size)
    assert_equal({ style: "thin" }, result[:borders][0][:left])
    assert_equal({ style: "thick" }, result[:borders][0][:right])
    assert_equal({ style: "dashed" }, result[:borders][0][:top])
    assert_equal({ style: "double" }, result[:borders][0][:bottom])
    assert_equal({ style: "medium" }, result[:borders][0][:diagonal])

    assert_equal(1, result[:cell_style_xfs].size)
    assert_equal(1, result[:cell_xfs].size)
    xf = result[:cell_xfs][0]
    assert_equal(14, xf[:num_fmt_id])
    assert_equal(true, xf[:apply_number_format])
    assert_equal(true, xf[:apply_font])
    assert_equal({ horizontal: "center", vertical: "top", wrap_text: true, text_rotation: 90 }, xf[:alignment])
  end

  test "styles_parser returns empty hash for nil input" do
    assert_equal({}, Xlsxrb::Ooxml::StylesParser.parse(nil))
    assert_equal({}, Xlsxrb::Ooxml::StylesParser.parse(""))
  end

  # --- WorkbookParser ---

  test "workbook_parser parses sheet list" do
    xml = <<~XML
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

  test "workbook_parser handles unnormalized names and fallback id attributes" do
    xml = <<~XML
      <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:other="http://example.com/rel">
        <sheets>
          <sheet name="Tom &amp; Jerry" sheetId="1" id="fallbackId1"/>
          <sheet name="Plain" sheetId="2" other:id="nsId2"/>
        </sheets>
      </workbook>
    XML

    sheets = Xlsxrb::Ooxml::WorkbookParser.parse(xml)
    assert_equal("Tom & Jerry", sheets[0][:name])
    assert_equal("fallbackId1", sheets[0][:r_id])
    assert_equal("nsId2", sheets[1][:r_id])

    assert_equal([], Xlsxrb::Ooxml::WorkbookParser.parse(""))
    assert_equal([], Xlsxrb::Ooxml::WorkbookParser.parse(nil))
  end

  # --- RelationshipsParser ---

  test "relationships_parser parses relationships" do
    xml = <<~XML
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
      </Relationships>
    XML

    rels = Xlsxrb::Ooxml::RelationshipsParser.parse(xml)
    assert_equal("worksheets/sheet1.xml", rels["rId1"])
    assert_equal("styles.xml", rels["rId2"])
  end

  test "relationships_parser handles missing attributes and empty inputs" do
    xml = <<~XML
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="validId" Target="sheet1.xml"/>
        <Relationship Id="missingTarget"/>
        <Relationship Target="missingId.xml"/>
      </Relationships>
    XML

    rels = Xlsxrb::Ooxml::RelationshipsParser.parse(xml)
    assert_equal({ "validId" => "sheet1.xml" }, rels)
    assert_equal({}, Xlsxrb::Ooxml::RelationshipsParser.parse(""))
    assert_equal({}, Xlsxrb::Ooxml::RelationshipsParser.parse(nil))
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

  test "worksheet_parser streaming each_row works with IO stream and chunks" do
    xml = <<~XML
      <?xml version="1.0" encoding="UTF-8"?>
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
          <row r="1"><c r="A1"><v>10</v></c></row>
          <row r="2"><c r="A2"><v>20</v></c></row>
          <row r="3"><c r="A3"><v>30</v></c></row>
        </sheetData>
      </worksheet>
    XML

    io = StringIO.new(xml)
    collected = []
    Xlsxrb::Ooxml::WorksheetParser.each_row(io, shared_strings: []) do |row|
      collected << [row[:index], row[0].value]
    end

    assert_equal([[0, 10], [1, 20], [2, 30]], collected)
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
