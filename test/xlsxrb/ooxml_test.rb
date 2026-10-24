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

  test "zip_reader normalizes Windows backslash path separators" do
    io = StringIO.new
    Xlsxrb::Ooxml::ZipWriter.open(io) do |w|
      w.add_entry("xl\\workbook.xml", "<workbook/>")
      w.add_entry("xl\\worksheets\\sheet1.xml", "<worksheet/>")
    end
    io.rewind

    reader = Xlsxrb::Ooxml::ZipReader.new(io)
    assert_equal(true, reader.entry?("xl/workbook.xml"))
    assert_equal(true, reader.entry?("xl\\workbook.xml"))
    assert_equal(true, reader.entry?("xl/worksheets/sheet1.xml"))
    assert_equal("<workbook/>", reader.read_entry("xl/workbook.xml"))
    assert_equal("<workbook/>", reader.read_entry("xl\\workbook.xml"))
    assert_equal("<worksheet/>", reader.read_entry("xl/worksheets/sheet1.xml"))

    stream = reader.open_entry_io("xl/workbook.xml")
    assert_equal("<workbook/>", stream.read)
    stream.close

    stream_bs = reader.open_entry_io("xl\\workbook.xml")
    assert_equal("<workbook/>", stream_bs.read)
    stream_bs.close
  end

  test "zip_reader matches part names case-insensitively per ECMA-376 Part 2" do
    io = StringIO.new
    Xlsxrb::Ooxml::ZipWriter.open(io) do |w|
      w.add_entry("xl/sharedstrings.xml", "<sst/>")
      w.add_entry("xl/STYLES.XML", "<styleSheet/>")
    end
    io.rewind

    reader = Xlsxrb::Ooxml::ZipReader.new(io)
    assert_equal(true, reader.entry?("xl/sharedStrings.xml"))
    assert_equal(true, reader.entry?("xl/sharedstrings.xml"))
    assert_equal(true, reader.entry?("xl/styles.xml"))
    assert_equal(true, reader.entry?("xl/STYLES.XML"))
    assert_equal("<sst/>", reader.read_entry("xl/sharedStrings.xml"))
    assert_equal("<sst/>", reader.read_entry("xl/sharedstrings.xml"))
    assert_equal("<styleSheet/>", reader.read_entry("xl/styles.xml"))
    assert_equal("<styleSheet/>", reader.read_entry("xl/STYLES.XML"))

    stream = reader.open_entry_io("xl/sharedStrings.xml")
    assert_equal("<sst/>", stream.read)
    stream.close
  end

  test "zip_reader normalizes UTF-16LE, UTF-16BE, and UTF-8 BOM encodings to UTF-8" do
    utf16le_content = "\uFEFF<root>UTF16LE Content</root>".encode(Encoding::UTF_16LE)
    utf16be_content = "\uFEFF<root>UTF16BE Content</root>".encode(Encoding::UTF_16BE)
    utf8_bom_content = "\xEF\xBB\xBF<root>UTF8 BOM Content</root>".b

    io = StringIO.new
    Xlsxrb::Ooxml::ZipWriter.open(io) do |w|
      w.add_binary_entry("utf16le.xml", utf16le_content.b)
      w.add_binary_entry("utf16be.xml", utf16be_content.b)
      w.add_binary_entry("utf8bom.xml", utf8_bom_content)
    end
    io.rewind

    reader = Xlsxrb::Ooxml::ZipReader.new(io)
    assert_equal("<root>UTF16LE Content</root>", reader.read_entry("utf16le.xml"))
    assert_equal("<root>UTF16BE Content</root>", reader.read_entry("utf16be.xml"))
    assert_equal("<root>UTF8 BOM Content</root>", reader.read_entry("utf8bom.xml"))

    stream_le = reader.open_entry_io("utf16le.xml")
    assert_equal("<root>UTF16LE Content</root>", stream_le.read)
    stream_le.close

    stream_be = reader.open_entry_io("utf16be.xml")
    assert_equal("<root>UTF16BE Content</root>", stream_be.read)
    stream_be.close

    stream_bom = reader.open_entry_io("utf8bom.xml")
    assert_equal("<root>UTF8 BOM Content</root>", stream_bom.read)
    stream_bom.close
  end

  test "zip_reader streaming yields UTF-8 chunks for XML and rels entries and respects encoding option" do
    xml_data = "<worksheet><sheetData><row r=\"1\"><c r=\"A1\"><is><t>日本語テスト</t></is></c></row></sheetData></worksheet>"
    txt_data = "plain binary text"

    io = StringIO.new
    Xlsxrb::Ooxml::ZipWriter.open(io) do |w|
      w.add_entry("xl/worksheets/sheet1.xml", xml_data)
      w.add_entry("_rels/.rels", "<Relationships/>")
      w.add_entry("data.txt", txt_data)
    end
    io.rewind

    reader = Xlsxrb::Ooxml::ZipReader.new(io)

    # XML entries default to UTF-8
    stream = reader.open_entry_io("xl/worksheets/sheet1.xml")
    chunk = stream.read
    assert_equal(Encoding::UTF_8, chunk.encoding)
    assert_match(/日本語/, chunk)
    stream.close

    # each_entry_chunk defaults to UTF-8 for XML
    chunks = []
    reader.each_entry_chunk("xl/worksheets/sheet1.xml", chunk_size: 32) do |c|
      assert_equal(Encoding::UTF_8, c.encoding)
      chunks << c
    end
    assert_equal(xml_data, chunks.join)

    # .rels defaults to UTF-8
    rels_stream = reader.open_entry_io("_rels/.rels")
    assert_equal(Encoding::UTF_8, rels_stream.read.encoding)
    rels_stream.close

    # Non-XML defaults to binary (ASCII-8BIT)
    txt_stream = reader.open_entry_io("data.txt")
    assert_equal(Encoding::ASCII_8BIT, txt_stream.read.encoding)
    txt_stream.close

    # Explicit encoding option can override
    raw_xml_stream = reader.open_entry_io("xl/worksheets/sheet1.xml", encoding: false)
    assert_equal(Encoding::ASCII_8BIT, raw_xml_stream.read.encoding)
    raw_xml_stream.close

    custom_txt_stream = reader.open_entry_io("data.txt", encoding: "UTF-8")
    assert_equal(Encoding::UTF_8, custom_txt_stream.read.encoding)
    custom_txt_stream.close
  end

  test "EntryStreamIO#read with byte length slices exact bytes on UTF-8 stream" do
    data = "日本語" * 20
    io = StringIO.new
    Xlsxrb::Ooxml::ZipWriter.open(io) do |w|
      w.add_entry("test.xml", data)
    end
    io.rewind

    reader = Xlsxrb::Ooxml::ZipReader.new(io)
    stream = reader.open_entry_io("test.xml", encoding: Encoding::UTF_8)
    part = stream.read(18)
    assert_equal(18, part.bytesize)
    assert_equal(6, part.chars.size)
    assert_equal("日本語日本語", part)
    stream.close
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

  test "zip_writer copy_raw_entry and zip_reader copy_to_writer preserves payload and metadata" do
    io1 = StringIO.new
    binary_data = "\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR".b * 20
    Xlsxrb::Ooxml::ZipWriter.open(io1) do |w|
      w.add_entry("text.txt", "Original text content")
      w.add_binary_entry("image.png", binary_data)
    end

    reader1 = Xlsxrb::Ooxml::ZipReader.open(StringIO.new(io1.string))
    io2 = StringIO.new
    Xlsxrb::Ooxml::ZipWriter.open(io2) do |w|
      reader1.copy_to_writer("text.txt", w)
      reader1.copy_to_writer("image.png", w)
    end

    reader2 = Xlsxrb::Ooxml::ZipReader.open(StringIO.new(io2.string))
    assert_equal("Original text content", reader2.read_entry("text.txt"))
    assert_equal(binary_data, reader2.read_entry("image.png").b)

    # Validate catalog metadata parity
    e1 = reader1.entry_catalog["image.png"]
    e2 = reader2.entry_catalog["image.png"]
    assert_equal(e1.crc32, e2.crc32)
    assert_equal(e1.compressed_size, e2.compressed_size)
    assert_equal(e1.uncompressed_size, e2.uncompressed_size)
    assert_equal(e1.method, e2.method)
  end

  test "zip_reader parses ZIP64 extra fields in central directory and local headers" do
    filename = "test.txt"
    payload = "Hello ZIP64!"
    crc = Zlib.crc32(payload)
    u_size = 5_000_000_000
    c_size = 4_000_000_000
    local_offset = 0

    # Local extra
    local_extra = [1, 16, u_size, c_size].pack("vvQ<Q<")
    local_hdr = "PK\x03\x04#{[45, 0, 0, 0, 0, crc, 0xFFFF_FFFF, 0xFFFF_FFFF, filename.bytesize, local_extra.bytesize].pack("vvvvvVVVvv")}#{filename}#{local_extra}"

    # Central Directory extra
    cd_extra = [1, 24, u_size, c_size, local_offset].pack("vvQ<Q<Q<")
    cd_hdr = "PK\x01\x02#{[45, 45, 0, 0, 0, 0, crc, 0xFFFF_FFFF, 0xFFFF_FFFF, filename.bytesize, cd_extra.bytesize, 0, 0, 0, 0, 0xFFFF_FFFF].pack("vvvvvvVVVvvvvvVV")}#{filename}#{cd_extra}"

    cd_offset = local_hdr.bytesize + payload.bytesize
    cd_size = cd_hdr.bytesize
    eocd = "PK\x05\x06#{[0, 0, 1, 1, cd_size, cd_offset, 0].pack("vvvvVVv")}"

    buf = local_hdr + payload + cd_hdr + eocd
    reader = Xlsxrb::Ooxml::ZipReader.new(StringIO.new(buf))
    entry = reader.entry_catalog["test.txt"]

    assert_equal u_size, entry.uncompressed_size
    assert_equal c_size, entry.compressed_size
    assert_equal local_offset, entry.local_header_offset

    # Local header fallback test
    local_extra_fallback = [1, 16, payload.bytesize, payload.bytesize].pack("vvQ<Q<")
    local_hdr_fallback = "PK\x03\x04#{[45, 0, 0, 0, 0, crc, 0xFFFF_FFFF, 0xFFFF_FFFF, filename.bytesize, local_extra_fallback.bytesize].pack("vvvvvVVVvv")}#{filename}#{local_extra_fallback}"
    reader_fallback = Xlsxrb::Ooxml::ZipReader.new(StringIO.new(local_hdr_fallback + payload))
    entry_fallback = reader_fallback.entry_catalog["test.txt"]

    assert_equal payload.bytesize, entry_fallback.uncompressed_size
    assert_equal payload.bytesize, entry_fallback.compressed_size
    assert_equal payload, reader_fallback.read_entry("test.txt")
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
    assert_equal("a &# b", Xlsxrb::Ooxml::XmlBuilder.unescape("a &# b"))

    # Decimal entities
    assert_equal("ABC", Xlsxrb::Ooxml::XmlBuilder.unescape("&#65;&#66;&#67;"))
    assert_equal("A", Xlsxrb::Ooxml::XmlBuilder.unescape("&#65;"))

    # Hexadecimal entities (lowercase, uppercase, multibyte)
    assert_equal("ABC", Xlsxrb::Ooxml::XmlBuilder.unescape("&#x41;&#x42;&#x43;"))
    assert_equal("a", Xlsxrb::Ooxml::XmlBuilder.unescape("&#x61;"))
    assert_equal("\x10", Xlsxrb::Ooxml::XmlBuilder.unescape("&#x10;"))
    emoji = Xlsxrb::Ooxml::XmlBuilder.unescape("&#x1F600;")
    assert_equal("😀", emoji)
    assert_equal(Encoding::UTF_8, emoji.encoding)

    # Out of range codepoints (RangeError rescued to match string)
    assert_equal("&#9999999999;", Xlsxrb::Ooxml::XmlBuilder.unescape("&#9999999999;"))
    assert_equal("&#xFFFFFFFFF;", Xlsxrb::Ooxml::XmlBuilder.unescape("&#xFFFFFFFFF;"))

    # Mixed entities and text
    assert_equal("Hello <World> & 😀!", Xlsxrb::Ooxml::XmlBuilder.unescape("Hello &lt;World&gt; &amp; &#x1F600;!"))

    # ISO/IEC 29500-1 ST_Xstring (_xHHHH_) escape sequences
    assert_equal("\r", Xlsxrb::Ooxml::XmlBuilder.unescape("_x000D_"))
    assert_equal("\n", Xlsxrb::Ooxml::XmlBuilder.unescape("_x000A_"))
    assert_equal("\t", Xlsxrb::Ooxml::XmlBuilder.unescape("_x0009_"))
    assert_equal(" ", Xlsxrb::Ooxml::XmlBuilder.unescape("_x0020_"))
    assert_equal("ABC", Xlsxrb::Ooxml::XmlBuilder.unescape("_x0041__x0042__x0043_"))
    assert_equal("abc", Xlsxrb::Ooxml::XmlBuilder.unescape("_x0061__x0062__x0063_"))
    assert_equal("_x000D_", Xlsxrb::Ooxml::XmlBuilder.unescape("_x005F_x000D_"))
    assert_equal("Line 1\rLine 2", Xlsxrb::Ooxml::XmlBuilder.unescape("Line 1_x000D_Line 2"))
    assert_equal("AT&T\rLine", Xlsxrb::Ooxml::XmlBuilder.unescape("AT&amp;T_x000D_Line"))
    assert_equal("_xZZZZ_", Xlsxrb::Ooxml::XmlBuilder.unescape("_xZZZZ_"))
    assert_equal("_x12_", Xlsxrb::Ooxml::XmlBuilder.unescape("_x12_"))
    assert_equal("_x12345_", Xlsxrb::Ooxml::XmlBuilder.unescape("_x12345_"))

    # WorksheetParser.decode_xml_entities delegation
    assert_equal("Tom & Jerry", Xlsxrb::Ooxml::WorksheetParser.decode_xml_entities("Tom &amp; Jerry"))
    assert_equal("Tom\rJerry", Xlsxrb::Ooxml::WorksheetParser.decode_xml_entities("Tom_x000D_Jerry"))
    assert_equal("Tom\nJerry", Xlsxrb::Ooxml::WorksheetParser.normalize_text("Tom_x000D_Jerry"))
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

  test "shared_strings_parser decodes ST_Xstring _xHHHH_ sequences and normalizes newlines" do
    xml = <<~XML
      <sst>
        <si><t>Line1_x000D_\nLine2</t></si>
        <si><t>Line1_x000D_Line2</t></si>
        <si><t>Literal_x005F_x000D_Escaped</t></si>
        <si><r><t>Run1_x000D_</t></r><r><t>Run2</t></r></si>
      </sst>
    XML

    strings = Xlsxrb::Ooxml::SharedStringsParser.parse(xml)
    assert_equal(%W[Line1\nLine2 Line1\nLine2 Literal_x000D_Escaped Run1\nRun2], strings)
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

  test "shared_strings_parser handles XML namespace prefixes" do
    xml = <<~XML
      <?xml version="1.0" encoding="utf-8"?>
      <x:sst count="3" uniqueCount="3" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <x:si><x:t>First</x:t></x:si>
        <x:si><x:r><x:t>Rich </x:t></x:r><x:r><x:t>Text</x:t></x:r></x:si>
        <x:si><x:t>Third</x:t><x:rPh sb="0" eb="2"><x:t>Furigana</x:t></x:rPh></x:si>
      </x:sst>
    XML

    strings = Xlsxrb::Ooxml::SharedStringsParser.parse(xml)
    assert_equal(["First", "Rich Text", "Third"], strings)

    events = Xlsxrb::Ooxml::SharedStringsParser.each_event(xml).to_a
    assert_equal(3, events.size)
    assert_equal(["First"], events[0].args)
    assert_equal(["Rich Text"], events[1].args)
    assert_equal(["Third"], events[2].args)
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

  test "relationships_parser handles duplicate xml declarations and malformed xml" do
    duplicate_decl_xml = %(<?xml version="1.0" encoding="UTF-8"?><?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>)
    assert_equal({ "rId1" => "worksheets/sheet1.xml" }, Xlsxrb::Ooxml::RelationshipsParser.parse(duplicate_decl_xml))

    malformed_xml = %(<Relationships><Relationship Target="worksheets/sheet2.xml" Type="http://..." Id="rId2"><broken xml>)
    assert_equal({ "rId2" => "worksheets/sheet2.xml" }, Xlsxrb::Ooxml::RelationshipsParser.parse(malformed_xml))
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

  test "worksheet_parser handles self-closing row tags preserving attributes and events" do
    xml = <<~XML
      <?xml version="1.0" encoding="UTF-8"?>
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
          <row r="1"><c r="A1"><v>100</v></c></row>
          <row r="2" ht="28.5" hidden="1" customHeight="1"/>
          <row r="3"><c r="A3"><v>300</v></c></row>
        </sheetData>
      </worksheet>
    XML

    # 1. String source (fast_scan_rows_direct)
    rows = Xlsxrb::Ooxml::WorksheetParser.each_row(xml).to_a
    assert_equal(3, rows.size)

    assert_equal(0, rows[0].index)
    assert_equal([100], rows[0].values)

    assert_equal(1, rows[1].index)
    assert_equal(28.5, rows[1].height)
    assert_equal(true, rows[1].hidden?)
    assert_equal(true, rows[1].custom_height?)
    assert_empty(rows[1].cells)
    assert_empty(rows[1].values)

    assert_equal(2, rows[2].index)
    assert_equal([300], rows[2].values)

    # 2. IO source (scan_rows_stream)
    io_rows = Xlsxrb::Ooxml::WorksheetParser.each_row(StringIO.new(xml)).to_a
    assert_equal(3, io_rows.size)
    assert_equal(1, io_rows[1].index)
    assert_equal(28.5, io_rows[1].height)
    assert_equal(true, io_rows[1].hidden?)
    assert_empty(io_rows[1].values)

    # 3. Events (each_event / fast_scan_events)
    events = Xlsxrb::Ooxml::WorksheetParser.each_event(xml).to_a
    row_starts = events.select { |e| e.type == :row_start }
    assert_equal(3, row_starts.size)
    assert_equal(1, row_starts[1].args[0])
    assert_equal(28.5, row_starts[1].args[1][:height])
    assert_equal(true, row_starts[1].args[1][:hidden])
    row_ends = events.select { |e| e.type == :row_end }
    assert_equal(3, row_ends.size)
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

  test "worksheet_parser parses shared formula cells with self-closing f tags" do
    xml = <<~XML
      <?xml version="1.0" encoding="UTF-8"?>
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
          <row r="1">
            <c r="A1" s="2"><f t="shared" ref="A1:C1" si="0">B1+C1</f><v>10</v></c>
            <c r="B1" s="2"><f t="shared" si="0"/><v>20</v></c>
            <c r="C1" s="2"><f t="shared" si="0" /><v>30</v></c>
          </row>
          <row r="2">
            <c s="1" r="A2"><f t="shared" si="0"/><v>40</v></c>
          </row>
        </sheetData>
      </worksheet>
    XML

    rows = Xlsxrb::Ooxml::WorksheetParser.parse(xml, shared_strings: [])
    assert_equal(2, rows.size)

    row1_cells = rows[0][:cells]
    assert_equal(3, row1_cells.size)
    assert_equal("A1", row1_cells[0][:ref])
    assert_equal(10, row1_cells[0][:value])
    assert_equal("B1+C1", row1_cells[0][:formula])

    assert_equal("B1", row1_cells[1][:ref])
    assert_equal(20, row1_cells[1][:value])
    assert_nil(row1_cells[1][:formula])

    assert_equal("C1", row1_cells[2][:ref])
    assert_equal(30, row1_cells[2][:value])
    assert_nil(row1_cells[2][:formula])

    row2_cells = rows[1][:cells]
    assert_equal(1, row2_cells.size)
    assert_equal("A2", row2_cells[0][:ref])
    assert_equal(40, row2_cells[0][:value])
    assert_nil(row2_cells[0][:formula])
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

  test "worksheet_writer emits fallback <v>0</v> for formula cells without cached value" do
    io = StringIO.new
    writer = Xlsxrb::Ooxml::WorksheetWriter.new(io)
    writer.start
    formula_obj = Xlsxrb::Elements::Formula.new(expression: "SUM(A1:A10)")
    formula_with_val = Xlsxrb::Elements::Formula.new(expression: "SUM(B1:B10)", cached_value: 42)
    writer.write_row(0, [
                       Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, formula: formula_obj),
                       Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 1, formula: formula_with_val)
                     ])
    writer.write_row_values(1, [{ formula: "A1*2" }, formula_obj, { formula: "NOW()" }])
    writer.finish

    xml = io.string
    assert_include(xml, '<c r="A1"><f>SUM(A1:A10)</f><v>0</v></c>')
    assert_include(xml, '<c r="B1"><f>SUM(B1:B10)</f><v>42</v></c>')
    assert_include(xml, '<c r="A2"><f>A1*2</f><v>0</v></c>')
    assert_include(xml, '<c r="B2"><f>SUM(A1:A10)</f><v>0</v></c>')
    assert_include(xml, '<c r="C2"><f>NOW()</f><v>0</v></c>')
  end

  test "worksheet_writer write_row_values properly serializes styled formula cells" do
    io = StringIO.new
    writer = Xlsxrb::Ooxml::WorksheetWriter.new(io)
    writer.start
    writer.write_row_values(0, [{ formula: "SUM(A1:A5)" }], styles: { 0 => "currency" }, style_map: { "currency" => 5 })
    writer.finish

    xml = io.string
    assert_include(xml, '<c r="A1" s="5"><f>SUM(A1:A5)</f><v>0</v></c>')
  end

  test "worksheet_writer write_row_values writes formula-prefixed strings as text instead of formulas" do
    io = StringIO.new
    sst = []
    sst_index = {}
    writer = Xlsxrb::Ooxml::WorksheetWriter.new(io)
    writer.start
    writer.write_row_values(0, ["=SUM(A1:A5)"], sst: sst, sst_index: sst_index)
    writer.finish

    xml = io.string
    assert_equal ["=SUM(A1:A5)"], sst
    assert_include(xml, '<c r="A1" t="s"><v>0</v></c>')
    assert_not_include(xml, "<f>")
  end

  test "extract_inline_text parses formatted rich text without confusing formatting tags with t elements" do
    is_xml = '<is><r><rPr><b/><sz val="11"/><vertAlign val="superscript"/></rPr><t>Bold</t></r><r><t>Normal</t></r></is>'
    text = Xlsxrb::Ooxml::WorksheetParser.send(:extract_inline_text, is_xml, 0, is_xml.bytesize)
    assert_equal "BoldNormal", text
  end

  test "worksheet_writer write_row_values escapes string formula cached_value" do
    io = StringIO.new
    writer = Xlsxrb::Ooxml::WorksheetWriter.new(io)
    writer.start
    formula = Xlsxrb.formula("CONCAT(A1, B1)", cached_value: "A & B < C >")
    writer.write_row_values(0, [formula])
    writer.finish

    xml = io.string
    assert_include(xml, '<c r="A1"><f>CONCAT(A1, B1)</f><v>A &amp; B &lt; C &gt;</v></c>')
  end

  test "worksheet_writer write_row_values properly serializes styled cell error formula" do
    io = StringIO.new
    writer = Xlsxrb::Ooxml::WorksheetWriter.new(io)
    writer.start
    formula = Xlsxrb.formula("1/0", cached_value: Xlsxrb::Elements::CellError.new("#DIV/0!"))
    writer.write_row_values(0, [formula], styles: { 0 => "warn" }, style_map: { "warn" => 3 })
    writer.finish

    xml = io.string
    assert_include(xml, '<c r="A1" s="3" t="e"><f>1/0</f><v>#DIV/0!</v></c>')
  end

  test "extract_inline_text handles UTF-8 strings with multibyte characters without byte drift" do
    is_xml = "<is><r><t>日本語</t></r><r><t>テスト</t></r></is>"
    text = Xlsxrb::Ooxml::WorksheetParser.send(:extract_inline_text, is_xml, 0, is_xml.bytesize)
    assert_equal "日本語テスト", text
  end

  test "SharedStringsParser handles rich text when first t is self-closing" do
    xml = %(<sst count="1" uniqueCount="1"><si><r><t/></r><r><t>Hello</t></r></si></sst>)
    strings = Xlsxrb::Ooxml::SharedStringsParser.parse(xml)
    assert_equal ["Hello"], strings
  end

  test "WorksheetListener excludes phonetic readings and ignores empty self-closing string cells" do
    xml = <<~XML
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
          <row r="1">
            <c r="A1" t="inlineStr"><is><t>東京</t><rPh sb="0" eb="2"><t>トウキョウ</t></rPh></is></c>
            <c r="B1" s="1" t="s"/>
          </row>
        </sheetData>
      </worksheet>
    XML

    listener = Xlsxrb::Ooxml::Reader::WorksheetListener.new([])
    parser = REXML::Parsers::SAX2Parser.new(xml)
    parser.listen(listener)
    parser.parse

    assert_equal "東京", listener.cells["A1"]
    assert_nil listener.cells["B1"]
  end

  test "SharedStringsListener excludes phonetic readings from parsed strings" do
    xml = <<~XML
      <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <si><t>部門1</t><rPh sb="0" eb="2"><t>ブモン</t></rPh></si>
      </sst>
    XML

    listener = Xlsxrb::Ooxml::Reader::SharedStringsListener.new
    parser = REXML::Parsers::SAX2Parser.new(xml)
    parser.listen(listener)
    parser.parse

    assert_equal ["部門1"], listener.strings
  end

  test "worksheet_writer correctly formats formula cells with boolean, zero, string, and nil cached_value" do
    io = StringIO.new
    writer = Xlsxrb::Ooxml::WorksheetWriter.new(io)
    writer.start

    formula_false = Xlsxrb::Elements::Formula.new(expression: "A1>10", cached_value: false)
    formula_true = Xlsxrb::Elements::Formula.new(expression: "A1>5", cached_value: true)
    formula_zero = Xlsxrb::Elements::Formula.new(expression: "SUM(B1:B5)", cached_value: 0)
    formula_str = Xlsxrb::Elements::Formula.new(expression: "CONCAT(A1,B1)", cached_value: "res")
    formula_nil = Xlsxrb::Elements::Formula.new(expression: "A1+1")

    writer.write_row(0, [
                       Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, formula: formula_false),
                       Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 1, formula: formula_true),
                       Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 2, formula: formula_zero),
                       Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 3, formula: formula_str),
                       Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 4, formula: formula_nil)
                     ])

    writer.write_row_values(1, [
                              formula_false,
                              formula_true,
                              formula_zero,
                              { formula: "A1>10", value: false },
                              { formula: "A1>5", value: true },
                              { formula: "SUM(A1:A5)", value: 0 }
                            ])
    writer.finish

    xml = io.string
    # Row 1 (write_row)
    assert_include(xml, '<c r="A1" t="b"><f>A1&gt;10</f><v>0</v></c>')
    assert_include(xml, '<c r="B1" t="b"><f>A1&gt;5</f><v>1</v></c>')
    assert_include(xml, '<c r="C1"><f>SUM(B1:B5)</f><v>0</v></c>')
    assert_include(xml, '<c r="D1"><f>CONCAT(A1,B1)</f><v>res</v></c>')
    assert_include(xml, '<c r="E1"><f>A1+1</f><v>0</v></c>')

    # Row 2 (write_row_values)
    assert_include(xml, '<c r="A2" t="b"><f>A1&gt;10</f><v>0</v></c>')
    assert_include(xml, '<c r="B2" t="b"><f>A1&gt;5</f><v>1</v></c>')
    assert_include(xml, '<c r="C2"><f>SUM(B1:B5)</f><v>0</v></c>')
    assert_include(xml, '<c r="D2" t="b"><f>A1&gt;10</f><v>0</v></c>')
    assert_include(xml, '<c r="E2" t="b"><f>A1&gt;5</f><v>1</v></c>')
    assert_include(xml, '<c r="F2"><f>SUM(A1:A5)</f><v>0</v></c>')
  end

  test "build_raw_cell_from_value retains false and zero cached_values with proper type" do
    f_false = Xlsxrb::Elements::Formula.new(expression: "A1>10", cached_value: false)
    f_zero = Xlsxrb::Elements::Formula.new(expression: "SUM(A1:A5)", cached_value: 0)
    sst = []
    sst_index = {}

    cell_false = Xlsxrb.build_raw_cell_from_value(0, 0, f_false, sst, sst_index)
    assert_equal(false, cell_false[:value])
    assert_equal("b", cell_false[:type])
    assert_equal("A1>10", cell_false[:formula])

    cell_zero = Xlsxrb.build_raw_cell_from_value(0, 1, f_zero, sst, sst_index)
    assert_equal(0, cell_zero[:value])
    assert_nil(cell_zero[:type])
    assert_equal("SUM(A1:A5)", cell_zero[:formula])
  end

  test "Ooxml::Reader resolves case-insensitive part names via ZipReader" do
    Dir.mktmpdir do |dir|
      xlsx_path = File.join(dir, "case_insensitive.xlsx")
      Xlsxrb::Ooxml::ZipWriter.open(xlsx_path) do |zw|
        zw.add_entry("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>')
        zw.add_entry("_rels/.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        zw.add_entry("xl/WORKBOOK.XML", '<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Main" sheetId="1" r:id="rId1"/></sheets></workbook>')
        zw.add_entry("xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
        zw.add_entry("xl/sharedstrings.xml", '<?xml version="1.0" encoding="UTF-8"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>CaseInsensitiveValue</t></si></sst>')
        zw.add_entry("xl/worksheets/SHEET1.XML", '<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>')
      end

      reader = Xlsxrb::Ooxml::Reader.new(xlsx_path)
      assert_equal(["Main"], reader.sheet_names)
      assert_equal({ "A1" => "CaseInsensitiveValue" }, reader.cells)
      assert_include(reader.entry_names, "xl/sharedstrings.xml")
    end
  end
end
