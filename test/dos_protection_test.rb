# frozen_string_literal: true

require "test_helper"
require "tempfile"
require "timeout"

class DosProtectionTest < Test::Unit::TestCase
  test "protects against XML Billion Laughs attack (entity expansion limit)" do
    billion_laughs_xml = <<~XML
      <?xml version="1.0" encoding="UTF-8"?>
      <!DOCTYPE lolz [
       <!ENTITY lol "lol">
       <!ELEMENT lolz (#PCDATA)>
       <!ENTITY lol1 "&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;">
       <!ENTITY lol2 "&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;">
       <!ENTITY lol3 "&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;">
       <!ENTITY lol4 "&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;">
       <!ENTITY lol5 "&lol4;&lol4;&lol4;&lol4;&lol4;&lol4;&lol4;&lol4;&lol4;&lol4;">
       <!ENTITY lol6 "&lol5;&lol5;&lol5;&lol5;&lol5;&lol5;&lol5;&lol5;&lol5;&lol5;">
       <!ENTITY lol7 "&lol6;&lol6;&lol6;&lol6;&lol6;&lol6;&lol6;&lol6;&lol6;&lol6;">
       <!ENTITY lol8 "&lol7;&lol7;&lol7;&lol7;&lol7;&lol7;&lol7;&lol7;&lol7;&lol7;">
       <!ENTITY lol9 "&lol8;&lol8;&lol8;&lol8;&lol8;&lol8;&lol8;&lol8;&lol8;&lol8;">
      ]>
      <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <si><t>&lol9;</t></si>
      </sst>
    XML

    zip_tempfile = Tempfile.new(["xlsxrb-dos", ".xlsx"])
    zip_path = zip_tempfile.path
    zip_tempfile.close

    begin
      generator = Xlsxrb::Ooxml::ZipGenerator.new(zip_path)
      generator.add_entry("xl/sharedStrings.xml", billion_laughs_xml)
      generator.add_entry("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>')
      generator.add_entry("xl/workbook.xml", '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>')
      generator.add_entry("xl/_rels/workbook.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
      generator.add_entry("xl/worksheets/sheet1.xml", '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>')
      generator.generate

      error = assert_raises(Xlsxrb::ParseError) do
        Xlsxrb.load(zip_path)
      end

      # REXML limits entity expansions; its RuntimeError is reported as ParseError.
      assert_match(/entity expansions exceeded/i, error.message)
    ensure
      FileUtils.rm_f(zip_path)
    end
  end

  # Note on Zip Bomb protection:
  # The `rubyzip` gem does not strictly prevent decompression of highly compressed archives (Zip Bombs)
  # on its own unless explicitly checked (e.g. tracking bytes read vs compressed size).
  # However, Xlsxrb provides a streaming read API (`Xlsxrb.read`) that avoids loading
  # the entire payload into memory at once. Users parsing untrusted files with `Xlsxrb.read`
  # should be aware of memory limitations.

  SML_NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
  REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

  test "the word hyperlink outside a tag does not stop the hyperlink scan from advancing" do
    [
      "<hyperlinks>hyperlink</hyperlinks>",
      %(<hyperlinks ref="A1">hyperlink</hyperlinks>),
      "<!-- hyperlinks> hyperlink -->"
    ].each do |markup|
      within(5, markup) { Xlsxrb.load(workbook_bytes(sheet: sheet_xml(%(<row r="1"><c r="A1"><v>1</v></c></row>), after: markup))) }
    end

    links = %(<hyperlinks><hyperlink ref="A1" location="S!A1" display="hyperlink"/>hyperlink<hyperlink ref="B2" location="S!B2"/></hyperlinks>)
    assert_equal(
      [{ ref: "A1", display: "hyperlink", location: "S!A1" }, { ref: "B2", location: "S!B2" }],
      Xlsxrb::Ooxml::WorksheetParser.parse_hyperlinks(sheet_xml("", after: links))
    )
  end

  test "deeply nested elements raise ParseError instead of taking time proportional to the depth squared" do
    depth = 40_000
    workbook = %(<workbook #{SML_NS}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>#{"<a>" * depth}#{"</a>" * depth}</workbook>)
    bytes = workbook_bytes("xl/workbook.xml" => workbook)

    within(5, "load") { assert_raise(Xlsxrb::ParseError) { Xlsxrb.load(bytes) } }
    error = within(5, "read") { assert_raise(Xlsxrb::ParseError) { Xlsxrb.read(bytes) { |sheet| sheet.each_row(&:values) } } }
    assert_match(/nested more than 256 levels/, error.message)

    nested = %(<workbook #{SML_NS}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>#{"<a>" * 200}#{"</a>" * 200}</workbook>)
    assert_equal ["S"], Xlsxrb.load(workbook_bytes("xl/workbook.xml" => nested)).sheets.map(&:name)
  end

  test "malformed XML in a part raises ParseError" do
    error = assert_raise(Xlsxrb::ParseError) { Xlsxrb.load(workbook_bytes("xl/workbook.xml" => "<workbook><sheets>")) }
    assert_match(/\AMalformed XML: /, error.message)
    assert_not_include error.message, "\n"
  end

  test "rows with unclosed cell or inline string tags raise ParseError instead of rescanning the row for each tag" do
    count = 40_000
    {
      "c" => %(<c r="A1">) * count,
      "c with f" => %(<c r="A1"><f>) * count,
      "c after a closed cell" => %(<c r="A1"><v>1</v></c>#{%(<c r="B1">) * count}),
      "is" => %(<c r="A1" t="inlineStr"><is></c>) * count
    }.each do |label, cells|
      bytes = workbook_bytes(sheet: sheet_xml(%(<row r="1">#{cells}</row>)))

      within(5, "#{label} read") { assert_raise(Xlsxrb::ParseError, label) { Xlsxrb.read(bytes) { |sheet| sheet.each_row(&:values) } } }
      within(5, "#{label} load") { assert_raise(Xlsxrb::ParseError, label) { Xlsxrb.load(bytes).sheets.first.rows.each(&:cells) } }
    end

    cells = %(<c r="A1" t="inlineStr"><is><t>hi</t></is></c><c r="B1" s="0"/><x:c r="C1"><x:v>2</x:v></x:c> <c r="D1"/> )
    bytes = workbook_bytes(sheet: sheet_xml(%(<row r="1">#{cells}</row><row r="2"><c r="A2"/><c r="B2"/></row>)))
    assert_equal [["hi", nil, 2, nil], [nil, nil]], Xlsxrb.read(bytes).first.each_row.map(&:values)
    assert_equal([["hi", nil, 2, nil], [nil, nil]], Xlsxrb.load(bytes).sheets.first.rows.map { |row| row.cells.map(&:value) })
  end

  test "a percent sign in a date format code is literal text, not a strftime directive" do
    code = "yyyy#{"%999999d" * 100}"
    formatted = within(5, "format") { Xlsxrb::NumberFormatter.format(45_000.5, code) }
    assert_operator formatted.bytesize, :<, 2000

    assert_equal "2023%15", Xlsxrb::NumberFormatter.format(45_000.5, "yyyy%d")
    assert_equal "2023-03-15 %", Xlsxrb::NumberFormatter.format(45_000, "yyyy-mm-dd %")
    assert_equal "50%", Xlsxrb::NumberFormatter.format(0.5, "0%")
  end

  test "format codes and values outside what Excel accepts are formatted as General" do
    assert_equal "1.5", Xlsxrb::NumberFormatter.format(1.5, "0.#{"0" * 2000}")
    assert_equal "1.50", Xlsxrb::NumberFormatter.format(1.5, "0.00#{" " * 1000}")
    assert_equal "1.#{"0" * 1022}", Xlsxrb::NumberFormatter.format(1, "0.#{"0" * 1022}")

    assert_equal "Infinity", Xlsxrb::NumberFormatter.format(Float::INFINITY, "yyyy-mm-dd")
    assert_equal "-Infinity", Xlsxrb::NumberFormatter.format(-Float::INFINITY, "0.00")
    assert_equal "NaN", Xlsxrb::NumberFormatter.format(Float::NAN, "0")
    assert_equal "1000000000000000", Xlsxrb::NumberFormatter.format(1e15, "yyyy-mm-dd")
    assert_equal "-5", Xlsxrb::NumberFormatter.format(-5, "yyyy-mm-dd")
    assert_equal "2023-03-15", Xlsxrb::NumberFormatter.format(45_000, "yyyy-mm-dd")
  end

  private

  def within(seconds, label, &)
    Timeout.timeout(seconds, Timeout::Error, "#{label} took more than #{seconds} s", &)
  end

  def sheet_xml(rows, before: "", after: "")
    %(<worksheet #{SML_NS}>#{before}<sheetData>#{rows}</sheetData>#{after}</worksheet>)
  end

  # A minimal workbook with one sheet named S and one shared string, "a".
  def workbook_bytes(parts = {})
    parts = parts.dup
    sheet = parts.delete(:sheet) || sheet_xml(%(<row r="1"><c r="A1"><v>1</v></c></row>))
    rels = %(<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">%s</Relationships>)
    defaults = {
      "[Content_Types].xml" => '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
      "_rels/.rels" => format(rels, %(<Relationship Id="rId1" Type="#{REL_TYPE}/officeDocument" Target="xl/workbook.xml"/>)),
      "xl/workbook.xml" => %(<workbook #{SML_NS}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>),
      "xl/_rels/workbook.xml.rels" => format(
        rels,
        %(<Relationship Id="rId1" Type="#{REL_TYPE}/worksheet" Target="worksheets/sheet1.xml"/>) +
        %(<Relationship Id="rId2" Type="#{REL_TYPE}/sharedStrings" Target="sharedStrings.xml"/>)
      ),
      "xl/sharedStrings.xml" => %(<sst #{SML_NS}><si><t>a</t></si></sst>),
      "xl/worksheets/sheet1.xml" => sheet
    }
    io = StringIO.new
    Xlsxrb::Ooxml::ZipWriter.open(io) do |zip|
      defaults.merge(parts).each { |name, content| zip.add_entry(name, content) }
    end
    io.string
  end
end
