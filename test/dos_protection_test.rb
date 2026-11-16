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

      error = assert_raises(RuntimeError) do
        Xlsxrb.load(zip_path)
      end

      # REXML limits entity expansions and raises RuntimeError.
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

  test "a percent sign in a date format code is literal text, not a strftime directive" do
    code = "yyyy#{"%999999d" * 100}"
    formatted = within(5, "format") { Xlsxrb::NumberFormatter.format(45_000.5, code) }
    assert_operator formatted.bytesize, :<, 2000

    assert_equal "2023%15", Xlsxrb::NumberFormatter.format(45_000.5, "yyyy%d")
    assert_equal "2023-03-15 %", Xlsxrb::NumberFormatter.format(45_000, "yyyy-mm-dd %")
    assert_equal "50%", Xlsxrb::NumberFormatter.format(0.5, "0%")
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
