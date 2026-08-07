# frozen_string_literal: true

require "test_helper"
require "tempfile"

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
        Xlsxrb.read(zip_path)
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
  # However, Xlsxrb provides a streaming read API (`Xlsxrb.foreach`) that avoids loading
  # the entire payload into memory at once. Users parsing untrusted files with `Xlsxrb.read`
  # should be aware of memory limitations.
end
