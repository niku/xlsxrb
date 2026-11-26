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
      within(5, markup) { Xlsxrb.load(StringIO.new(workbook_bytes(sheet: sheet_xml(%(<row r="1"><c r="A1"><v>1</v></c></row>), after: markup)))) }
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

    within(5, "load") { assert_raise(Xlsxrb::ParseError) { Xlsxrb.load(StringIO.new(bytes)) } }
    error = within(5, "read") { assert_raise(Xlsxrb::ParseError) { Xlsxrb.read(StringIO.new(bytes)) { |sheet| sheet.each_row(&:values) } } }
    assert_match(/nested more than 256 levels/, error.message)

    nested = %(<workbook #{SML_NS}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>#{"<a>" * 200}#{"</a>" * 200}</workbook>)
    assert_equal ["S"], Xlsxrb.load(StringIO.new(workbook_bytes("xl/workbook.xml" => nested))).sheets.map(&:name)
  end

  test "malformed XML in a part raises ParseError" do
    error = assert_raise(Xlsxrb::ParseError) { Xlsxrb.load(StringIO.new(workbook_bytes("xl/workbook.xml" => "<workbook><sheets>"))) }
    assert_match(/\AMalformed XML: /, error.message)
    assert_not_include error.message, "\n"
  end

  test "a row number above Excel's last row raises ParseError" do
    ["1048577", "4000000000", "9" * 200_000].each do |number|
      bytes = workbook_bytes(sheet: sheet_xml(%(<row r="#{number}"><c r="A1"><v>1</v></c></row>)))

      within(5, "load #{number.size}") { assert_raise(Xlsxrb::ParseError) { Xlsxrb.load(StringIO.new(bytes)) } }
      within(5, "read #{number.size}") do
        assert_raise(Xlsxrb::ParseError) { Xlsxrb.read(StringIO.new(bytes), pad_empty_rows: true) { |sheet| sheet.each_row(&:values) } }
      end
    end

    last_row = workbook_bytes(sheet: sheet_xml(%(<row r="1048576"><c r="A1048576"><v>1</v></c></row>)))
    assert_equal [1_048_575], Xlsxrb.load(StringIO.new(last_row)).sheets.first.rows.map(&:index)
  end

  test "cell values of more digits than a number has are not converted, and shared string indexes out of range are empty" do
    digits = "7" * 2_000_000
    cells = [
      %(<c r="A1"><v>1234567890123456789</v></c>),
      %(<c r="B1"><v>12345678901234567890123</v></c>),
      %(<c r="C1" t="n"><v>-12345678901234567890123</v></c>),
      %(<c r="D1"><v>#{digits}</v></c>),
      %(<c r="E1" t="s"><v>99999999999999999999999</v></c>),
      %(<c r="F1" t="s"><v>-1</v></c>),
      %(<c r="G1" t="s"><v>0</v></c>)
    ].join
    bytes = workbook_bytes(sheet: sheet_xml(%(<row r="1">#{cells}</row>)))

    values = within(5, "read") { Xlsxrb.read(StringIO.new(bytes)).first.each_row.first.values }
    # Integers keep every digit, as before; only text too long to be a number is left as text.
    assert_equal [1_234_567_890_123_456_789, 12_345_678_901_234_567_890_123, -12_345_678_901_234_567_890_123, digits, "", "", "a"], values
    assert_equal values, within(5, "load") { Xlsxrb.load(StringIO.new(bytes)).sheets.first.rows.first.cells.map(&:value) }

    typed = workbook_bytes(sheet: sheet_xml(%(<row r="1"><c r="A1" t="n"><v>#{digits}</v></c></row>)))
    assert_raise(Xlsxrb::ParseError) { Xlsxrb.load(StringIO.new(typed)).sheets.first.rows.first.cells.map(&:value) }
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

      within(5, "#{label} read") { assert_raise(Xlsxrb::ParseError, label) { Xlsxrb.read(StringIO.new(bytes)) { |sheet| sheet.each_row(&:values) } } }
      within(5, "#{label} load") { assert_raise(Xlsxrb::ParseError, label) { Xlsxrb.load(StringIO.new(bytes)).sheets.first.rows.each(&:cells) } }
    end

    cells = %(<c r="A1" t="inlineStr"><is><t>hi</t></is></c><c r="B1" s="0"/><x:c r="C1"><x:v>2</x:v></x:c> <c r="D1"/> )
    bytes = workbook_bytes(sheet: sheet_xml(%(<row r="1">#{cells}</row><row r="2"><c r="A2"/><c r="B2"/></row>)))
    assert_equal [["hi", nil, 2, nil], [nil, nil]], Xlsxrb.read(StringIO.new(bytes)).first.each_row.map(&:values)
    assert_equal([["hi", nil, 2, nil], [nil, nil]], Xlsxrb.load(StringIO.new(bytes)).sheets.first.rows.map { |row| row.cells.map(&:value) })
  end

  test "shared strings without text are read in time proportional to the part" do
    omit "Time limits do not hold under RBS runtime checks, which wrap every method call" if defined?(RBS::Test)

    count = 200_000
    empty = %(<sst #{SML_NS}>#{"<si></si>" * count}<si><t>last</t></si></sst>)
    runs = %(<sst #{SML_NS}><si>#{"<r><t>a</t></r>" * count}</si><si><t>x</t><rPh sb="0" eb="1"><t>y</t></rPh></si></sst>)

    strings = within(5, "empty") { Xlsxrb::Ooxml::SharedStringsParser.parse(empty) }
    assert_equal count + 1, strings.size
    assert_equal ["", "last"], strings.last(2).map(&:to_s)

    strings = within(5, "runs") { Xlsxrb::Ooxml::SharedStringsParser.parse(runs) }
    assert_equal ["a" * count, "x"], strings.map(&:to_s)
  end

  test "streamed rows are found in time proportional to the sheet when markup spans many chunks" do
    omit "Time limits do not hold under RBS runtime checks, which wrap every method call" if defined?(RBS::Test)

    megabyte = 1024 * 1024
    {
      "elements named like rows" => "<rowX" * (8 * megabyte / 5),
      "a row without an end tag" => %(<row r="1">#{"A" * (16 * megabyte)}),
      "a comment before the rows" => "<!--#{"x" * (16 * megabyte)}-->"
    }.each do |label, markup|
      bytes = workbook_bytes(sheet: sheet_xml(%(#{markup}<row r="5"><c r="A5"><v>1</v></c></row>)))
      rows = within(15, label) { Xlsxrb.read(StringIO.new(bytes)).first.each_row.map(&:values) }

      assert_equal(label == "a row without an end tag" ? 1 : [[1]], label == "a row without an end tag" ? rows.size : rows, label)
    end
  end

  test "a streamed row larger than the buffer limit raises ParseError" do
    bytes = workbook_bytes(sheet: sheet_xml(%(<row r="1">#{"A" * (51 * 1024 * 1024)})))

    error = assert_raise(Xlsxrb::ParseError) { Xlsxrb.read(StringIO.new(bytes)) { |sheet| sheet.each_row(&:values) } }
    assert_match(/exceeds 52428800 bytes/, error.message)
  end

  test "dimension stops searching a sheet that has neither a dimension nor sheet data" do
    omit "Time limits do not hold under RBS runtime checks, which wrap every method call" if defined?(RBS::Test)

    filler = %(<worksheet #{SML_NS}>#{"<a/>" * (16 * 1024 * 1024 / 4)}</worksheet>)
    filler_bytes = workbook_bytes(sheet: filler)
    assert_nil(within(10, "dimension") { Xlsxrb.read(StringIO.new(filler_bytes)).first.dimension })

    with_dimension = sheet_xml(%(<row r="1"><c r="A1"><v>1</v></c></row>), before: %(<dimension ref="A1:C3"/>))
    assert_equal "A1:C3", Xlsxrb.read(StringIO.new(workbook_bytes(sheet: with_dimension))).first.dimension
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

  test "style indexes outside the style table are read as unstyled" do
    styles = [
      %(<styleSheet #{SML_NS}><numFmts count="1"><numFmt numFmtId="164" formatCode="0.00"/></numFmts>),
      %(<cellStyleXfs count="1"><xf numFmtId="164"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" xfId="0"/>),
      %(<xf numFmtId="0" xfId="#{"9" * 30}"/><xf numFmtId="0" xfId="-1"/></cellXfs></styleSheet>)
    ].join
    cells = ["0", "1", "2", "9" * 30, "-1"].each_with_index.map { |s, i| %(<c r="#{("A".ord + i).chr}1" s="#{s}"><v>1.5</v></c>) }.join
    rels = [
      %(<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">),
      %(<Relationship Id="rId1" Type="#{REL_TYPE}/worksheet" Target="worksheets/sheet1.xml"/>),
      %(<Relationship Id="rId2" Type="#{REL_TYPE}/sharedStrings" Target="sharedStrings.xml"/>),
      %(<Relationship Id="rId3" Type="#{REL_TYPE}/styles" Target="styles.xml"/></Relationships>)
    ].join
    bytes = workbook_bytes(sheet: sheet_xml(%(<row r="1">#{cells}</row>)), "xl/styles.xml" => styles, "xl/_rels/workbook.xml.rels" => rels)

    # The first style takes its format from its parent; the others name no style or no parent.
    expected = ["1.50", "1.5", "1.5", "1.5", "1.5"]
    assert_equal expected, Xlsxrb.read(StringIO.new(bytes)).first.each_row.first.formatted_values
    assert_equal expected, Xlsxrb.load(StringIO.new(bytes)).sheets.first.rows.first.cells.map(&:formatted_value)
  end

  test "bytes that are not UTF-8 are read as U+FFFD" do
    cells = %(<c r="A1" t="inlineStr"><is><t>a\xFFb</t></is></c><c r="B1" t="str"><f>"\xFF"</f><v>\xFE</v></c><c r="C1" t="s"><v>0</v></c>)
    bytes = workbook_bytes(
      sheet: sheet_xml(%(<row r="1">#{cells}</row>), before: %(<dimension ref="A1:C1"/>)).b,
      "xl/sharedStrings.xml" => %(<sst #{SML_NS}><si><t>\xC3</t></si></sst>).b
    )
    expected = ["a�b", "�", "�"]

    streamed = Xlsxrb.read(StringIO.new(bytes)).first
    assert_equal "A1:C1", streamed.dimension
    assert_equal expected, streamed.each_row.first.values
    loaded = Xlsxrb.load(StringIO.new(bytes)).sheets.first.rows.first.cells
    assert_equal expected, loaded.map(&:value)
    assert_equal "\"�\"", loaded[1].formula.to_s
    assert_nothing_raised { Xlsxrb.write(StringIO.new, Xlsxrb.load(StringIO.new(bytes))) }

    # Outside element content the part is not well-formed, which is a ParseError at most.
    damaged = workbook_bytes("xl/_rels/workbook.xml.rels" => %(<Relationships \xFF/>).b, "xl/workbook.xml" => %(<workbook \xFF="1"/>).b)
    [-> { Xlsxrb.load(StringIO.new(damaged)) }, -> { Xlsxrb.read(StringIO.new(damaged)) { |sheet| sheet.each_row.to_a } }].each do |reader|
      reader.call
    rescue Xlsxrb::Error
      nil
    end
  end

  test "damaged parts raise no error other than Xlsxrb::Error" do
    styles = [
      %(<styleSheet #{SML_NS}><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>),
      %(<cellStyleXfs count="1"><xf numFmtId="0"/></cellStyleXfs>),
      %(<cellXfs count="2"><xf numFmtId="0" xfId="0"/><xf numFmtId="164" xfId="0"/></cellXfs></styleSheet>)
    ].join
    rows = %(<row r="1" ht="20"><c r="A1" t="s"><v>0</v></c><c r="B1" s="1"><v>45000</v></c><c r="C1" t="b"><v>1</v></c></row>) +
           %(<row r="2"><c r="A2" t="inlineStr"><is><t>x</t></is></c><c r="B2"><f>A1+1</f><v>2</v></c><c r="C2" t="d"><v>2024-01-02T03:04:05</v></c></row>)
    after = %(<mergeCells count="1"><mergeCell ref="A5:B6"/></mergeCells><hyperlinks><hyperlink ref="A1" location="S!A2"/></hyperlinks>)
    parts = {
      "xl/styles.xml" => styles,
      "xl/sharedStrings.xml" => %(<sst #{SML_NS}><si><t>a</t></si></sst>),
      "xl/workbook.xml" => %(<workbook #{SML_NS}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>),
      "xl/worksheets/sheet1.xml" => sheet_xml(rows, before: %(<dimension ref="A1:C2"/><cols><col min="1" max="2" width="12"/></cols>), after: after)
    }
    replacements = ["-1", "0", "9" * 25, "1e999", '"', ">", "<", "&", "\xFF".b, ""]
    random = Random.new(20_261_117)

    400.times do |round|
      name = parts.keys.sample(random: random)
      part = parts[name].b
      position = random.rand(part.bytesize)
      part[position, 1 + random.rand(8)] = replacements.sample(random: random).b
      bytes = workbook_bytes(parts.merge(name => part))
      begin
        within(10, "round #{round}") do
          Xlsxrb.read(StringIO.new(bytes)) do |sheet|
            sheet.dimension
            sheet.each_row(&:formatted_values)
          end
          Xlsxrb.load(StringIO.new(bytes)).sheets.each { |sheet| sheet.rows.each { |row| row.cells.each(&:formatted_value) } }
        end
      rescue Xlsxrb::Error
        next
      rescue StandardError => e
        flunk "round #{round} damaged #{name} at #{position}: #{e.class}: #{e.message}"
      end
    end
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
      defaults.merge(parts).each { |name, content| zip.add_binary_entry(name, content.b) }
    end
    io.string
  end
end
