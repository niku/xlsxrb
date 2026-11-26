# frozen_string_literal: true

require "test_helper"
require "pbt"
require "rexml/document"
require "nokogiri"
require "stringio"
require "timeout"
require "digest"

class SecurityAndFuzzingPbtTest < Test::Unit::TestCase
  ADVERSARIAL_SAFE_FRAGMENTS = [
    '1" injected="1"><injected/><x y="',
    '" injected="1" ',
    "' injected='1' ",
    "</t><injected/><t>",
    "</v></c><injected/><c><v>",
    "</f><injected/>",
    "</si><injected/>",
    "</text><injected/>",
    "</vt:lpwstr><injected/>",
    "</x:Anchor><injected/>",
    "&amp;",
    "&lt;injected/&gt;",
    "<![CDATA[<injected/>]]>",
    "<!--injected-->",
    "a&b<c>d\"e'f",
    "日本語🚀",
    "line1\nline2"
  ].freeze

  ADVERSARIAL_CONTROL_FRAGMENTS = (
    ADVERSARIAL_SAFE_FRAGMENTS + [
      "\x00",
      "\x01",
      "\x07",
      "\x0B",
      "\x0C",
      "\x1F",
      "_x000A_",
      "_x0000_",
      "  padded  "
    ]
  ).freeze

  ZIP_SENTINELS = [
    "\x00\x00\x00\x00".b,
    "\xFF\xFF\xFF\xFF".b,
    "\x7F\xFF\xFF\xFF".b,
    "PK\x03\x04".b,
    "PK\x01\x02".b,
    "PK\x05\x06".b
  ].freeze

  CFB_SENTINELS = [
    [0].pack("V"),
    [1].pack("V"),
    [0xFFFFFFFD].pack("V"),
    [0xFFFFFFFE].pack("V"),
    [0xFFFFFFFF].pack("V"),
    [0x7FFFFFFF].pack("V")
  ].freeze

  INVALID_UTF8_CHUNKS = [
    "\xFF\xFE".b,
    "\xC0\xAF".b,
    "\xE3\x81".b,
    "\x80".b,
    "\xF4\x90\x80\x80".b
  ].freeze

  SEED_WORKSHEET_XML = <<~XML
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <dimension ref="A1:D3"/>
      <sheetData>
        <row r="1">
          <c r="A1" t="s"><v>0</v></c>
          <c r="B1" t="n"><v>42.5</v></c>
          <c r="C1" t="b"><v>1</v></c>
        </row>
        <row r="2">
          <c r="A2" t="inlineStr"><is><t xml:space="preserve"> Hello 世界 🚀 </t></is></c>
          <c r="B2" t="inlineStr"><is><r><rPr><b/></rPr><t>Rich</t></r><r><t>Text</t></r></is></c>
          <c r="C2" t="e"><v>#N/A</v></c>
          <c r="D2" t="str"><f>1+1</f><v>two</v></c>
        </row>
      </sheetData>
    </worksheet>
  XML

  SEED_SST_XML = <<~XML
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
      <si><t xml:space="preserve">alpha &amp; omega</t></si>
      <si><r><rPr><b/></rPr><t>日本語</t></r><r><t>テキスト</t></r><rPh sb="0" eb="3"><t>ニホンゴ</t></rPh></si>
    </sst>
  XML

  SEED_STYLES_XML = <<~XML
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <numFmts count="1"><numFmt numFmtId="164" formatCode="0.00_ "/></numFmts>
      <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
      <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
      <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
      <cellXfs count="1"><xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/></cellXfs>
    </styleSheet>
  XML

  FUZZ_TIMEOUT_SECONDS = 1.0
  STANDARD_SEED_MUTEX = Mutex.new

  def setup
    omit "Skip property-based tests under RBS runtime testing to prevent OOM / slowdown" if defined?(RBS::Test)
  end

  def silence_deprecation_stderr
    STANDARD_SEED_MUTEX.synchronize do
      original_stderr = $stderr
      $stderr = StringIO.new
      yield
    ensure
      $stderr = original_stderr
    end
  end

  def populate_adversarial_workbook(workbook, payload_one, payload_two, raw_payload, sheet_name, custom_key)
    workbook.core_properties(title: payload_one, creator: payload_two, subject: payload_one, description: payload_two, keywords: payload_one)
    workbook.app_properties(company: payload_one, manager: payload_two)
    workbook.custom_property(custom_key, payload_two, type: :string)

    quoted_sheet = "'#{sheet_name.gsub("'", "''")}'"

    workbook.sheet(sheet_name) do |s|
      s.style(:adv, bold: true)
      s.row([payload_one, "Sales"])
      s.row([payload_two, 42], cell_styles: [:adv, nil])
      s.row([
              Xlsxrb.rich_text({ text: payload_one, font: { bold: true, name: payload_two[0, 20] } }, { text: payload_two }),
              Xlsxrb.formula("1+1", cached_value: payload_one)
            ])
      s.row([raw_payload, Xlsxrb.rich_text({ text: raw_payload })])

      s.hyperlink("A1", "https://example.com/?q=#{payload_one}", display: payload_two, tooltip: payload_one)
      s.comment("A2", payload_one, author: payload_two)
      s.threaded_comment("A3", payload_one, author: payload_two, id: "{11111111-1111-1111-1111-111111111111}")
      s.header_footer(odd_header: payload_one, odd_footer: payload_two)
      s.data_validation(
        "C1:C5",
        type: :list,
        formula1: '"A,B"',
        prompt_title: payload_one[0, 30],
        prompt: payload_two,
        error_title: payload_two[0, 30],
        error: payload_one
      )
      s.conditional_format("D1:D5", type: :contains_text, text: payload_one, dxf: { fill: { fg_color: "FFFFFF00" } })
      s.table("A1:B2", name: "AdvTable", display_name: "AdvTable", columns: [{ name: payload_one }, { name: "Sales" }])
      s.chart(
        type: :bar,
        title: payload_one,
        from_column: 5,
        from_row: 0,
        to_column: 10,
        to_row: 10,
        series: [
          {
            name: payload_two,
            cat_ref: "#{quoted_sheet}!$A$2:$A$2",
            val_ref: "#{quoted_sheet}!$B$2:$B$2"
          }
        ],
        category_axis: { title: payload_one },
        value_axis: { title: payload_two }
      )
      s.shape(preset: "rect", text: payload_one, from_column: 5, from_row: 12, to_column: 9, to_row: 16)
      s.pivot_table(
        "A1:B2",
        name: payload_one,
        field_names: [payload_one, "Sales"],
        row_fields: [payload_one],
        data_fields: [{ field: "Sales", name: payload_two }],
        ref: "F20"
      )
    end

    workbook.defined_name("AdvDefined", sheet: sheet_name, columns: 0, rows: 0..1)
  end

  def assert_well_formed_and_clean_xml(bytes, mode_label)
    reader = Xlsxrb::Ooxml::ZipReader.open(StringIO.new(bytes))
    names = reader.entry_names
    xml_parts = names.grep(/\.(xml|rels|vml)\z/)

    expected_parts = %w[
      docProps/core.xml
      docProps/app.xml
      docProps/custom.xml
      xl/worksheets/sheet1.xml
      xl/comments1.xml
      xl/drawings/vmlDrawing1.vml
      xl/threadedComments/threadedComment1.xml
      xl/persons/person.xml
      xl/tables/table1.xml
      xl/charts/chart1.xml
      xl/drawings/drawing1.xml
      xl/pivotTables/pivotTable1.xml
    ]
    expected_parts.each do |part|
      assert_include xml_parts, part, "#{mode_label}: missing #{part}"
    end

    xml_parts.each do |part|
      content = reader.read_entry(part)
      rexml_doc = REXML::Document.new(content)
      assert_empty REXML::XPath.match(rexml_doc, "//*[local-name()='injected'] | //@injected"),
                   "#{mode_label} REXML injection in #{part}"

      nokogiri_doc = Nokogiri::XML(content, &:strict)
      assert_empty nokogiri_doc.xpath("//*[local-name()='injected'] | //@injected"),
                   "#{mode_label} Nokogiri injection in #{part}"
    end
  ensure
    reader&.close
  end

  def assert_adversarial_roundtrip(bytes, payload_one, payload_two, sheet_name, custom_key, mode_label)
    wb = Xlsxrb.load(StringIO.new(bytes))
    assert_equal payload_one, wb.core_properties.title, "#{mode_label}: core title"
    assert_equal payload_two, wb.core_properties.creator, "#{mode_label}: core creator"
    assert_equal payload_one, wb.app_properties.company, "#{mode_label}: app company"
    assert_equal payload_two, wb.app_properties.manager, "#{mode_label}: app manager"
    assert_equal payload_two, wb.custom_properties[custom_key]&.value, "#{mode_label}: custom property"

    ws = wb[sheet_name]
    assert_not_nil ws, "#{mode_label}: sheet missing"
    assert_equal payload_one, ws["A1"].value, "#{mode_label}: A1"
    assert_equal payload_two, ws["A2"].value, "#{mode_label}: A2"
    assert_equal(payload_one + payload_two, ws["A3"].value.to_s, "#{mode_label}: A3 RichText")
    assert_equal "https://example.com/?q=#{payload_one}", ws.hyperlink("A1")&.url, "#{mode_label}: hyperlink url"
    assert_equal payload_two, ws.hyperlink("A1")&.display, "#{mode_label}: hyperlink display"
    assert_equal payload_one, ws.hyperlink("A1")&.tooltip, "#{mode_label}: hyperlink tooltip"
    assert_equal payload_one, ws.comment("A2")&.text, "#{mode_label}: comment text"
    assert_equal payload_two, ws.comment("A2")&.author, "#{mode_label}: comment author"
    assert_equal payload_one, ws.threaded_comments.first&.text, "#{mode_label}: threaded comment text"
    assert_equal payload_one, ws.header_footer.odd_header, "#{mode_label}: odd_header"
    assert_equal payload_two, ws.header_footer.odd_footer, "#{mode_label}: odd_footer"

    dv = ws.data_validations.first
    assert_not_nil dv, "#{mode_label}: data validation missing"
    assert_equal payload_one[0, 30].gsub("\n", "_x000a_"), dv.prompt_title, "#{mode_label}: dv prompt_title"
    assert_equal payload_two.gsub("\n", "_x000a_"), dv.prompt, "#{mode_label}: dv prompt"
    assert_equal payload_two[0, 30].gsub("\n", "_x000a_"), dv.error_title, "#{mode_label}: dv error_title"
    assert_equal payload_one.gsub("\n", "_x000a_"), dv.error, "#{mode_label}: dv error"

    assert_equal payload_one, ws.charts.first&.title, "#{mode_label}: chart title"
    assert_equal payload_one, ws.shapes.first&.text, "#{mode_label}: shape text"

    streamed_rows = nil
    Xlsxrb.read(StringIO.new(bytes)) do |sheet|
      next unless sheet.name == sheet_name

      streamed_rows = sheet.each_row.map(&:values)
    end
    assert_equal payload_one, streamed_rows[0][0], "#{mode_label}: streamed A1"
    assert_equal payload_two, streamed_rows[1][0], "#{mode_label}: streamed A2"
  end

  def cached_archive_seeds
    @cached_archive_seeds ||= begin
      minimal = Xlsxrb.write do |w|
        w.sheet("Minimal") do |s|
          s.row(["Alpha", 1, true])
          s.row(["Beta", 2.5, "日本語🚀"])
        end
      end

      rich = Xlsxrb.write(shared_strings: false) do |w|
        w.sheet("Rich") do |s|
          s.style(:hdr, bold: true, fill_color: "DDEEFF")
          s.row(%w[Item Total], cell_styles: %i[hdr hdr])
          s.row([Xlsxrb.rich_text({ text: "Hi", font: { bold: true } }, { text: " there" }), Xlsxrb.formula("1+2", cached_value: 3)])
          s.hyperlink("A1", "https://example.com", display: "Link")
          s.comment("A2", "Note", author: "Tester")
          s.chart(type: :bar, title: "Sales", from_column: 3, from_row: 0, to_column: 7, to_row: 8,
                  series: [{ name: "S", cat_ref: "Rich!$A$1:$A$2", val_ref: "Rich!$B$1:$B$2" }])
        end
      end

      agile = Xlsxrb.write(password: "secret", encryption_mode: :agile) do |w|
        w.sheet("Secret") { |s| s.row(["Protected", 123]) }
      end

      standard = nil
      silence_deprecation_stderr do
        standard = Xlsxrb.write(password: "secret", encryption_mode: :standard) do |w|
          w.sheet("Legacy") { |s| s.row(["Standard", 456]) }
        end
      end

      {
        minimal: [minimal.b, nil],
        rich: [rich.b, nil],
        agile: [agile.b, "secret"],
        standard: [standard.b, "secret"]
      }
    end
  end

  def cached_cfb_seeds
    @cached_cfb_seeds ||= begin
      mini = Xlsxrb::Ooxml::Cfb::Writer.write(
        "EncryptionInfo" => ("E" * 256).b,
        "EncryptedPackage" => ("P" * 1024).b
      ).b

      regular = Xlsxrb::Ooxml::Cfb::Writer.write(
        "EncryptionInfo" => ("E" * 600).b,
        "EncryptedPackage" => ("P" * 12_000).b,
        "ExtraStream" => ("X" * 5_000).b
      ).b

      { mini: mini, regular: regular }
    end
  end

  def mutate_binary_bytes(source, ops_count, rng, sentinels:)
    mode = rng.rand(10)
    return random_binary_blob(rng.rand(0..512), rng) if mode == 9

    buf = source.dup.b
    ops_count.times do
      break if buf.empty? && rng.rand(2).zero?

      case rng.rand(5)
      when 0
        # Truncate, biasing toward ZIP/CFB header and tail boundaries
        limit = if buf.bytesize > 64 && rng.rand(3).zero?
                  rng.rand(0..64)
                else
                  rng.rand(0..buf.bytesize)
                end
        buf = buf.byteslice(0, limit) || +"".b
      when 1
        # Flip or overwrite 1..8 bytes
        next if buf.empty?

        rng.rand(1..8).times do
          pos = rng.rand(buf.bytesize)
          val = case rng.rand(3)
                when 0 then rng.rand(256)
                when 1 then 0x00
                else 0xFF
                end
          buf.setbyte(pos, val)
        end
      when 2
        # Splice a boundary 4-byte sentinel
        next if buf.bytesize < 4

        pos = rng.rand(0..(buf.bytesize - 4))
        sentinel = sentinels[rng.rand(sentinels.size)]
        buf[pos, sentinel.bytesize] = sentinel
      when 3
        # Prepend or append random garbage
        garbage = random_binary_blob(rng.rand(1..32), rng)
        buf = rng.rand(2).zero? ? (garbage + buf) : (buf + garbage)
      else
        # Excise a slice from the middle
        next if buf.bytesize < 2

        pos = rng.rand(buf.bytesize)
        len = rng.rand(1..[buf.bytesize - pos, 64].min)
        buf.slice!(pos, len)
      end
    end
    buf
  end

  def random_binary_blob(length, rng)
    Array.new(length) { rng.rand(256) }.pack("C*").b
  end

  def mutate_xml_bytes(seed_xml, kind, rng)
    case kind
    when :truncate
      pos = rng.rand(0..seed_xml.bytesize)
      seed_xml.b.byteslice(0, pos) || +"".b
    when :invalid_utf8
      buf = seed_xml.b.dup
      rng.rand(1..6).times do
        chunk = INVALID_UTF8_CHUNKS[rng.rand(INVALID_UTF8_CHUNKS.size)]
        pos = rng.rand(0..buf.bytesize)
        buf.insert(pos, chunk)
      end
      buf
    when :corrupt_delimiters
      buf = seed_xml.b.dup
      delimiters = ["<", ">", '"', "'", "/", "&"]
      rng.rand(1..10).times do
        break if buf.empty?

        pos = rng.rand(buf.bytesize)
        if rng.rand(2).zero?
          buf.slice!(pos, 1)
        else
          buf.setbyte(pos, delimiters[rng.rand(delimiters.size)].getbyte(0))
        end
      end
      buf
    when :deep_nesting
      depth = rng.rand(50..180)
      "#{"<a>" * depth}payload#{"</a>" * depth}".b
    when :unclosed_tags
      count = rng.rand(50..300)
      "<worksheet><sheetData><row r=\"1\">#{'<c r="A1" ' * count}".b
    else
      random_binary_blob(rng.rand(0..256), rng)
    end
  end

  test "property: adversarial XML breakout and element-closing payloads cannot inject XML across workbook features" do
    frag_arb = Pbt.constant_from(*ADVERSARIAL_SAFE_FRAGMENTS)
    safe_payload_arb = Pbt.array(frag_arb, min: 2, max: 4).map { |parts| "x#{parts.join("_")}z" }

    raw_frag_arb = Pbt.constant_from(*ADVERSARIAL_CONTROL_FRAGMENTS)
    raw_payload_arb = Pbt.array(raw_frag_arb, min: 2, max: 4).map(&:join)

    Pbt.assert(num_runs: 25) do
      Pbt.property(safe_payload_arb, safe_payload_arb, raw_payload_arb, Pbt.boolean) do |p1, p2, raw_p, shared_strings|
        sheet_name = Xlsxrb.sanitize_sheet_name("S_<&>\"_#{p1}")
        custom_key = "Note_#{Digest::MD5.hexdigest(p1)[0, 8]}_#{p1[0, 12]}"

        streamed_bytes = Xlsxrb.write(shared_strings: shared_strings) do |w|
          populate_adversarial_workbook(w, p1, p2, raw_p, sheet_name, custom_key)
        end

        built_wb = Xlsxrb.build do |w|
          populate_adversarial_workbook(w, p1, p2, raw_p, sheet_name, custom_key)
        end
        built_bytes = Xlsxrb.write(built_wb, shared_strings: shared_strings)

        [["streamed", streamed_bytes], ["built", built_bytes]].each do |mode_label, bytes|
          assert_well_formed_and_clean_xml(bytes, "#{mode_label} (shared_strings=#{shared_strings})")
          assert_adversarial_roundtrip(bytes, p1, p2, sheet_name, custom_key, "#{mode_label} (shared_strings=#{shared_strings})")
        end
      end
    end
  end

  test "property: truncated and byte-mutated XLSX/ZIP archives fail cleanly without crashes or hangs" do
    seeds = cached_archive_seeds
    kind_arb = Pbt.constant_from(*seeds.keys)

    Pbt.assert(num_runs: 100) do
      Pbt.property(kind_arb, Pbt.integer(min: 1, max: 6), Pbt.integer(min: 0, max: (2**31) - 1)) do |kind, ops_count, seed|
        source, password = seeds.fetch(kind)
        rng = Random.new(seed)
        mutated = mutate_binary_bytes(source, ops_count, rng, sentinels: ZIP_SENTINELS)

        Timeout.timeout(FUZZ_TIMEOUT_SECONDS) do
          begin
            wb = Xlsxrb.load(StringIO.new(mutated), password: password, max_uncompressed_size: 2 * 1024 * 1024)
            wb.sheets.each { |s| s.rows.each(&:values) }
          rescue Xlsxrb::Error
            # Expected rejection for corrupted or malformed archives.
          end

          begin
            Xlsxrb.read(StringIO.new(mutated), password: password, max_uncompressed_size: 2 * 1024 * 1024) do |sheet|
              sheet.each_row(&:values)
            end
          rescue Xlsxrb::Error
            # Expected rejection for corrupted or malformed archives.
          end
        end
      end
    end
  end

  test "property: truncated and byte-mutated CFB containers fail cleanly in Cfb::Reader" do
    seeds = cached_cfb_seeds
    kind_arb = Pbt.constant_from(*seeds.keys)

    Pbt.assert(num_runs: 100) do
      Pbt.property(kind_arb, Pbt.integer(min: 1, max: 6), Pbt.integer(min: 0, max: (2**31) - 1)) do |kind, ops_count, seed|
        source = seeds.fetch(kind)
        rng = Random.new(seed)
        mutated = mutate_binary_bytes(source, ops_count, rng, sentinels: CFB_SENTINELS)

        Timeout.timeout(FUZZ_TIMEOUT_SECONDS) do
          reader = Xlsxrb::Ooxml::Cfb::Reader.new(mutated)
          reader.stream_names.each { |name| reader.read_stream(name) }
        rescue Xlsxrb::Error
          # Expected rejection for corrupted CFB streams or headers.
        end
      end
    end
  end

  test "property: malformed, truncated, and invalid-UTF-8 worksheet, SST, and styles XML fail cleanly without crashes or ReDoS" do
    xml_seeds = {
      worksheet: SEED_WORKSHEET_XML,
      sst: SEED_SST_XML,
      styles: SEED_STYLES_XML
    }.freeze
    seed_arb = Pbt.constant_from(*xml_seeds.keys)
    mutation_arb = Pbt.constant_from(:truncate, :invalid_utf8, :corrupt_delimiters, :deep_nesting, :unclosed_tags, :random)

    Pbt.assert(num_runs: 100) do
      Pbt.property(seed_arb, mutation_arb, Pbt.integer(min: 0, max: (2**31) - 1)) do |xml_kind, mut_kind, seed|
        rng = Random.new(seed)
        mutated = mutate_xml_bytes(xml_seeds.fetch(xml_kind), mut_kind, rng)
        sst = %w[alpha beta]
        row_count = 0

        Timeout.timeout(FUZZ_TIMEOUT_SECONDS) do
          [
            -> { Xlsxrb::Ooxml::WorksheetParser.parse(mutated, shared_strings: sst) },
            -> { Xlsxrb::Ooxml::WorksheetParser.each_row(mutated, shared_strings: sst) { |_row| row_count += 1 } },
            -> { Xlsxrb::Ooxml::WorksheetParser.each_row(StringIO.new(mutated), shared_strings: sst) { |_row| row_count += 1 } },
            -> { Xlsxrb::Ooxml::WorksheetParser.each_event(mutated, shared_strings: sst) { |_ev| row_count += 1 } },
            -> { Xlsxrb::Ooxml::SharedStringsParser.parse(mutated) },
            -> { Xlsxrb::Ooxml::StylesParser.parse(mutated) },
            -> { Xlsxrb::Ooxml::XmlParser.parse(mutated, Xlsxrb::Ooxml::XmlParser::BaseListener.new) }
          ].each do |invoke|
            invoke.call
          rescue Xlsxrb::Error
            # Expected rejection for malformed XML inputs.
          end
        end
        assert_operator row_count, :>=, 0
      end
    end
  end
end
