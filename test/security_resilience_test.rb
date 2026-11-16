# frozen_string_literal: true

require "test_helper"
require "timeout"

class SecurityResilienceTest < Test::Unit::TestCase
  # Helper to assert regex matching finishes within a strict time limit (tens of ms)
  def assert_redos_resistant(label, max_duration: 0.1, &)
    t0 = Process.clock_gettime(Process::CLOCK_MONOTONIC)
    Timeout.timeout(max_duration, &)
    elapsed = Process.clock_gettime(Process::CLOCK_MONOTONIC) - t0
    assert_operator elapsed, :<, max_duration, "#{label} took #{elapsed.round(4)}s, exceeding #{max_duration}s (possible ReDoS)"
  end

  # 1. ReDoS audit on WorksheetParser cell regular expressions
  def test_worksheet_parser_cell_fast_re_redos_resistance
    matches = 0
    # Unclosed cell tag with 20,000 characters of attributes
    malicious_unclosed = "<c r=\"A1\" #{'s="1" ' * 2000} t=\"s\""
    assert_redos_resistant("CELL_FAST_RE unclosed attributes") do
      malicious_unclosed.scan(Xlsxrb::Ooxml::WorksheetParser::CELL_FAST_RE) { |*m| matches += m.size }
    end

    # Unclosed inner value tag with 20,000 characters
    malicious_inner = "<c r=\"A1\"><v>#{"9" * 20_000}"
    assert_redos_resistant("CELL_FAST_RE unclosed value tag") do
      malicious_inner.scan(Xlsxrb::Ooxml::WorksheetParser::CELL_FAST_RE) { |*m| matches += m.size }
    end

    # Unclosed inlineStr tag
    malicious_is = "<c r=\"A1\" t=\"inlineStr\"><is><t>#{"A" * 20_000}"
    assert_redos_resistant("CELL_FAST_RE unclosed is/t tag") do
      malicious_is.scan(Xlsxrb::Ooxml::WorksheetParser::CELL_FAST_RE) { |*m| matches += m.size }
    end
    assert_operator matches, :>=, 0
  end

  def test_worksheet_parser_cell_generic_re_redos_resistance
    matches = 0
    malicious_unclosed = "<c #{'custom="val" ' * 2000} r=\"A1\""
    assert_redos_resistant("CELL_GENERIC_RE unclosed tag") do
      malicious_unclosed.scan(Xlsxrb::Ooxml::WorksheetParser::CELL_GENERIC_RE) { |*m| matches += m.size }
    end

    malicious_f = "<c r=\"A1\"><f>#{"SUM(" * 3000}"
    assert_redos_resistant("CELL_GENERIC_RE unclosed formula tag") do
      malicious_f.scan(Xlsxrb::Ooxml::WorksheetParser::CELL_GENERIC_RE) { |*m| matches += m.size }
    end
    assert_operator matches, :>=, 0
  end

  # 2. ReDoS audit on NumberFormatter patterns
  def test_number_formatter_redos_resistance
    # Deeply nested bracket syntax
    nested_brackets = "#{"[" * 5000}Red#{"ready" * 1000}"
    assert_redos_resistant("NumberFormatter nested brackets") do
      Xlsxrb::NumberFormatter.format(123.45, nested_brackets)
    end

    # Extremely long format code with repeated zeroes and scientific notations
    long_zeros = "0.#{"0" * 20_000}E+00"
    assert_redos_resistant("NumberFormatter long scientific format") do
      Xlsxrb::NumberFormatter.format(123.45, long_zeros)
    end

    # Repeated escaped characters
    long_escapes = "#{"\\$" * 10_000}#,##0.00"
    assert_redos_resistant("NumberFormatter long escaped prefix") do
      Xlsxrb::NumberFormatter.format(42, long_escapes)
    end
  end

  # 3. ReDoS audit on XML Builder escaping / unescaping
  def test_xml_builder_escape_and_unescape_redos_resistance
    # Long strings with unclosed entity candidates
    unclosed_entities = "&" * 50_000
    assert_redos_resistant("XmlBuilder.unescape repetitive ampersands") do
      Xlsxrb::Ooxml::XmlBuilder.unescape(unclosed_entities)
    end

    # Huge mixed content
    large_special_chars = ("<>&\"'abc" * 10_000)
    assert_redos_resistant("XmlBuilder.escape large mixed special characters") do
      Xlsxrb::Ooxml::XmlBuilder.escape(large_special_chars)
    end
  end

  # 4. ReDoS audit on Relationships scanning
  def test_relationship_tag_scanning_redos_resistance
    matches = 0
    malicious_rel = "<Relationship #{'Id="rId" ' * 3000}"
    assert_redos_resistant("Relationships scanning unclosed tag") do
      malicious_rel.scan(%r{<Relationship\b([^>]*)/?>}) { |*m| matches += m.size }
    end
    assert_operator matches, :>=, 0
  end

  # 5. Coordinate patterns ReDoS audit
  def test_coordinate_address_pattern_redos_resistance
    long_coord = ("A" * 10_000) + ("9" * 10_000)
    assert_redos_resistant("CELL_ADDRESS_PATTERN with 20k characters") do
      Xlsxrb::Ooxml::Writer::CELL_ADDRESS_PATTERN.match(long_coord)
    end
  end

  # 6. Formula injection mitigation test (leading formula characters escaping)
  def test_formula_injection_defense
    formula_values = [
      "=cmd|'/c calc'!A1",
      "+10+20",
      "-10-20",
      "@SUM(1,2)"
    ]

    Tempfile.create(["formula_injection", ".xlsx"]) do |f|
      Xlsxrb.write(f.path) do |w|
        w.sheet("Sheet1") do |s|
          formula_values.each do |val|
            s.row([val])
          end
        end
      end

      # Strings are always written as text cells, never as formulas, so Excel shows them
      # verbatim and does not evaluate command prefixes.
      sheet_xml = Xlsxrb::Ooxml::ZipReader.open(f.path).read_entry("xl/worksheets/sheet1.xml")
      assert_not_match(/<f[ >]/, sheet_xml)
      assert_equal(formula_values.size, sheet_xml.scan(' t="s"').size)

      wb = Xlsxrb.load(f.path)
      rows = wb.sheets[0].rows
      formula_values.each_with_index do |expected, idx|
        assert_equal expected, rows[idx].cells[0].value
      end
    end
  end

  # 7. Loading xlsxrb must not loosen REXML entity expansion limits set by the application
  def test_require_preserves_rexml_entity_expansion_limits
    script = <<~RUBY
      require "rexml/document"
      REXML::Security.entity_expansion_limit = 0
      REXML::Security.entity_expansion_text_limit = 512
      require "xlsxrb"
      print REXML::Security.entity_expansion_limit, ",", REXML::Security.entity_expansion_text_limit
      REXML::Security.entity_expansion_limit = 3
      print ",", REXML::Security.entity_expansion_limit
    RUBY
    lib_dir = File.expand_path("../lib", __dir__)
    stdout, stderr, status = Open3.capture3(RbConfig.ruby, "-I", lib_dir, "-e", script)
    assert_true status.success?, stderr
    assert_equal "0,512,3", stdout
  end

  # Values that reach attribute values, relationship targets, and ZIP entry names.
  ATTRIBUTE_BREAKOUT = '1" injected="1"><injected/><x y="'

  def test_image_extension_cannot_reach_entry_names_or_relationships
    ["png/../../../evil", 'png"/><Relationship Id="rId9" Target="file:///x" TargetMode="External', "p\u0001g", ""].each do |ext|
      error = assert_raise(ArgumentError, ext.inspect) do
        write_both_modes { |sheet| sheet.image("PNGDATA", ext: ext) }
      end
      assert_match(/Invalid image extension/, error.message)

      assert_raise(ArgumentError, ext.inspect) do
        Xlsxrb.write(StringIO.new) do |w|
          w.sheet("S") do |sheet|
            sheet.background_image = { data: "PNGDATA", ext: ext }
            sheet.row([1])
          end
        end
      end
    end

    write_both_modes { |sheet| sheet.image("PNGDATA", ext: "jpeg") }.each do |bytes|
      assert_include entry_names(bytes), "xl/media/image1.jpeg"
    end
  end

  def test_zip_writer_rejects_entry_names_that_escape_the_extraction_directory
    ["../evil", "/etc/passwd", "xl/../../evil", "xl/media/..", "xl\\media\\a.png", "xl/a\u0000.xml", "xl/a\n.xml", ""].each do |name|
      Xlsxrb::Ooxml::ZipWriter.open(StringIO.new) do |zip|
        assert_raise(Xlsxrb::ZipError, name.inspect) { zip.add_entry(name, "x") }
        assert_raise(Xlsxrb::ZipError, name.inspect) { zip.add_binary_entry(name, "x") }
        assert_raise(Xlsxrb::ZipError, name.inspect) { zip.start_entry(name) }
        assert_raise(Xlsxrb::ZipError, name.inspect) do
          zip.copy_raw_entry(name, StringIO.new("x"), data_offset: 0, compressed_size: 1, uncompressed_size: 1, crc32: 0, method: 0)
        end
      end
    end

    io = StringIO.new
    Xlsxrb::Ooxml::ZipWriter.open(io) do |zip|
      zip.add_entry("[Content_Types].xml", "x")
      zip.add_entry("xl/a..b/c...xml", "x")
    end
    assert_equal ["[Content_Types].xml", "xl/a..b/c...xml"], entry_names(io.string)
  end

  def test_formula_attributes_are_escaped
    omit "Passes values the signatures do not allow, which RBS runtime checks reject first" if defined?(RBS::Test)

    files = write_both_modes do |sheet|
      sheet.row([Xlsxrb.array_formula("1+1", ref: ATTRIBUTE_BREAKOUT)])
      sheet.row([Xlsxrb.shared_formula("1+1", shared_index: ATTRIBUTE_BREAKOUT, ref: "A2:A3")])
      sheet.row([{ formula: "1+1", type: :array, ref: ATTRIBUTE_BREAKOUT }])
    end
    files.each do |bytes|
      doc = REXML::Document.new(read_part(bytes, "xl/worksheets/sheet1.xml"))
      formulas = REXML::XPath.match(doc, "//*[local-name()='f']")

      assert_equal 3, formulas.size
      assert_equal ATTRIBUTE_BREAKOUT, formulas[0].attributes["ref"]
      assert_equal ATTRIBUTE_BREAKOUT, formulas[1].attributes["si"]
      assert_equal ATTRIBUTE_BREAKOUT, formulas[2].attributes["ref"]
      assert_empty REXML::XPath.match(doc, "//*[local-name()='injected'] | //@injected")
    end
  end

  def test_row_outline_level_is_escaped
    omit "Passes values the signatures do not allow, which RBS runtime checks reject first" if defined?(RBS::Test)

    write_both_modes { |sheet| sheet.row([1], outline_level: ATTRIBUTE_BREAKOUT) }.each do |bytes|
      doc = REXML::Document.new(read_part(bytes, "xl/worksheets/sheet1.xml"))
      row = REXML::XPath.first(doc, "//*[local-name()='row']")

      assert_equal ATTRIBUTE_BREAKOUT, row.attributes["outlineLevel"]
      assert_nil row.attributes["injected"]
    end
  end

  def test_threaded_comment_attributes_are_escaped
    files = write_both_modes do |sheet|
      sheet.row([1])
      sheet.threaded_comment("A1", "hi", id: ATTRIBUTE_BREAKOUT, parent_id: ATTRIBUTE_BREAKOUT, timestamp: ATTRIBUTE_BREAKOUT)
    end
    files.each do |bytes|
      doc = REXML::Document.new(read_part(bytes, "xl/threadedComments/threadedComment1.xml"))
      comment = REXML::XPath.first(doc, "//*[local-name()='threadedComment']")

      assert_equal ATTRIBUTE_BREAKOUT, comment.attributes["dT"]
      assert_include comment.attributes["id"].downcase, "<injected/>"
      assert_include comment.attributes["parentId"].downcase, "<injected/>"
      assert_nil comment.attributes["injected"]
      assert_empty REXML::XPath.match(doc, "//*[local-name()='injected']")
    end
  end

  private

  # Writes one sheet with Xlsxrb.build and with the streaming writer; returns both files.
  def write_both_modes(&)
    built = StringIO.new
    Xlsxrb.write(built, Xlsxrb.build { |b| b.sheet("S", &) })
    streamed = StringIO.new
    Xlsxrb.write(streamed) { |w| w.sheet("S", &) }
    [built.string, streamed.string]
  end

  def entry_names(bytes)
    Xlsxrb::Ooxml::ZipReader.open(bytes, &:entry_names)
  end

  def read_part(bytes, name)
    Xlsxrb::Ooxml::ZipReader.open(bytes) { |zip| zip.read_entry(name) }
  end
end
