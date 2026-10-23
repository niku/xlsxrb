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
      Xlsxrb.write(f.path, escape_formulas: true) do |w|
        w.sheet("Sheet1") do |s|
          formula_values.each do |val|
            s.row([val])
          end
        end
      end

      # In-memory and streaming read must correctly preserve the exact verbatim string
      # without interpreting or executing command prefixes
      wb = Xlsxrb.read(f.path)
      rows = wb.sheets[0].load.rows
      formula_values.each_with_index do |expected, idx|
        assert_equal expected, rows[idx].cells[0].value
      end
    end
  end
end
