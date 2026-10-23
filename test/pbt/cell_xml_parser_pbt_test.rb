# frozen_string_literal: true

require "test_helper"
require "pbt"
require "rexml/document"
require "tempfile"

class CellXmlParserPbtTest < Test::Unit::TestCase
  def setup
    omit "Skip property-based tests under RBS runtime testing to prevent OOM / slowdown" if defined?(RBS::Test)
  end

  # Generates varied cell configurations:
  # - empty self-closing cell (<c r="A1" s="1"/>)
  # - string with phonetic furigana (<c r="B1" t="inlineStr"><is><t>Hello</t><rPh><t>ふりがな</t></rPh></is></c>)
  # - rich text runs with formatting tags (<c r="C1" t="inlineStr"><is><r><rPr><b/><sz val="11"/></rPr><t>Part1</t></r><r><t>Part2</t></r></is></c>)
  # - numeric cell (integer or float)
  # - boolean cell (t="b")
  # - attributes with arbitrary ordering (r, s, t)
  # - optional namespace prefixes (<x:c>, <x:is>, <x:t>)
  def arbitrary_cell_xml_row
    cell_gen = Pbt.tuple(
      Pbt.choose(0..5), # cell variant
      Pbt.integer(min: 0, max: 1000), # int value
      Pbt.boolean, # boolean or namespace toggle
      Pbt.alphanumeric_string(min: 1, max: 10), # text1
      Pbt.alphanumeric_string(min: 1, max: 10), # text2 / furigana
      Pbt.choose(0..5) # attribute permutation (6 permutations of r, s, t)
    )

    Pbt.array(cell_gen, min: 1, max: 6)
  end

  def build_row_xml(cell_specs)
    cells_xml = +""
    col_names = %w[A B C D E F]

    cell_specs.each_with_index do |spec, idx|
      variant, int_val, ns_toggle, text1, text2, perm = spec
      col_letter = col_names[idx] || "A"
      ref = "#{col_letter}1"
      p = ns_toggle ? "x:" : ""

      attrs = [
        %(r="#{ref}"),
        %(s="1")
      ]

      case variant
      when 0 # Self-closing empty cell
        cells_xml << "<#{p}c #{attrs.join(" ")}/>"
      when 1 # Simple numeric cell
        cells_xml << "<#{p}c #{attrs.join(" ")}><#{p}v>#{int_val}</#{p}v></#{p}c>"
      when 2 # Boolean cell
        attrs << %(t="b")
        attrs.shuffle!
        val_str = int_val.even? ? "1" : "0"
        cells_xml << "<#{p}c #{attrs.join(" ")}><#{p}v>#{val_str}</#{p}v></#{p}c>"
      when 3 # Inline text with phonetic reading (<rPh>)
        attrs << %(t="inlineStr")
        # Permute attributes based on perm
        attrs.rotate!(perm % attrs.size)
        cells_xml << "<#{p}c #{attrs.join(" ")}><#{p}is><#{p}t>#{text1}</#{p}t><#{p}rPh sb=\"0\" eb=\"2\"><#{p}t>#{text2}</#{p}t></#{p}rPh></#{p}is></#{p}c>"
      when 4 # Rich text with formatting tags (<rPr>, <b/>)
        attrs << %(t="inlineStr")
        cells_xml << "<#{p}c #{attrs.join(" ")}><#{p}is><#{p}r><#{p}rPr><b/><sz val=\"11\"/></#{p}rPr><#{p}t>#{text1}</#{p}t></#{p}r><#{p}r><#{p}t>#{text2}</#{p}t></#{p}r></#{p}is></#{p}c>"
      when 5 # Self-closing empty cell with t="s"
        attrs << %(t="s")
        cells_xml << "<#{p}c #{attrs.join(" ")}/>"
      end
    end

    <<~XML
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
          <row r="1">#{cells_xml}</row>
        </sheetData>
      </worksheet>
    XML
  end

  test "property: fast scan parser and listener parser yield identical cell values on generated XML" do
    Pbt.assert(num_runs: 100) do
      Pbt.property(arbitrary_cell_xml_row) do |cell_specs|
        xml = build_row_xml(cell_specs)

        # 1. Parse via fast streaming scan
        fast_sheet = Xlsxrb::StreamSheet.new("Sheet1", xml.b, [])
        fast_row = fast_sheet.each_row.first
        fast_values = fast_row ? fast_row.values : []

        # 2. Parse via core listener / SAX parser
        listener = Xlsxrb::Ooxml::Reader::WorksheetListener.new([])
        parser = REXML::Parsers::SAX2Parser.new(xml)
        parser.listen(listener)
        parser.parse
        listener_cells = listener.cells
        listener_values = (0...cell_specs.size).map do |col_idx|
          ref = "#{%w[A B C D E F][col_idx]}1"
          listener_cells[ref]
        end

        # Trim trailing nils to align arrays
        fast_values.pop while fast_values.last.nil? && fast_values.size > listener_values.size
        listener_values.pop while listener_values.last.nil? && listener_values.size > fast_values.size

        # Compare plain values (converting RichText to string for uniform comparison)
        normalized_listener = listener_values.map { |v| v.is_a?(Xlsxrb::Elements::RichText) ? v.to_s : v }
        normalized_fast = fast_values.map { |v| v.is_a?(Xlsxrb::Elements::RichText) ? v.to_s : v }

        assert_equal normalized_listener, normalized_fast, "Mismatch between fast parser and listener parser on XML:\n#{xml}"
      end
    end
  end
end
