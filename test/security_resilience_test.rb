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

  def test_hash_cell_ref_and_type_are_escaped
    io = StringIO.new
    writer = Xlsxrb::Ooxml::WorksheetWriter.new(io)
    writer.write_row(0, [{ value: 1, column_index: 0, ref: ATTRIBUTE_BREAKOUT }])
    writer.write_row(1, [{ value: 1, column_index: 0, type: ATTRIBUTE_BREAKOUT }])
    writer.finish
    doc = REXML::Document.new(io.string)
    cells = REXML::XPath.match(doc, "//*[local-name()='c']")

    assert_equal 2, cells.size
    assert_equal ATTRIBUTE_BREAKOUT, cells[0].attributes["r"]
    assert_equal ATTRIBUTE_BREAKOUT, cells[1].attributes["t"]
    assert_empty REXML::XPath.match(doc, "//*[local-name()='injected'] | //@injected")
  end

  def test_style_ids_from_a_style_map_are_written_only_as_integers
    style_map = { "bold" => ATTRIBUTE_BREAKOUT, "__xlsxrb_time" => ATTRIBUTE_BREAKOUT, "__xlsxrb_date" => ATTRIBUTE_BREAKOUT }
    time = Time.utc(2026, 1, 2, 3, 4, 5)
    date = Date.new(2026, 1, 2)
    io = StringIO.new
    writer = Xlsxrb::Ooxml::WorksheetWriter.new(io)
    writer.write_row_values(0, [time, date, 1], style_map: style_map)
    writer.write_row_values(1, [time, date, 1, "x", nil], styles: "bold", style_map: style_map)
    writer.write_row_values(2, [time, date, 1], styles: [nil, nil, "bold"], style_map: style_map, attrs: { style_index: "bold" })
    writer.write_row(3, [{ value: 1, column_index: 0, style_index: ATTRIBUTE_BREAKOUT }], attrs: { style_index: ATTRIBUTE_BREAKOUT })
    writer.finish
    doc = REXML::Document.new(io.string)

    assert_equal 4, REXML::XPath.match(doc, "//*[local-name()='row']").size
    assert_empty REXML::XPath.match(doc, "//@s")
    assert_empty REXML::XPath.match(doc, "//*[local-name()='injected'] | //@injected")
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

  def test_option_values_cannot_add_attributes_or_elements
    attack = ATTRIBUTE_BREAKOUT
    font = { size: attack, underline: attack, vert_align: attack, family: attack, scheme: attack }
    writers = {
      "streamed" => ->(io, &block) { Xlsxrb.write(io, &block) },
      "built" => ->(io, &block) { Xlsxrb.write(io, Xlsxrb.build(&block)) }
    }
    writers.each do |label, write|
      io = StringIO.new
      write.call(io) do |w|
        w.custom_property("Count", attack, type: :number)
        w.custom_property("Ratio", attack, type: :float)
        w.custom_property("When", attack, type: :date)
        w.sheet("S") do |s|
          s.row(%w[Region Sales])
          s.row(["East", 1])
          s.row([Xlsxrb.rich_text({ text: "a", font: font }, { text: "b", font: { color: { theme: attack } } }, { text: "c", font: { color: { indexed: attack } } })])
          s.chart(type: :pie, from_column: 3, from_row: 0, to_column: 8, to_row: 12,
                  series: [{ cat_ref: "S!$A$2:$A$2", val_ref: "S!$B$2:$B$2", data_points: [{ idx: attack, fill_color: "4F81BD" }] }])
          s.pivot_table(attack, field_names: %w[Region Sales], row_fields: ["Region"],
                                data_fields: [{ field: "Sales", name: "Sum", subtotal: attack }], ref: attack, name: "P")
        end
      end
      xml_parts = entry_names(io.string).grep(/\.(xml|rels|vml)\z/)

      assert_include xml_parts, "docProps/custom.xml", label
      assert_include xml_parts, "xl/charts/chart1.xml", label
      assert_include xml_parts, "xl/pivotTables/pivotTable1.xml", label
      xml_parts.each do |name|
        doc = REXML::Document.new(read_part(io.string, name))

        assert_empty REXML::XPath.match(doc, "//*[local-name()='injected'] | //@injected"), "#{label}: #{name}"
      end
    end
  end

  def test_values_written_as_element_names_are_checked
    series = { cat_ref: "S!$A$1:$B$1", val_ref: "S!$A$2:$B$2" }
    injected = "barChart><injected/></c:barChart><c:barChart"

    assert_raise_with_message(ArgumentError, /Unknown chart series type/) do
      write_both_modes { |sheet| sheet.chart(type: :bar, series: [series.merge(type: injected)]) }
    end
    # A chart type that is not known is written as a bar chart, as before.
    write_both_modes { |sheet| sheet.chart(type: injected, series: [series]) }.each do |bytes|
      chart = REXML::Document.new(read_part(bytes, "xl/charts/chart1.xml"))
      assert_empty REXML::XPath.match(chart, "//injected")
      assert_equal 1, REXML::XPath.match(chart, "//c:barChart").size
    end
    # Symbols and element names of known types are both accepted for a series.
    write_both_modes { |sheet| sheet.chart(type: :combo, series: [series.merge(type: :line), series.merge(type: "barChart")]) }.each do |bytes|
      chart = REXML::Document.new(read_part(bytes, "xl/charts/chart1.xml"))
      assert_equal([1, 1], ["//c:lineChart", "//c:barChart"].map { |path| REXML::XPath.match(chart, path).size })
    end
  end

  def test_background_image_string_is_never_read_as_a_path
    Tempfile.create(["secret", ".txt"]) do |secret|
      secret.write("DB_PASSWORD=hunter2")
      secret.close

      write_both_modes { |sheet| sheet.background_image = secret.path }.each do |bytes|
        media = entry_names(bytes).grep(%r{\Axl/media/})
        assert_equal ["xl/media/sheet1_bg.png"], media
        assert_equal secret.path.b, read_part(bytes, media.first).b
      end
    end
  end

  def test_shape_color_transforms_are_checked
    shape = ->(transforms) { write_both_modes { |sheet| sheet.shape(preset: "rect", fill_color: "FF0000", fill_color_transforms: transforms) } }

    assert_raise_with_message(ArgumentError, /Unknown color transform/) do
      shape.call([{ type: "tint/><injected/><a:tint", val: 1 }])
    end
    shape.call([{ type: "lumMod", val: ATTRIBUTE_BREAKOUT }]).each do |bytes|
      drawing = REXML::Document.new(read_part(bytes, "xl/drawings/drawing1.xml"))
      assert_empty REXML::XPath.match(drawing, "//injected")
      assert_equal([ATTRIBUTE_BREAKOUT], REXML::XPath.match(drawing, "//a:lumMod").map { |element| element.attributes["val"] })
    end
  end

  def test_element_content_values_are_escaped
    payload = "1</v><injected/>"
    sym_payload = :"1</v><injected/>"

    files = write_both_modes do |sheet|
      sheet.style(:styled, bold: true)
      sheet.row([Xlsxrb.formula("1", cached_value: payload), { formula: "1", value: payload }])
      sheet.row([Xlsxrb.formula("1", cached_value: payload), { formula: "1", value: payload }], cell_styles: :styled)
    end
    files.each do |bytes|
      sheet_doc = REXML::Document.new(read_part(bytes, "xl/worksheets/sheet1.xml"))
      assert_empty REXML::XPath.match(sheet_doc, "//*[local-name()='injected']")
      assert_equal Array.new(4, payload), REXML::XPath.match(sheet_doc, "//*[local-name()='v']").map(&:text)
    end

    io = StringIO.new
    ws = Xlsxrb::Ooxml::WorksheetWriter.new(io)
    ws.write_row(0, [
                   { ref: "A1", value: payload, type: "str" },
                   { ref: "B1", value: payload, type: "e" }
                 ])
    ws.write_row_values(1, [{ formula: "1", value: sym_payload }])
    ws.write_row_values(2, [{ formula: "1", value: sym_payload }], styles: 0)
    ws.finish
    ws_doc = REXML::Document.new(io.string)
    assert_empty REXML::XPath.match(ws_doc, "//*[local-name()='injected']")
    assert_equal Array.new(4, payload), REXML::XPath.match(ws_doc, "//*[local-name()='v']").map(&:text)

    writer = Xlsxrb::Ooxml::Writer.new
    writer.set_cell("A1", 1)
    writer.add_comment("A1", "note", anchor: [1, "0</x:Anchor><injected/>"])
    writer.app_property(:heading_pairs, [["Worksheets", "1</vt:i4><injected/>"]])
    writer.app_property(:titles_of_parts, ["Sheet1"])
    Tempfile.create(["element_content_escape", ".xlsx"]) do |f|
      writer.write(f.path)
      bytes = File.binread(f.path)

      vml_doc = REXML::Document.new(read_part(bytes, "xl/drawings/vmlDrawing1.vml"))
      assert_empty REXML::XPath.match(vml_doc, "//*[local-name()='injected']")
      assert_equal "1, 0</x:Anchor><injected/>", REXML::XPath.first(vml_doc, "//*[local-name()='Anchor']").text

      app_doc = REXML::Document.new(read_part(bytes, "docProps/app.xml"))
      assert_empty REXML::XPath.match(app_doc, "//*[local-name()='injected']")
      assert_equal "1</vt:i4><injected/>", REXML::XPath.first(app_doc, "//*[local-name()='i4']").text
    end
  end

  def test_hot_path_attribute_and_shared_strings_breakout
    omit "Passes values the signatures do not allow, which RBS runtime checks reject first" if defined?(RBS::Test)

    rt = Xlsxrb::Elements::RichText.new(
      runs: [
        {
          text: ATTRIBUTE_BREAKOUT,
          font: {
            underline: ATTRIBUTE_BREAKOUT,
            vert_align: ATTRIBUTE_BREAKOUT,
            size: ATTRIBUTE_BREAKOUT,
            color: ATTRIBUTE_BREAKOUT,
            name: ATTRIBUTE_BREAKOUT,
            family: ATTRIBUTE_BREAKOUT,
            scheme: ATTRIBUTE_BREAKOUT
          }
        }
      ]
    )
    shared_obj = Xlsxrb.shared_formula("1", shared_index: ATTRIBUTE_BREAKOUT, ref: ATTRIBUTE_BREAKOUT)
    row_attrs = { height: ATTRIBUTE_BREAKOUT, outline_level: ATTRIBUTE_BREAKOUT }

    io = StringIO.new
    ws = Xlsxrb::Ooxml::WorksheetWriter.new(io)
    ws.write_row(0, [{ column_index: 0, formula: shared_obj }, { column_index: 1, value: rt }], attrs: row_attrs)
    ws.write_row_values(1, [shared_obj, rt], sst: nil, sst_index: nil, attrs: row_attrs)
    ws.write_row_values(2, [shared_obj, rt], styles: [0, 0], sst: nil, sst_index: nil, attrs: row_attrs)
    ws.finish

    ws_doc = REXML::Document.new(io.string)
    assert_empty REXML::XPath.match(ws_doc, "//injected")
    rows = REXML::XPath.match(ws_doc, "//row")
    assert_equal(Array.new(3, ATTRIBUTE_BREAKOUT), rows.map { |r| r.attributes["ht"] })
    assert_equal(Array.new(3, ATTRIBUTE_BREAKOUT), rows.map { |r| r.attributes["outlineLevel"] })
    formulas = REXML::XPath.match(ws_doc, "//f")
    assert_equal(Array.new(3, ATTRIBUTE_BREAKOUT), formulas.map { |f| f.attributes["si"] })
    assert_equal(Array.new(3, ATTRIBUTE_BREAKOUT), formulas.map { |f| f.attributes["ref"] })
    assert_equal(Array.new(3, ATTRIBUTE_BREAKOUT), REXML::XPath.match(ws_doc, "//u").map { |u| u.attributes["val"] })
    assert_equal(Array.new(3, ATTRIBUTE_BREAKOUT), REXML::XPath.match(ws_doc, "//sz").map { |s| s.attributes["val"] })
    assert_equal(Array.new(3, ATTRIBUTE_BREAKOUT), REXML::XPath.match(ws_doc, "//t").map(&:text))

    streamed = StringIO.new
    Xlsxrb.write(streamed) do |wb|
      wb.sheet("S") { |sheet| sheet.row([ATTRIBUTE_BREAKOUT, rt]) }
    end
    sst_doc = REXML::Document.new(read_part(streamed.string, "xl/sharedStrings.xml"))
    assert_empty REXML::XPath.match(sst_doc, "//injected")
    assert_equal [ATTRIBUTE_BREAKOUT, ATTRIBUTE_BREAKOUT], REXML::XPath.match(sst_doc, "//t").map(&:text)
    assert_equal ATTRIBUTE_BREAKOUT, REXML::XPath.first(sst_doc, "//u").attributes["val"]
    assert_equal ATTRIBUTE_BREAKOUT, REXML::XPath.first(sst_doc, "//vertAlign").attributes["val"]
    assert_equal ATTRIBUTE_BREAKOUT, REXML::XPath.first(sst_doc, "//sz").attributes["val"]
    assert_equal ATTRIBUTE_BREAKOUT, REXML::XPath.first(sst_doc, "//rFont").attributes["val"]
    assert_equal ATTRIBUTE_BREAKOUT, REXML::XPath.first(sst_doc, "//family").attributes["val"]
    assert_equal ATTRIBUTE_BREAKOUT, REXML::XPath.first(sst_doc, "//scheme").attributes["val"]
  end

  def test_xml_builder_attributes_and_text_are_escaped_once
    writer = Xlsxrb::Ooxml::Writer.new

    persons_xml = writer.generate_persons_xml(
      [
        {
          id: ATTRIBUTE_BREAKOUT,
          display_name: ATTRIBUTE_BREAKOUT,
          user_id: ATTRIBUTE_BREAKOUT,
          provider_id: ATTRIBUTE_BREAKOUT
        }
      ]
    )
    refute_includes persons_xml, "&amp;quot;"
    refute_includes persons_xml, "&amp;lt;"
    persons_doc = REXML::Document.new(persons_xml)
    assert_empty REXML::XPath.match(persons_doc, "//*[local-name()='injected']")
    person = REXML::XPath.first(persons_doc, "//*[local-name()='person']")
    assert_equal ATTRIBUTE_BREAKOUT, person.attributes["displayName"]
    assert_equal ATTRIBUTE_BREAKOUT, person.attributes["id"]
    assert_equal ATTRIBUTE_BREAKOUT, person.attributes["userId"]
    assert_equal ATTRIBUTE_BREAKOUT, person.attributes["providerId"]

    drawing_xml = writer.generate_drawing_xml(
      [
        {
          kind: :chart,
          rid_index: 1,
          chart: {
            from_column: 0,
            from_row: 0,
            to_column: 5,
            to_row: 10,
            frame_macro: ATTRIBUTE_BREAKOUT
          }
        }
      ]
    )
    refute_includes drawing_xml, "&amp;quot;"
    refute_includes drawing_xml, "&amp;lt;"
    drawing_doc = REXML::Document.new(drawing_xml)
    assert_empty REXML::XPath.match(drawing_doc, "//*[local-name()='injected']")
    assert_equal ATTRIBUTE_BREAKOUT, REXML::XPath.first(drawing_doc, "//*[local-name()='graphicFrame']").attributes["macro"]

    chart_xml = writer.generate_chart_xml(
      {
        type: :bar,
        title: { text: ATTRIBUTE_BREAKOUT },
        series: [
          {
            name: "S1",
            fill_color: ATTRIBUTE_BREAKOUT,
            line_color: { scheme: ATTRIBUTE_BREAKOUT }
          }
        ]
      }
    )
    refute_includes chart_xml, "&amp;quot;"
    refute_includes chart_xml, "&amp;lt;"
    chart_doc = REXML::Document.new(chart_xml)
    assert_empty REXML::XPath.match(chart_doc, "//*[local-name()='injected']")
    assert_equal ATTRIBUTE_BREAKOUT, REXML::XPath.first(chart_doc, "//*[local-name()='t']").text
    assert_equal ATTRIBUTE_BREAKOUT, REXML::XPath.first(chart_doc, "//*[local-name()='srgbClr']").attributes["val"]
    assert_equal ATTRIBUTE_BREAKOUT, REXML::XPath.first(chart_doc, "//*[local-name()='schemeClr']").attributes["val"]
  end

  HOT_PATH_STRING_WRITERS = [
    "lib/xlsxrb/ooxml/worksheet_writer.rb",
    "lib/xlsxrb/ooxml/workbook_writer.rb"
  ].freeze
  RAW_XML_START_TAG = %r{<[a-zA-Z_][\w:.-]*[\s>/]}

  # Values that the string-concatenating hot-path writers put in an attribute without
  # escaping them, by file, each with the reason it cannot hold a character that needs
  # escaping. All other XML writers build XML through XmlBuilder and may not construct raw
  # attributes at all.
  UNESCAPED_ATTRIBUTE_VALUES = {
    "lib/xlsxrb/ooxml/worksheet_writer.rb" => {
      "col_ref" => "column letters and a row number computed here, or the ref of a Hash cell, escaped where it is taken",
      "row_num_str" => "Integer#to_s",
      "type" => "a literal, or the type of a Hash cell, escaped where it is taken",
      "attrs[:style_index]" => "written only if it is an Integer",
      "style_id" => "an Integer: resolve_style_id returns nothing else",
      "cell_style_id" => "an Integer: resolve_style_id returns nothing else",
      "row_style_id" => "written only if it is an Integer",
      "time_style_id" => "written only if it is an Integer",
      "date_style_id" => "written only if it is an Integer"
    },
    "lib/xlsxrb/ooxml/workbook_writer.rb" => {
      "count_str" => "Integer#to_s"
    }
  }.freeze

  ESCAPE_CALL = /\A(?:xml_escape|xml_escape_text|escape_xml|escape|XmlBuilder\.escape|Xlsxrb::Ooxml::XmlBuilder\.escape)\(.*\)\z/
  # Forms that cannot produce a character that needs escaping: an Integer, a constant,
  # a choice between two literals, and Base64.
  SAFE_ATTRIBUTE_VALUES = [
    /\.to_i\z/, /\AInteger\(.*\)\z/, /\.(?:size|length|count|bytesize)\z/, /\A\d+\z/,
    /\A[A-Z][A-Z0-9_]*\z/,
    /\?\s*(?:"[^"#]*"|\d+|:\w+)\s*:\s*(?:"[^"#]*"|\d+|:\w+)\z/,
    /\.pack\("m0"\)\z/
  ].freeze
  # An interpolation in an attribute value: ="...#{ , and a concatenation: ="' << value
  ATTRIBUTE_INTERPOLATION = /=\\?["'](?:[^"'#]|#(?!\{))*#\{/
  ATTRIBUTE_CONCATENATION = /="' *<< *([^<\n]+?) *(?=<<|$)/

  # Row and shared-string hot paths still build XML strings for throughput, so this reads
  # every file under lib/ and fails if any raw attribute construction appears outside
  # HOT_PATH_STRING_WRITERS, or if an attribute value in HOT_PATH_STRING_WRITERS is
  # neither escaped, nor of a safe form, nor listed in UNESCAPED_ATTRIBUTE_VALUES.
  #
  # Static inspection does not extend to element content (`>#{...}</`) in
  # string-interpolating writers because most element-content interpolations insert
  # pre-built child XML fragments that must not be escaped; a regex on source lines cannot
  # distinguish an unescaped text variable from a variable holding child XML elements
  # without maintaining an exhaustive list of fragment expressions.
  def test_writers_escape_every_attribute_value
    root = File.expand_path("..", __dir__)
    seen = 0
    unused = UNESCAPED_ATTRIBUTE_VALUES.to_h { |file, values| [file, values.keys] }
    offenses = []
    Dir[File.join(root, "lib/**/*.rb")].each do |path|
      file = path.delete_prefix("#{root}/")
      listed = UNESCAPED_ATTRIBUTE_VALUES.fetch(file, {})
      hot_path = HOT_PATH_STRING_WRITERS.include?(file)
      File.foreach(path, encoding: Encoding::UTF_8).with_index(1) do |line, number|
        attribute_values(line).each do |expression|
          unless hot_path
            offenses << "#{file}:#{number}: raw attribute outside hot-path writers: #{expression}"
            next
          end

          seen += 1
          value = expression.delete_suffix(".to_s")
          next if value.match?(ESCAPE_CALL) || SAFE_ATTRIBUTE_VALUES.any? { |form| value.match?(form) }

          if listed.key?(value)
            unused[file].delete(value)
          else
            offenses << "#{file}:#{number}: #{expression}"
          end
        end
      end
    end

    assert_operator seen, :>, 80, "the patterns no longer match the attributes the hot-path writers write"
    assert_empty offenses, "written into an attribute as given; build XML with XmlBuilder, or escape each hot-path value or list it in UNESCAPED_ATTRIBUTE_VALUES"
    assert_empty unused.reject { |_file, values| values.empty? }, "listed in UNESCAPED_ATTRIBUTE_VALUES but not found"
  end

  def test_xml_builder_writers_emit_no_raw_start_tags
    root = File.expand_path("..", __dir__)
    files = %w[lib/xlsxrb/ooxml/writer.rb lib/xlsxrb/ooxml/crypto/agile.rb] +
            Dir[File.join(root, "lib/xlsxrb/ooxml/writer/*.rb")].map { |p| p.delete_prefix("#{root}/") }
    offenses = []
    files.each do |file|
      File.foreach(File.join(root, file), encoding: Encoding::UTF_8).with_index(1) do |line, number|
        next if line.lstrip.start_with?("#")

        offenses << "#{file}:#{number}: #{line.strip}" if line.match?(RAW_XML_START_TAG)
      end
    end

    assert_empty offenses, "XML writers migrated to XmlBuilder must not construct raw XML start tags"
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

  # The expressions a line of source writes into attribute values.
  def attribute_values(line)
    values = line.scan(ATTRIBUTE_CONCATENATION).flatten
    position = 0
    while (match = line.match(ATTRIBUTE_INTERPOLATION, position))
      first = match.end(0)
      depth = 1
      last = first
      while last < line.size && depth.positive?
        depth += 1 if line[last] == "{"
        depth -= 1 if line[last] == "}"
        last += 1
      end
      values << line[first...(last - 1)].strip if depth.zero?
      position = last
    end
    values
  end
end
