# frozen_string_literal: true

require "test_helper"
require "pbt"
require "stringio"
require "tempfile"
require "zlib"
require_relative "../support/excel_rules"
require_relative "../support/xsd_validator"

class ParserFormulaZipPbtTest < Test::Unit::TestCase
  def setup
    omit "Skip property-based tests under RBS runtime testing to prevent OOM / slowdown" if defined?(RBS::Test)
  end

  # --- Item 9a: Xlsxrb.expand_formula AST-based property test ---

  SHEET_PREFIXES = [
    nil,
    "Sheet2!",
    "S1!",
    "Q1!",
    "'Q1'!",
    "'Sales 2026'!",
    "'O''Reilly'!"
  ].freeze

  LITERAL_TOKENS = [
    '"A1+$B$2"',
    '"say ""C3"" ok"',
    '"S1!B10"',
    '"SUM(D4:E5)"'
  ].freeze

  FUNC_NAMES = %w[SUM LOG10 ATAN2].freeze

  def col_index_to_letter(col_idx)
    Xlsxrb::Elements::Cell._column_letter(col_idx)
  end

  def render_ref_node(node)
    sheet = node[:sheet] || ""
    col_part = "#{node[:col_abs]}#{col_index_to_letter(node[:col])}"
    row_part = "#{node[:row_abs]}#{node[:row]}"
    "#{sheet}#{col_part}#{row_part}"
  end

  def render_formula_ast(ast)
    case ast[:kind]
    when :cell_ref
      render_ref_node(ast)
    when :range_ref
      sheet = ast[:sheet] || ""
      left = "#{ast[:c1_abs]}#{col_index_to_letter(ast[:c1])}#{ast[:r1_abs]}#{ast[:r1]}"
      right = "#{ast[:c2_abs]}#{col_index_to_letter(ast[:c2])}#{ast[:r2_abs]}#{ast[:r2]}"
      "#{sheet}#{left}:#{right}"
    when :str_lit
      ast[:text]
    when :func_call
      rendered_args = ast[:args].map { |arg| render_formula_ast(arg) }.join(",")
      "#{ast[:name]}(#{rendered_args})"
    when :bin_op
      "#{render_formula_ast(ast[:left])}#{ast[:op]}#{render_formula_ast(ast[:right])}"
    end
  end

  def shift_formula_ast(ast, row_delta, col_delta)
    case ast[:kind]
    when :cell_ref
      new_col = ast[:col_abs].empty? ? (ast[:col] + col_delta).clamp(0, 16_383) : ast[:col]
      new_row = ast[:row_abs].empty? ? (ast[:row] + row_delta).clamp(1, 1_048_576) : ast[:row]
      ast.merge(col: new_col, row: new_row)
    when :range_ref
      new_c1 = ast[:c1_abs].empty? ? (ast[:c1] + col_delta).clamp(0, 16_383) : ast[:c1]
      new_r1 = ast[:r1_abs].empty? ? (ast[:r1] + row_delta).clamp(1, 1_048_576) : ast[:r1]
      new_c2 = ast[:c2_abs].empty? ? (ast[:c2] + col_delta).clamp(0, 16_383) : ast[:c2]
      new_r2 = ast[:r2_abs].empty? ? (ast[:r2] + row_delta).clamp(1, 1_048_576) : ast[:r2]
      ast.merge(c1: new_c1, r1: new_r1, c2: new_c2, r2: new_r2)
    when :str_lit
      ast
    when :func_call
      ast.merge(args: ast[:args].map { |arg| shift_formula_ast(arg, row_delta, col_delta) })
    when :bin_op
      ast.merge(
        left: shift_formula_ast(ast[:left], row_delta, col_delta),
        right: shift_formula_ast(ast[:right], row_delta, col_delta)
      )
    end
  end

  def build_formula_ast(spec)
    sheet_a, col_a, row_a, cabs_a, rabs_a,
      sheet_b, col_b, row_b, cabs_b, rabs_b,
      func_name, op, lit, shape_idx = spec

    ref_a = { kind: :cell_ref, sheet: sheet_a, col_abs: cabs_a, col: col_a, row_abs: rabs_a, row: row_a }
    ref_b = { kind: :cell_ref, sheet: sheet_b, col_abs: cabs_b, col: col_b, row_abs: rabs_b, row: row_b }
    rng = {
      kind: :range_ref,
      sheet: sheet_a,
      c1_abs: cabs_a,
      c1: col_a,
      r1_abs: rabs_a,
      r1: row_a,
      c2_abs: cabs_b,
      c2: col_a + 2,
      r2_abs: rabs_b,
      r2: row_a + 3
    }
    lit_node = { kind: :str_lit, text: lit }

    case shape_idx % 4
    when 0
      { kind: :bin_op, op: op, left: { kind: :func_call, name: func_name, args: [ref_a, ref_b] }, right: lit_node }
    when 1
      { kind: :bin_op, op: "+", left: { kind: :func_call, name: func_name, args: [rng] }, right: ref_b }
    when 2
      { kind: :bin_op, op: "&", left: lit_node, right: { kind: :bin_op, op: "+", left: ref_a, right: ref_b } }
    else
      { kind: :func_call, name: "IF", args: [ref_a, rng, lit_node] }
    end
  end

  test "property: Xlsxrb.expand_formula algebraic invariants, anchoring, literals, and sheet-qualified references" do
    ref_parts = Pbt.tuple(
      Pbt.constant_from(*SHEET_PREFIXES),
      Pbt.integer(min: 10, max: 35),
      Pbt.integer(min: 10, max: 40),
      Pbt.constant_from("", "$"),
      Pbt.constant_from("", "$")
    )
    formula_gen = Pbt.tuple(
      ref_parts,
      ref_parts,
      Pbt.constant_from(*FUNC_NAMES),
      Pbt.constant_from("+", "-", "*", "&"),
      Pbt.constant_from(*LITERAL_TOKENS),
      Pbt.integer(min: 0, max: 11),
      Pbt.integer(min: -5, max: 5),
      Pbt.integer(min: -5, max: 5),
      Pbt.integer(min: -3, max: 3),
      Pbt.integer(min: -3, max: 3)
    )

    Pbt.assert(num_runs: 150) do
      Pbt.property(formula_gen) do |parts_a, parts_b, func_name, op, lit, shape_idx, dr1, dc1, dr2, dc2|
        ast = build_formula_ast([*parts_a, *parts_b, func_name, op, lit, shape_idx])
        formula = render_formula_ast(ast)

        assert_equal(formula, Xlsxrb.expand_formula(formula, row_delta: 0, column_delta: 0))
        assert_equal(formula, Xlsxrb.expand_formula(formula, from_ref: "K10", to_ref: "K10"))

        expected_one = render_formula_ast(shift_formula_ast(ast, dr1, dc1))
        by_delta = Xlsxrb.expand_formula(formula, row_delta: dr1, column_delta: dc1)
        assert_equal(expected_one, by_delta, "Oracle mismatch for formula=#{formula.inspect} dr=#{dr1} dc=#{dc1}")

        from_ref = "#{col_index_to_letter(10)}20"
        to_ref = "#{col_index_to_letter(10 + dc1)}#{20 + dr1}"
        by_refs = Xlsxrb.expand_formula(formula, from_ref: from_ref, to_ref: to_ref)
        assert_equal(expected_one, by_refs, "from_ref/to_ref mismatch for formula=#{formula.inspect}")

        restored = Xlsxrb.expand_formula(by_delta, row_delta: -dr1, column_delta: -dc1)
        assert_equal(formula, restored, "Invertibility failed for formula=#{formula.inspect}")

        step_two = Xlsxrb.expand_formula(by_delta, row_delta: dr2, column_delta: dc2)
        combined = Xlsxrb.expand_formula(formula, row_delta: dr1 + dr2, column_delta: dc1 + dc2)
        assert_equal(combined, step_two, "Additivity failed for formula=#{formula.inspect}")

        min_expected = render_formula_ast(shift_formula_ast(ast, -100, -100))
        assert_equal(min_expected, Xlsxrb.expand_formula(formula, row_delta: -100, column_delta: -100))

        max_expected = render_formula_ast(shift_formula_ast(ast, 2_000_000, 20_000))
        assert_equal(max_expected, Xlsxrb.expand_formula(formula, row_delta: 2_000_000, column_delta: 20_000))
      end
    end
  end

  # --- Item 9b: Shared and array formulas across writers, events, listeners, and roundtrip ---

  test "property: shared and array formulas across writers, events, listeners, and roundtrip" do
    gen = Pbt.tuple(
      Pbt.integer(min: 2, max: 4),
      Pbt.constant_from("A1*2", "$A$1+B1", "A1+$B1"),
      Pbt.constant_from("SUM(A1:A2*C1:C2)", "A1:A2+1"),
      Pbt.boolean
    )

    Pbt.assert(num_runs: 25) do
      Pbt.property(gen) do |slave_count, master_expr, array_expr, use_self_closing_slave|
        writer = Xlsxrb::Ooxml::Writer.new
        writer.set_cell("A1", 10)
        writer.set_cell("B1", 100)
        writer.set_cell("C1", 2)
        writer.set_cell(
          "D1",
          Xlsxrb::Elements::Formula.new(
            expression: master_expr,
            type: :shared,
            ref: "D1:D#{slave_count + 1}",
            shared_index: 0,
            cached_value: 20
          )
        )
        slave_count.times do |idx|
          row_num = idx + 2
          writer.set_cell("A#{row_num}", 10 + row_num)
          writer.set_cell("B#{row_num}", 100 + row_num)
          writer.set_cell("C#{row_num}", 2 + row_num)
          writer.set_cell(
            "D#{row_num}",
            Xlsxrb::Elements::Formula.new(
              expression: "",
              type: :shared,
              shared_index: 0,
              cached_value: 20 + row_num
            )
          )
        end
        writer.set_cell(
          "E1",
          Xlsxrb::Elements::Formula.new(
            expression: array_expr,
            type: :array,
            ref: "E1",
            cached_value: 99
          )
        )

        xlsx_bytes = Tempfile.create(["pbt_shared_formula", ".xlsx"]) do |tmp|
          tmp.close
          writer.write(tmp.path)
          File.binread(tmp.path)
        end
        assert_empty(ExcelRules.check(xlsx_bytes).map(&:to_s))
        assert_empty(XsdValidator.check(xlsx_bytes))

        raw_sheet_xml = Xlsxrb::Ooxml::ZipReader.new(StringIO.new(xlsx_bytes)).read_entry("xl/worksheets/sheet1.xml")
        raw_sheet_xml = raw_sheet_xml.gsub(%r{<f t="shared" si="0"/>}, '<f t="shared" si="0"></f>') unless use_self_closing_slave

        event_formulas = {}
        Xlsxrb::Ooxml::WorksheetParser.each_event(raw_sheet_xml, shared_strings: []) do |ev|
          event_formulas[ev.args[0]] = ev.args[4] if ev.type == :cell && ev.args[4]
        end

        listener = Xlsxrb::Ooxml::Reader::WorksheetListener.new([])
        Xlsxrb::Ooxml::XmlParser.parse(raw_sheet_xml, listener)
        listener_cells = listener.cells

        expected_master = master_expr
        assert_equal(expected_master, event_formulas["D1"])
        assert_equal(expected_master, listener_cells["D1"].expression)
        assert_equal(array_expr, event_formulas["E1"])
        assert_equal(array_expr, listener_cells["E1"].expression)
        assert_equal(:array, listener_cells["E1"].type)

        slave_count.times do |idx|
          target_ref = "D#{idx + 2}"
          expected_slave = Xlsxrb.expand_formula(master_expr, from_ref: "D1", to_ref: target_ref)
          assert_equal(expected_slave, event_formulas[target_ref], "each_event slave mismatch at #{target_ref}")
          assert_equal(expected_slave, listener_cells[target_ref].expression, "listener slave mismatch at #{target_ref}")
        end

        loaded = Xlsxrb.load(StringIO.new(xlsx_bytes))
        ws = loaded.sheet(0)
        assert_equal(array_expr, ws["E1"].formula_expression)
        assert_equal(:array, ws["E1"].formula.type)

        resaved = loaded.to_data
        assert_empty(ExcelRules.check(resaved).map(&:to_s))
        assert_empty(XsdValidator.check(resaved))
        reloaded = Xlsxrb.load(StringIO.new(resaved)).sheet(0)
        assert_equal(master_expr, reloaded["D1"].formula_expression)
        assert_equal(array_expr, reloaded["E1"].formula_expression)
        slave_count.times do |idx|
          target_ref = "D#{idx + 2}"
          expected_slave = Xlsxrb.expand_formula(master_expr, from_ref: "D1", to_ref: target_ref)
          assert_equal(expected_slave, reloaded[target_ref].formula_expression)
        end
      end
    end
  end

  # --- Item 10: WorksheetParser multi-path equivalence + chunk boundary splitting ---

  # Simulates arbitrarily chunked IO streams across read(len) calls.
  class ChunkedXmlIo
    def initialize(bytes, chunk_size)
      @bytes = bytes.b
      @chunk_size = chunk_size
      @pos = 0
    end

    def read(len = nil)
      return nil if @pos >= @bytes.bytesize

      take = len ? [len, @chunk_size, @bytes.bytesize - @pos].min : (@bytes.bytesize - @pos)
      slice = @bytes.byteslice(@pos, take)
      @pos += take
      slice
    end
  end

  CELL_VARIANTS = %i[
    num_plain
    num_explicit
    num_empty_pair
    num_empty_self
    bool_num
    bool_word
    shared_str
    inline_plain
    inline_entities_crlf
    inline_rich
    inline_rph
    inline_cdata
    inline_empty_tag
    inline_empty_t
    formula_str
    err_cell
    self_closing_empty
  ].freeze

  def build_generated_worksheet(specs, opts)
    prefix = opts[:use_prefix] ? "x:" : ""
    omit_r = opts[:omit_r]
    sst = [
      "Alpha & Beta <one>",
      "日本語テスト🌸",
      "Line1\nLine2"
    ]

    row_xmls = []
    specs.each_with_index do |row_spec, row_idx|
      r_attr = omit_r ? "" : %( r="#{row_idx + 1}")
      extra_row_attrs = if row_idx.zero? && !omit_r
                          ' ht="22.5" customHeight="1" hidden="1" collapsed="1" outlineLevel="2" s="1"'
                        else
                          ""
                        end

      if row_spec[:self_closing]
        row_xmls << "<#{prefix}row#{r_attr}/>"
        next
      end

      cells_buf = +""
      col_cursor = 0
      row_spec[:cells].each_with_index do |cell_spec, idx|
        col_cursor += cell_spec[:gap]
        col_letter = col_index_to_letter(col_cursor)
        ref_attr = omit_r ? nil : %(r="#{col_letter}#{row_idx + 1}")
        style_attr = %(s="#{(idx % 2) + 1}")
        cells_buf << render_generated_cell(prefix, cell_spec, ref_attr, style_attr, sst)
        col_cursor += 1
      end

      row_xmls << "<#{prefix}row#{r_attr}#{extra_row_attrs}>#{cells_buf}</#{prefix}row>"
    end

    decoy = opts[:with_decoy] ? "<#{prefix}rowHeader/><!-- hyperlink decoy -->" : ""
    links_xml = if opts[:with_hyperlinks]
                  <<~HL
                    <#{prefix}hyperlinks>
                      <#{prefix}hyperlink ref="A1" r:id="rId1" display="A &amp; B" tooltip="Tip &lt;1&gt; &quot;hi&quot;"/>
                      <#{prefix}hyperlink ref="B1" location="Sheet2!A1" display="Go 日本語"/>
                    </#{prefix}hyperlinks>
                  HL
                else
                  ""
                end

    xml = <<~XML
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <#{prefix}worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        #{decoy}
        <#{prefix}sheetData>
          #{row_xmls.join("\n    ")}
        </#{prefix}sheetData>
        #{links_xml}
      </#{prefix}worksheet>
    XML
    [xml, sst]
  end

  def render_generated_cell(prefix, spec, ref_attr, style_attr, _sst)
    t_attr = case spec[:variant]
             when :num_explicit, :num_empty_pair, :num_empty_self then 't="n"'
             when :bool_num, :bool_word then 't="b"'
             when :shared_str then 't="s"'
             when :inline_plain, :inline_entities_crlf, :inline_rich, :inline_rph, :inline_cdata, :inline_empty_tag, :inline_empty_t
               't="inlineStr"'
             when :formula_str then 't="str"'
             when :err_cell then 't="e"'
             end

    attrs = [ref_attr, style_attr, t_attr].compact
    attrs.rotate!(spec[:perm] % attrs.size) unless attrs.empty?
    attr_str = attrs.empty? ? "" : " #{attrs.join(" ")}"

    case spec[:variant]
    when :num_plain, :num_explicit
      "<#{prefix}c#{attr_str}><#{prefix}v>#{spec[:num]}</#{prefix}v></#{prefix}c>"
    when :num_empty_pair
      "<#{prefix}c#{attr_str}><#{prefix}v></#{prefix}v></#{prefix}c>"
    when :num_empty_self
      "<#{prefix}c#{attr_str}><#{prefix}v/></#{prefix}c>"
    when :bool_num
      "<#{prefix}c#{attr_str}><#{prefix}v>#{spec[:num].even? ? "1" : "0"}</#{prefix}v></#{prefix}c>"
    when :bool_word
      "<#{prefix}c#{attr_str}><#{prefix}v>#{spec[:num].even? ? "true" : "false"}</#{prefix}v></#{prefix}c>"
    when :shared_str
      "<#{prefix}c#{attr_str}><#{prefix}v>#{spec[:num] % 3}</#{prefix}v></#{prefix}c>"
    when :inline_plain
      "<#{prefix}c#{attr_str}><#{prefix}is><#{prefix}t>Hello_🌸_#{spec[:num]}</#{prefix}t></#{prefix}is></#{prefix}c>"
    when :inline_entities_crlf
      "<#{prefix}c#{attr_str}><#{prefix}is><#{prefix}t>A &amp; B &lt;tag&gt;\r\nL2</#{prefix}t></#{prefix}is></#{prefix}c>"
    when :inline_rich
      "<#{prefix}c#{attr_str}><#{prefix}is><#{prefix}r><#{prefix}rPr><b/></#{prefix}rPr><#{prefix}t>R1</#{prefix}t></#{prefix}r><#{prefix}r><#{prefix}t>R2</#{prefix}t></#{prefix}r></#{prefix}is></#{prefix}c>"
    when :inline_rph
      "<#{prefix}c#{attr_str}><#{prefix}is><#{prefix}t>漢字</#{prefix}t><#{prefix}rPh sb=\"0\" eb=\"2\"><#{prefix}t>かんじ</#{prefix}t></#{prefix}rPh></#{prefix}is></#{prefix}c>"
    when :inline_cdata
      "<#{prefix}c#{attr_str}><#{prefix}is><#{prefix}t><![CDATA[C & <D>]]></#{prefix}t></#{prefix}is></#{prefix}c>"
    when :inline_empty_tag
      "<#{prefix}c#{attr_str}><#{prefix}is/></#{prefix}c>"
    when :inline_empty_t
      "<#{prefix}c#{attr_str}><#{prefix}is><#{prefix}t/></#{prefix}is></#{prefix}c>"
    when :formula_str
      "<#{prefix}c#{attr_str}><#{prefix}f>CONCAT(&quot;x&quot;,&quot;y&quot;)</#{prefix}f><#{prefix}v>xy &amp; z</#{prefix}v></#{prefix}c>"
    when :err_cell
      "<#{prefix}c#{attr_str}><#{prefix}v>#DIV/0!</#{prefix}v></#{prefix}c>"
    when :self_closing_empty
      "<#{prefix}c#{attr_str}/>"
    end
  end

  def normalize_cell_value(val)
    case val
    when Xlsxrb::Elements::RichText then val.to_s
    when Xlsxrb::Elements::CellError then val.code
    else val
    end
  end

  def snapshot_stream_rows(xml_source, sst, trim:, pad_rows:, pad_cells:)
    rows = []
    Xlsxrb::Ooxml::WorksheetParser.each_row(
      xml_source,
      shared_strings: sst,
      trim_empty_rows: trim,
      pad_empty_rows: pad_rows,
      pad_empty_cells: pad_cells
    ) do |row|
      fast_vals = row.values.map { |v| normalize_cell_value(v) }
      cells = row.cells
      mat_vals = row.values.map { |v| normalize_cell_value(v) }
      fast_trimmed = fast_vals.dup
      fast_trimmed.pop while !fast_trimmed.empty? && fast_trimmed.last.nil?
      mat_trimmed = mat_vals.dup
      mat_trimmed.pop while !mat_trimmed.empty? && mat_trimmed.last.nil?
      assert_equal(fast_trimmed, mat_trimmed, "StreamRow#values fast vs materialized mismatch on row #{row.index}")

      rows << {
        index: row.index,
        attrs: [row.height, row.hidden?, row.custom_height?, row.collapsed?, row.outline_level, row.style_index],
        cells: cells.map { |cell| [cell.column_index, normalize_cell_value(cell.value), cell.formula_expression, cell.style_index] }
      }
    end
    rows
  end

  test "property: WorksheetParser multi-path equivalence across String, chunked IO, each_event, and WorksheetListener" do
    cell_gen = Pbt.tuple(
      Pbt.constant_from(*CELL_VARIANTS),
      Pbt.integer(min: 0, max: 500),
      Pbt.integer(min: 0, max: 2),
      Pbt.integer(min: 0, max: 5)
    ).map { |variant, num, gap, perm| { variant: variant, num: num, gap: gap, perm: perm } }

    row_gen = Pbt.tuple(
      Pbt.boolean,
      Pbt.array(cell_gen, min: 1, max: 4)
    ).map { |self_closing, cells| { self_closing: self_closing, cells: cells } }

    sheet_gen = Pbt.tuple(
      Pbt.array(row_gen, min: 1, max: 5),
      Pbt.boolean,
      Pbt.boolean,
      Pbt.boolean,
      Pbt.boolean,
      Pbt.constant_from(1, 2, 3, 5, 7, 13, 29, 64),
      Pbt.boolean,
      Pbt.boolean,
      Pbt.boolean
    )

    Pbt.assert(num_runs: 60) do
      Pbt.property(sheet_gen) do |row_specs, use_prefix, omit_r, with_decoy, with_hl, chunk_size, trim, pad_rows, pad_cells|
        opts = { use_prefix: use_prefix, omit_r: omit_r, with_decoy: with_decoy, with_hyperlinks: with_hl }
        xml, sst = build_generated_worksheet(row_specs, opts)

        str_rows = snapshot_stream_rows(xml, sst, trim: trim, pad_rows: pad_rows, pad_cells: pad_cells)
        io_rows = snapshot_stream_rows(ChunkedXmlIo.new(xml, chunk_size), sst, trim: trim, pad_rows: pad_rows, pad_cells: pad_cells)
        assert_equal(str_rows, io_rows, "String vs ChunkedXmlIo(chunk=#{chunk_size}) mismatch:\n#{xml}")

        enum_chunks = xml.b.scan(/.{1,#{chunk_size}}/m).each
        enum_rows = snapshot_stream_rows(enum_chunks, sst, trim: trim, pad_rows: pad_rows, pad_cells: pad_cells)
        assert_equal(str_rows, enum_rows, "String vs Enumerator(chunk=#{chunk_size}) mismatch")

        base_rows = snapshot_stream_rows(xml, sst, trim: false, pad_rows: false, pad_cells: false)
        event_rows = []
        current_cells = nil
        current_row_idx = nil
        Xlsxrb::Ooxml::WorksheetParser.each_event(xml, shared_strings: sst) do |ev|
          case ev.type
          when :row_start
            current_row_idx = ev.args[0]
            current_cells = []
          when :cell
            c_ref, _c_type, c_style, c_val, c_formula = ev.args
            col_i = c_ref ? Xlsxrb::Elements::Cell._parse_ref(c_ref)&.[](1) : current_cells.size
            current_cells << [col_i, normalize_cell_value(c_val), c_formula, c_style]
          when :row_end
            event_rows << { index: current_row_idx, cells: current_cells }
          end
        end
        assert_equal(
          base_rows.map { |entry| { index: entry[:index], cells: entry[:cells] } },
          event_rows,
          "each_row vs each_event mismatch:\n#{xml}"
        )

        unless omit_r
          listener = Xlsxrb::Ooxml::Reader::WorksheetListener.new(sst)
          Xlsxrb::Ooxml::XmlParser.parse(xml, listener)
          expected_non_nil = {}
          base_rows.each do |entry|
            entry[:cells].each do |col_i, val, f_expr, _|
              next if val.nil? && f_expr.nil?

              ref = "#{col_index_to_letter(col_i)}#{entry[:index] + 1}"
              expected_non_nil[ref] = val
            end
          end
          actual_non_nil = {}
          listener.cells.each do |ref, cell_val|
            if cell_val.is_a?(Xlsxrb::Elements::Formula)
              actual_non_nil[ref] = cell_val.cached_value
            else
              v = normalize_cell_value(cell_val)
              actual_non_nil[ref] = v unless v.nil?
            end
          end
          assert_equal(expected_non_nil, actual_non_nil, "WorksheetParser vs WorksheetListener mismatch:\n#{xml}")
        end

        hl_parsed = Xlsxrb::Ooxml::WorksheetParser.parse_hyperlinks(xml)
        hl_listener = Xlsxrb::Ooxml::Reader::HyperlinksListener.new
        Xlsxrb::Ooxml::XmlParser.parse(xml, hl_listener)
        assert_equal(hl_listener.links, hl_parsed, "parse_hyperlinks vs HyperlinksListener mismatch")
      end
    end
  end

  # --- Item 11: WorkbookParser vs WorkbookListener & RelationshipsParser vs RelsListener ---

  def escape_attr(str)
    str.to_s.gsub("&", "&amp;").gsub('"', "&quot;").gsub("<", "&lt;").gsub(">", "&gt;")
  end

  def normalize_sheet_state(state)
    case state
    when :hidden, "hidden" then :hidden
    when :very_hidden, "veryHidden" then :very_hidden
    else :visible
    end
  end

  def normalize_workbook_parser_result(parsed)
    {
      sheets: parsed[:sheets].map { |sh| [sh[:name], sh[:r_id], normalize_sheet_state(sh[:state])] },
      date1904: parsed[:date1904] ? true : false,
      code_name: parsed[:code_name],
      active_tab: parsed[:active_tab],
      first_sheet: parsed[:first_sheet],
      show_sheet_tabs: parsed[:show_sheet_tabs],
      tab_ratio: parsed[:tab_ratio],
      full_calc_on_load: parsed[:full_calc_on_load],
      workbook_protection: parsed[:workbook_protection],
      defined_names: parsed[:defined_names].map do |dn|
        [dn[:name], dn[:hidden] ? true : false, dn[:local_sheet_id], dn[:comment], dn[:value]]
      end
    }
  end

  def normalize_workbook_listener_result(listener)
    prot = listener.workbook_protection&.except(:password)
    {
      sheets: listener.sheets.map { |sh| [sh[:name], sh[:rid], normalize_sheet_state(sh[:state])] },
      date1904: listener.workbook_properties[:date1904] ? true : false,
      code_name: listener.workbook_properties[:code_name],
      active_tab: listener.workbook_views[:active_tab],
      first_sheet: listener.workbook_views[:first_sheet],
      show_sheet_tabs: listener.workbook_views[:show_sheet_tabs],
      tab_ratio: listener.workbook_views[:tab_ratio],
      full_calc_on_load: listener.calc_properties[:full_calc_on_load],
      workbook_protection: prot,
      defined_names: listener.defined_names.map do |dn|
        [dn[:name], dn[:hidden] ? true : false, dn[:local_sheet_id], dn[:comment], dn[:value]]
      end
    }
  end

  def build_workbook_xml(spec)
    sheet_count, use_prefix, date1904, prot_mode, active_tab, show_tabs, full_calc, dn_count = spec
    px = use_prefix ? "x:" : ""
    states = [nil, "hidden", "veryHidden"]
    sheets_xml = (1..sheet_count).map do |idx|
      st_attr = (st = states[(idx - 1) % states.size]) ? %( state="#{st}") : ""
      sh_name = "売上 #{idx} & <Chart>"
      %(<#{px}sheet name="#{escape_attr(sh_name)}" sheetId="#{idx}" r:id="rId#{idx}"#{st_attr}/>)
    end.join

    prot_xml = case prot_mode
               when :none then ""
               when :legacy then %(<#{px}workbookProtection workbookPassword="ABCD" lockStructure="1" lockWindows="0"/>)
               when :agile
                 %(<#{px}workbookProtection workbookAlgorithmName="SHA-512" workbookHashValue="abc==" workbookSaltValue="def==" workbookSpinCount="10000" lockStructure="true"/>)
               end

    dns_xml = if dn_count.positive?
                items = (1..dn_count).map do |idx|
                  lsi = idx == 1 ? "" : %( localSheetId="#{idx - 2}")
                  hid = idx.even? ? ' hidden="1"' : ""
                  comment_attr = %( comment="#{escape_attr("Note #{idx} & \"tip\"")}")
                  val = escape_attr("'売上 1'!$A$#{idx}:$B$#{idx + 3}")
                  %(<#{px}definedName name="Range_#{idx}"#{lsi}#{hid}#{comment_attr}>#{val}</#{px}definedName>)
                end
                "<#{px}definedNames>#{items.join}</#{px}definedNames>"
              else
                ""
              end

    <<~XML
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <#{px}workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <#{px}workbookPr date1904="#{date1904 ? "1" : "0"}" codeName="ThisWorkbook_日本語"/>
        #{prot_xml}
        <#{px}bookViews><#{px}workbookView activeTab="#{active_tab % sheet_count}" firstSheet="0" showSheetTabs="#{show_tabs ? "1" : "0"}" tabRatio="550"/></#{px}bookViews>
        <#{px}sheets>#{sheets_xml}</#{px}sheets>
        #{dns_xml}
        <#{px}calcPr fullCalcOnLoad="#{full_calc ? "true" : "false"}"/>
      </#{px}workbook>
    XML
  end

  test "property: WorkbookParser vs WorkbookListener and RelationshipsParser vs RelsListener equivalence plus malformed XML defense" do
    wb_gen = Pbt.tuple(
      Pbt.integer(min: 1, max: 4),
      Pbt.boolean,
      Pbt.boolean,
      Pbt.constant_from(:none, :legacy, :agile),
      Pbt.integer(min: 0, max: 3),
      Pbt.boolean,
      Pbt.boolean,
      Pbt.integer(min: 0, max: 3),
      Pbt.integer(min: 1, max: 4),
      Pbt.boolean
    )

    Pbt.assert(num_runs: 60) do
      Pbt.property(wb_gen) do |sc, up, d1904, pmode, atab, stabs, fcalc, dnc, rel_count, extra_decl|
        wb_xml = build_workbook_xml([sc, up, d1904, pmode, atab, stabs, fcalc, dnc])
        wb_parsed = Xlsxrb::Ooxml::WorkbookParser.parse_with_properties(wb_xml)
        wb_listener = Xlsxrb::Ooxml::Reader::WorkbookListener.new
        Xlsxrb::Ooxml::XmlParser.parse(wb_xml, wb_listener)
        assert_equal(
          normalize_workbook_parser_result(wb_parsed),
          normalize_workbook_listener_result(wb_listener),
          "WorkbookParser vs WorkbookListener mismatch:\n#{wb_xml}"
        )

        rels_entries_xml = (1..rel_count).map do |idx|
          target = escape_attr("worksheets/sheet_#{idx}.xml?a=1&b=2")
          type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
          %(<Relationship Id="rId#{idx}" Type="#{type}" Target="#{target}"/>)
        end.join
        clean_rels_xml = <<~XML
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            #{rels_entries_xml}
          </Relationships>
        XML
        rels_parsed = Xlsxrb::Ooxml::RelationshipsParser.parse_all(clean_rels_xml)
        rels_listener = Xlsxrb::Ooxml::Reader::RelsListener.new
        Xlsxrb::Ooxml::XmlParser.parse(clean_rels_xml, rels_listener)
        assert_equal(rels_listener.relationships, rels_parsed[:entries])

        if extra_decl
          dup_decl_rels = "<?xml version=\"1.0\"?>#{clean_rels_xml}"
          assert_equal(rels_parsed, Xlsxrb::Ooxml::RelationshipsParser.parse_all(dup_decl_rels))
        end

        bad_rels_samples = [
          "<Relationships><Relationship Id='rId1' Type='t' Target='sheet1.xml'/><Relationship Id='rId2'",
          "<?xml version=\"1.0\"?><Relationships><Relationship Id=\"rId1\" Type=\"t\" Target=\"x\xFF\"/>"
        ]
        bad_rels_samples.each do |bad_xml|
          res = Xlsxrb::Ooxml::RelationshipsParser.parse_all(bad_xml)
          assert_kind_of(Hash, res)
          assert_includes(res[:targets].keys, "rId1")
        end

        deep_open = "<a>" * 270
        deep_close = "</a>" * 270
        bad_wb_samples = [
          "<workbook><sheets><sheet name='A'</workbook>",
          "<workbook>#{deep_open}#{deep_close}</workbook>",
          "<workbook><sheets><sheet name='A'\xFF/></sheets></workbook>"
        ]
        bad_wb_samples.each do |bad_wb|
          assert_raise(Xlsxrb::ParseError) { Xlsxrb::Ooxml::WorkbookParser.parse_with_properties(bad_wb) }
        end
      end
    end
  end

  # --- Item 12: ZipReader CD-less Data-Descriptor fallback Stateful PBT ---

  ZipState = Data.define(:entries, :dirty)

  # Builds CD-less ZIP archives (stored, deflated, signature/nosig data descriptors, ZIP64 DD, dirs, BOM).
  class CdLessZipSut
    SIG_DD = [0x50, 0x4B, 0x07, 0x08].pack("C4").freeze
    FORBIDDEN_CRCS = [
      "PK\x07\x08".unpack1("V"),
      "PK\x03\x04".unpack1("V"),
      "PK\x01\x02".unpack1("V")
    ].freeze

    attr_reader :entries_model

    def initialize
      @chunks = []
      @entries_model = []
    end

    def add_stored(name, payload)
      data = payload.b
      crc = Zlib.crc32(data)
      header = [20, 0, 0, 0, 33, crc, data.bytesize, data.bytesize, name.bytesize, 0].pack("vvvvvVVVvv")
      @chunks << "PK\x03\x04".b << header << name.b << data
      @entries_model << { name: name, data: data, method: 0, crc32: crc, csize: data.bytesize, usize: data.bytesize }
      nil
    end

    def add_deflated(name, payload)
      data = payload.b
      crc = Zlib.crc32(data)
      comp = raw_deflate(data)
      header = [20, 0, 8, 0, 33, crc, comp.bytesize, data.bytesize, name.bytesize, 0].pack("vvvvvVVVvv")
      @chunks << "PK\x03\x04".b << header << name.b << comp
      @entries_model << { name: name, data: data, method: 8, crc32: crc, csize: comp.bytesize, usize: data.bytesize }
      nil
    end

    def add_dd(name, payload, with_sig:, zip64:)
      data = payload.b
      unless with_sig
        # Guard against a CRC matching a ZIP record signature when descriptor has no signature.
        data += "\x00".b while FORBIDDEN_CRCS.include?(Zlib.crc32(data))
      end
      crc = Zlib.crc32(data)
      comp = raw_deflate(data)
      ver = zip64 ? 45 : 20
      header = [ver, 0x08, 8, 0, 33, 0, 0, 0, name.bytesize, 0].pack("vvvvvVVVvv")
      desc = if zip64
               [crc, comp.bytesize, data.bytesize].pack("VQ<Q<")
             else
               [crc, comp.bytesize, data.bytesize].pack("VVV")
             end
      @chunks << "PK\x03\x04".b << header << name.b << comp
      @chunks << SIG_DD if with_sig
      @chunks << desc
      @entries_model << { name: name, data: data, method: 8, crc32: crc, csize: comp.bytesize, usize: data.bytesize }
      nil
    end

    def add_dir(dir_name)
      header = [20, 0, 0, 0, 33, 0, 0, 0, dir_name.bytesize, 0].pack("vvvvvVVVvv")
      @chunks << "PK\x03\x04".b << header << dir_name.b
      nil
    end

    def add_bom_xml(name, raw_bytes, utf8_expected)
      crc = Zlib.crc32(raw_bytes)
      comp = raw_deflate(raw_bytes)
      header = [20, 0x08, 8, 0, 33, 0, 0, 0, name.bytesize, 0].pack("vvvvvVVVvv")
      desc = [crc, comp.bytesize, raw_bytes.bytesize].pack("VVV")
      @chunks << "PK\x03\x04".b << header << name.b << comp << SIG_DD << desc
      @entries_model << {
        name: name,
        data: utf8_expected.b,
        method: 8,
        crc32: crc,
        csize: comp.bytesize,
        usize: raw_bytes.bytesize
      }
      nil
    end

    def archive_bytes
      @chunks.join.b
    end

    private

    def raw_deflate(data)
      deflater = Zlib::Deflate.new(Zlib::DEFAULT_COMPRESSION, -Zlib::MAX_WBITS)
      deflater.deflate(data, Zlib::FINISH)
    ensure
      deflater&.close
    end
  end

  def make_payload(size, seed_byte)
    return "".b if size.zero?

    pattern = [seed_byte.ord, ((seed_byte.ord + 13) % 251) + 1].pack("CC")
    (pattern * ((size / 2) + 1)).byteslice(0, size)
  end

  BOM_EXPECTED_TEXT = "<root>Hello 日本語 &amp; OK</root>"
  BOM_SAMPLES = {
    utf8: "\xEF\xBB\xBF#{BOM_EXPECTED_TEXT}",
    utf16_le: (+"\xFF\xFE").b << BOM_EXPECTED_TEXT.encode(Encoding::UTF_16LE).b,
    utf16_be: (+"\xFE\xFF").b << BOM_EXPECTED_TEXT.encode(Encoding::UTF_16BE).b
  }.freeze

  def verify_cd_less_zip_reader(zip, model)
    expected_names = model.map { |entry| entry[:name] }
    assert_equal(expected_names, zip.entry_names)

    model.each do |entry|
      name = entry[:name]
      expected = entry[:data]
      cat_entry = zip.entry_catalog[name]
      assert_not_nil(cat_entry)
      assert_equal(entry[:csize], cat_entry.compressed_size)
      assert_equal(entry[:usize], cat_entry.uncompressed_size)
      assert_equal(entry[:crc32], cat_entry.crc32)

      read_full = zip.read_entry(name)
      assert_equal(expected, read_full.b, "read_entry mismatch for #{name}")
      assert_equal(read_full, zip.read_entry(name.upcase))
      assert_equal(read_full, zip.read_entry(name.tr("/", "\\")))

      stream = zip.open_entry_io(name, encoding: false)
      begin
        zero_chunk = stream.read(0)
        assert_equal("".b, zero_chunk)
        assert_false(zero_chunk.frozen?, "read(0) must return an unfrozen String")

        collected = +"".b
        [1, 7, 513].each do |read_sz|
          part = stream.read(read_sz)
          collected << part if part
        end
        rest = stream.read(nil)
        collected << rest if rest
        assert_true(stream.eof?)
        assert_nil(stream.read(1))
        assert_equal(expected, collected, "EntryStreamIO piecewise read mismatch for #{name}")
      ensure
        stream.close
        assert_true(stream.closed?)
      end

      chunked = +"".b
      zip.each_entry_chunk(name, chunk_size: 17, encoding: false) { |ch| chunked << ch }
      assert_equal(expected, chunked, "each_entry_chunk mismatch for #{name}")
    end
  end

  test "stateful property: ZipReader CD-less data-descriptor fallback, EntryStreamIO, copy_to_writer, and ZipBombError" do
    init_state = ZipState.new(entries: [], dirty: false)

    Pbt.assert(num_runs: 25) do
      Pbt.stateful(max_steps: 8) do |m|
        m.initial_state Pbt.constant(init_state)
        m.sut { |_init| CdLessZipSut.new }

        m.command :add_stored, weight: 2 do |c|
          c.arguments(size: Pbt.integer(min: 0, max: 2000), seed: Pbt.constant_from("a", "Z", "\x7F"))
          c.next_state do |s, args|
            name = "stored/part_#{s.entries.size}.bin"
            s.with(entries: [*s.entries, { name: name, size: args[:size] }], dirty: true)
          end
          c.run do |sut, args|
            name = "stored/part_#{sut.entries_model.size}.bin"
            sut.add_stored(name, make_payload(args[:size], args[:seed]))
          end
        end

        m.command :add_deflated, weight: 2 do |c|
          c.arguments(size: Pbt.integer(min: 0, max: 5000), seed: Pbt.constant_from("b", "M", "\x20"))
          c.next_state do |s, args|
            name = "deflated/part_#{s.entries.size}.bin"
            s.with(entries: [*s.entries, { name: name, size: args[:size] }], dirty: true)
          end
          c.run do |sut, args|
            name = "deflated/part_#{sut.entries_model.size}.bin"
            sut.add_deflated(name, make_payload(args[:size], args[:seed]))
          end
        end

        m.command :add_dd_sig, weight: 3 do |c|
          c.arguments(
            size: Pbt.constant_from(0, 17, 4095, 4096, 4097, 6500),
            seed: Pbt.constant_from("c", "Q"),
            zip64: Pbt.boolean
          )
          c.next_state do |s, args|
            name = "dd_sig/part_#{s.entries.size}.bin"
            s.with(entries: [*s.entries, { name: name, size: args[:size] }], dirty: true)
          end
          c.run do |sut, args|
            name = "dd_sig/part_#{sut.entries_model.size}.bin"
            sut.add_dd(name, make_payload(args[:size], args[:seed]), with_sig: true, zip64: args[:zip64])
          end
        end

        m.command :add_dd_nosig, weight: 3 do |c|
          c.arguments(
            size: Pbt.constant_from(1, 300, 4100, 5500),
            seed: Pbt.constant_from("d", "R"),
            zip64: Pbt.boolean
          )
          c.next_state do |s, args|
            name = "dd_nosig/part_#{s.entries.size}.bin"
            s.with(entries: [*s.entries, { name: name, size: args[:size] }], dirty: true)
          end
          c.run do |sut, args|
            name = "dd_nosig/part_#{sut.entries_model.size}.bin"
            sut.add_dd(name, make_payload(args[:size], args[:seed]), with_sig: false, zip64: args[:zip64])
          end
        end

        m.command :add_dir, weight: 1 do |c|
          c.next_state { |s| s }
          c.run { |sut, _args| sut.add_dir("dir_#{sut.entries_model.size}/") }
        end

        m.command :add_bom_xml, weight: 2 do |c|
          c.arguments(bom_kind: Pbt.constant_from(:utf8, :utf16_le, :utf16_be))
          c.next_state do |s, _args|
            name = "bom/part_#{s.entries.size}.xml"
            s.with(entries: [*s.entries, { name: name, size: BOM_EXPECTED_TEXT.bytesize }], dirty: true)
          end
          c.run do |sut, args|
            name = "bom/part_#{sut.entries_model.size}.xml"
            sut.add_bom_xml(name, BOM_SAMPLES.fetch(args[:bom_kind]), BOM_EXPECTED_TEXT)
          end
        end

        m.invariant(:zip_reader_cd_less_consistency) do |state, sut|
          next unless state.dirty

          archive = sut.archive_bytes
          model = sut.entries_model

          Xlsxrb::Ooxml::ZipReader.open(archive) do |zip|
            verify_cd_less_zip_reader(zip, model)

            repacked_io = StringIO.new(+"".b)
            Xlsxrb::Ooxml::ZipWriter.open(repacked_io) do |zw|
              zip.entry_names.each { |name| zip.copy_to_writer(name, zw) }
            end
            Xlsxrb::Ooxml::ZipReader.open(repacked_io.string) do |repacked|
              assert_equal(zip.read_all, repacked.read_all)
            end
          end

          chunk_enum = Enumerator.new do |yielder|
            off = 0
            while off < archive.bytesize
              yielder << archive.byteslice(off, 193)
              off += 193
            end
          end
          stream_zip = Xlsxrb::Ooxml::ZipReader.new(chunk_enum)
          begin
            verify_cd_less_zip_reader(stream_zip, model)
          ensure
            stream_zip.close
          end

          oversized = model.find { |entry| entry[:data].bytesize > 64 }
          if oversized
            assert_raise(Xlsxrb::ZipBombError) do
              Xlsxrb::Ooxml::ZipReader.open(archive, max_uncompressed_size: 64, &:read_all)
            end
          end
        end
      end
    end
  end
end
