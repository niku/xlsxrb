# frozen_string_literal: true

require "test_helper"
require "pbt"
require "tempfile"
require "date"
require "time"
require_relative "../support/excel_rules"

class PbtTest < Test::Unit::TestCase
  def setup
    omit "Skip property-based tests under RBS runtime testing to prevent OOM / slowdown" if defined?(RBS::Test)
  end

  # --- PBT Arbitraries & Generators for Union Types and Structures ---

  # Generates sheet names (1-31 chars, alphanumeric)
  def sheet_name_generator
    Pbt.alphanumeric_string(min: 1, max: 31)
  end

  # Generates raw descriptors for the entire union type domain:
  # [type_code (0..11), int_val, float_val, bool_val, str_val, sub_choice (0..7)]
  def cell_value_descriptor_generator
    Pbt.tuple(
      Pbt.choose(0..11),
      Pbt.integer,
      Pbt.float.filter { |f| !f.nan? && !f.infinite? },
      Pbt.boolean,
      Pbt.printable_ascii_string(max: 20),
      Pbt.choose(0..7)
    )
  end

  # Maps a raw descriptor tuple to an actual cell value in the union type domain:
  # - 0: nil
  # - 1: Integer
  # - 2: Float
  # - 3: Boolean (true/false)
  # - 4: ASCII String
  # - 5: Special Strings (XML entities <>&"', UTF-8, Japanese, Emoji, leading '=')
  # - 6: Date
  # - 7: Time
  # - 8: Elements::Formula (expression, cached_value, calculate_always)
  # - 9: Hash formula representation ({ formula:, value:, calculate_always: })
  # - 10: Elements::CellError (#N/A, #DIV/0!, #VALUE!, #REF!, etc.)
  # - 11: Elements::RichText (multiple runs with font formatting)
  def resolve_cell_value(tuple)
    type_code, int_val, float_val, bool_val, str_val, sub_choice = tuple
    case type_code
    when 0
      nil
    when 1
      int_val
    when 2
      (float_val % 10_000.0).round(4)
    when 3
      bool_val
    when 4
      str_val
    when 5 # Special strings
      case sub_choice
      when 0 then "<tag>&'\"</tag>"
      when 1 then "10 < 20 & 30 > 5"
      when 2 then "Quote: \"Hello & World\""
      when 3 then "Line1\nLine2\tTabbed"
      when 4 then "日本語テキスト（漢字・ひらがな）_#{str_val}"
      when 5 then "Emoji 🎉 🚀 📊 💡 🧪 ✨_#{str_val}"
      when 6 then "Mixed: UTF8-日本語-Emoji-👍-<xml>&_#{str_val}"
      when 7 then "=SUM(A1:A#{(int_val.abs % 50) + 1})"
      end
    when 6 # Date
      year = 1900 + (int_val.abs % 150)
      month = 1 + (sub_choice % 12)
      day = 1 + (int_val.abs % 28)
      Date.new(year, month, day)
    when 7 # Time
      year = 1900 + (int_val.abs % 150)
      month = 1 + (sub_choice % 12)
      day = 1 + (int_val.abs % 28)
      hour = int_val.abs % 24
      min = (sub_choice * 7) % 60
      sec = int_val.abs % 60
      Time.utc(year, month, day, hour, min, sec)
    when 8 # Elements::Formula
      expressions = %w[SUM(A1:A10) AVERAGE(B1:B10) A1*B1+C1 IF(A1>0,1,0) MAX(D1:D5) CONCATENATE(A1,B1)]
      expr = expressions[sub_choice % expressions.size]
      cached = bool_val ? (int_val % 1000) : nil
      Xlsxrb::Elements::Formula.new(expression: expr, cached_value: cached, calculate_always: bool_val)
    when 9 # Hash formula
      expressions = %w[SUM(A1:A5) COUNT(C1:C10) B1+B2 PRODUCT(A1:B2)]
      expr = expressions[sub_choice % expressions.size]
      { formula: expr, value: bool_val ? (int_val.abs % 500) : nil, calculate_always: bool_val }
    when 10 # Elements::CellError
      code = Xlsxrb::Elements::VALID_ERROR_CODES[sub_choice % Xlsxrb::Elements::VALID_ERROR_CODES.size]
      Xlsxrb::Elements::CellError.new(code: code)
    when 11 # Elements::RichText
      runs = [
        { text: str_val.empty? ? "Run1" : str_val, font: { bold: bool_val, color: "FF0000" } },
        { text: " Run2_#{sub_choice}", font: { italic: !bool_val, sz: 12 } }
      ]
      Xlsxrb::Elements::RichText.new(runs: runs)
    end
  end

  # Style options generator and resolver
  def style_options_descriptor_generator
    Pbt.tuple(
      Pbt.boolean, # bold
      Pbt.boolean, # italic
      Pbt.boolean, # underline
      Pbt.integer(min: 8, max: 24), # sz
      Pbt.choose(0..3), # color
      Pbt.choose(0..3), # fill_color
      Pbt.choose(0..2), # horizontal
      Pbt.choose(0..2), # vertical
      Pbt.choose(0..3)  # number_format
    )
  end

  def resolve_style_options(tuple)
    bold, italic, underline, sz, color_idx, fill_idx, h_idx, v_idx, nf_idx = tuple
    colors = %w[FF0000 00FF00 0000FF FFFF00]
    horizontals = %i[left center right]
    verticals = %i[top center bottom]
    nfs = ["#,##0.00", "0.0%", "yyyy-mm-dd", "@"]
    {
      bold: bold,
      italic: italic,
      underline: underline,
      font_size: sz,
      font_color: colors[color_idx],
      fill_color: colors[fill_idx],
      align_horizontal: horizontals[h_idx],
      align_vertical: verticals[v_idx],
      number_format: nfs[nf_idx]
    }
  end

  # Sheet options generator and resolver
  def sheet_options_descriptor_generator
    Pbt.tuple(
      Pbt.boolean, # fit_to_page
      Pbt.choose(0..3), # tab_color
      Pbt.integer(min: 50, max: 200), # zoom_scale
      Pbt.integer(min: 0, max: 20), # freeze_row
      Pbt.integer(min: 0, max: 10), # freeze_col
      Pbt.boolean, # protect_sheet
      Pbt.alphanumeric_string(min: 1, max: 10) # password
    )
  end

  def resolve_sheet_options(tuple)
    fit_to_page, tab_color_idx, zoom_scale, freeze_row, freeze_col, protect_sheet, password = tuple
    colors = %w[FF0000 00FF00 0000FF FFFF00]
    {
      fit_to_page: fit_to_page,
      tab_color: colors[tab_color_idx],
      zoom_scale: zoom_scale,
      freeze_row: freeze_row,
      freeze_col: freeze_col,
      protect_sheet: protect_sheet,
      password: password
    }
  end

  # --- Property Tests ---

  test "property-based test: comprehensive build and read consistency across full union type domain" do
    Pbt.assert(num_runs: 35) do
      Pbt.property(
        Pbt.array(sheet_name_generator, min: 1, max: 3).filter { |a| a.size == a.uniq.size && !a.empty? },
        Pbt.array(Pbt.array(cell_value_descriptor_generator, min: 1, max: 5), min: 1, max: 5)
      ) do |sheet_names, raw_rows_data|
        rows_data = raw_rows_data.map do |row|
          row.map { |cell_desc| resolve_cell_value(cell_desc) }
        end

        workbook = Xlsxrb.build do |w|
          sheet_names.each do |sname|
            w.sheet(sname) do |s|
              rows_data.each do |row_data|
                s.row(row_data)
              end
            end
          end
        end

        tmp = Tempfile.new(["pbt_union_test", ".xlsx"])
        begin
          Xlsxrb.write(tmp.path, workbook)

          read_wb = Xlsxrb.load(tmp.path)

          assert_equal sheet_names.size, read_wb.sheets.size

          sheet_names.each_with_index do |sname, sheet_idx|
            sheet = read_wb.sheets[sheet_idx]
            assert_equal sname, sheet.name

            rows_data.each_with_index do |expected_row, row_idx|
              actual_row = sheet.rows.find { |r| r.index == row_idx }

              next if expected_row.all?(&:nil?)

              expected_row.each_with_index do |expected_val, col_idx|
                actual_cell = actual_row&.cells&.find { |c| c.column_index == col_idx }

                if expected_val.nil?
                  assert_nil actual_cell&.value
                  next
                end

                assert_not_nil actual_cell, "Cell at row #{row_idx}, col #{col_idx} should exist for expected #{expected_val.inspect}"

                actual_val = actual_cell.value
                actual_formula = actual_cell.formula

                case expected_val
                when Xlsxrb::Elements::Formula
                  expected_expr = expected_val.expression.delete_prefix("=")
                  assert_equal expected_expr, actual_formula&.expression
                  assert_equal expected_val.cached_value, actual_val unless expected_val.cached_value.nil?
                when Hash
                  expected_expr = expected_val[:formula].delete_prefix("=")
                  assert_equal expected_expr, actual_formula&.expression
                when Xlsxrb::Elements::CellError
                  assert_equal expected_val.code, actual_val.to_s
                when Xlsxrb::Elements::RichText
                  assert_equal expected_val.to_s, actual_val.to_s
                when Date
                  assert_equal expected_val, actual_cell.to_date
                when Time
                  actual_time = actual_cell.to_time
                  assert_not_nil actual_time
                  assert_in_delta expected_val.to_i, actual_time.to_i, 2
                when Float
                  assert_in_delta expected_val, actual_val.to_f, 0.001
                when Integer, true, false
                  assert_equal expected_val, actual_val
                when String
                  assert_equal expected_val, actual_val.to_s
                end
              end
            end
          end
        ensure
          tmp.close!
        end
      end
    end
  end

  test "property-based test: stream writer vs workbook builder equivalence for union types" do
    Pbt.assert(num_runs: 25) do
      Pbt.property(
        sheet_name_generator,
        Pbt.array(Pbt.array(cell_value_descriptor_generator, min: 1, max: 5), min: 1, max: 5)
      ) do |sheet_name, raw_rows_data|
        rows_data = raw_rows_data.map do |row|
          row.map { |cell_desc| resolve_cell_value(cell_desc) }
        end

        Dir.mktmpdir do |dir|
          streaming_path = File.join(dir, "streaming_union.xlsx")
          in_memory_path = File.join(dir, "in_memory_union.xlsx")

          # 1. Streaming write
          Xlsxrb.write(streaming_path) do |wb|
            wb.sheet(sheet_name) do |s|
              rows_data.each { |r| s.row(r) }
            end
          end

          # 2. In-memory write
          wb = Xlsxrb.build do |b|
            b.sheet(sheet_name) do |s|
              rows_data.each { |r| s.row(r) }
            end
          end
          Xlsxrb.write(in_memory_path, wb)

          # Invariant: Both write modes yield identical sheets and cell values
          streaming_wb = Xlsxrb.load(streaming_path)
          in_memory_wb = Xlsxrb.load(in_memory_path)

          assert_equal 1, streaming_wb.sheets.size
          assert_equal 1, in_memory_wb.sheets.size
          assert_equal streaming_wb.sheets[0].name, in_memory_wb.sheets[0].name

          streaming_sheet = streaming_wb.sheets[0]
          in_memory_sheet = in_memory_wb.sheets[0]

          assert_equal streaming_sheet.rows.size, in_memory_sheet.rows.size
          streaming_sheet.rows.each_with_index do |s_row, r_idx|
            m_row = in_memory_sheet.rows[r_idx]
            assert_equal s_row.index, m_row.index
            assert_equal s_row.cells.size, m_row.cells.size

            s_row.cells.each_with_index do |s_cell, c_idx|
              m_cell = m_row.cells[c_idx]
              assert_equal s_cell.column_index, m_cell.column_index
              assert_equal s_cell.ref, m_cell.ref
              assert_equal s_cell.formula, m_cell.formula

              if s_cell.value.is_a?(Float)
                assert_in_delta s_cell.value, m_cell.value, 0.0001
              else
                assert_equal s_cell.value, m_cell.value
              end
            end
          end
        end
      end
    end
  end

  test "property-based test: cell coordinate indexing bijection and reference parsing" do
    Pbt.assert(num_runs: 50) do
      Pbt.property(
        Pbt.choose(0..16_383),
        Pbt.choose(0..1_048_575)
      ) do |col_idx, row_idx|
        # 1. Bijection: column_letter <-> column_index
        letter = Xlsxrb.column_letter(col_idx)
        assert_equal col_idx, Xlsxrb.column_index(letter)
        assert_equal col_idx, Xlsxrb.column_index(letter.downcase)
        assert_equal col_idx, Xlsxrb.column_index(letter.to_sym)

        # 2. Reference bijection: parse_ref <-> ref
        ref = "#{letter}#{row_idx + 1}"
        parsed = Xlsxrb.parse_ref(ref)
        assert_equal [row_idx, col_idx], parsed

        parsed_lower = Xlsxrb.parse_ref(ref.downcase)
        assert_equal [row_idx, col_idx], parsed_lower

        # 3. Cell object reference and validation
        cell = Xlsxrb::Elements::Cell.new(row_index: row_idx, column_index: col_idx, value: 42)
        assert_equal ref, cell.ref
        assert_true cell.valid?
        assert_empty cell.errors
      end
    end
  end

  test "property-based test: cell model type coercions and accessor invariants" do
    Pbt.assert(num_runs: 40) do
      Pbt.property(
        cell_value_descriptor_generator,
        Pbt.choose(0..100),
        Pbt.choose(0..50)
      ) do |cell_desc, row_idx, col_idx|
        val = resolve_cell_value(cell_desc)
        cell = Xlsxrb::Elements::Cell.new(row_index: row_idx, column_index: col_idx, value: val)

        assert_true cell.valid?
        assert_equal val, cell.value
        assert_equal row_idx, cell.row_index
        assert_equal col_idx, cell.column_index

        # Coercion invariant checks
        case val
        when Integer
          assert_equal val, cell.to_i
          assert_equal val.to_f, cell.to_f
          assert_equal val.to_s, cell.to_s
        when Float
          assert_equal val.to_i, cell.to_i
          assert_equal val, cell.to_f
        when Date
          assert_equal val, cell.to_date
        when Time
          assert_equal val, cell.to_time
        when Xlsxrb::Elements::CellError
          assert_equal val.code, cell.to_s
        when Xlsxrb::Elements::RichText, true, false, String
          assert_equal val.to_s, cell.to_s
        when nil
          assert_equal "", cell.to_s
        end
      end
    end
  end

  test "property-based test: row and column struct options and indexing invariants" do
    Pbt.assert(num_runs: 40) do
      Pbt.property(
        Pbt.choose(0..1000),
        Pbt.tuple(
          Pbt.choose(0..1), Pbt.choose(0..409), # height (nil or 0..409)
          Pbt.boolean, # hidden
          Pbt.boolean, # custom_height
          Pbt.choose(0..1), Pbt.choose(0..7) # outline_level (nil or 0..7)
        ),
        Pbt.tuple(
          Pbt.choose(0..1), Pbt.choose(0..255), # width (nil or 0..255)
          Pbt.boolean, # hidden
          Pbt.boolean, # custom_width
          Pbt.choose(0..1), Pbt.choose(0..7) # outline_level (nil or 0..7)
        )
      ) do |idx, row_desc, col_desc|
        row_height = row_desc[0].zero? ? nil : row_desc[1]
        row_hidden = row_desc[2]
        row_custom_height = row_desc[3]
        row_outline = row_desc[4].zero? ? nil : row_desc[5]

        col_width = col_desc[0].zero? ? nil : col_desc[1]
        col_hidden = col_desc[2]
        col_custom_width = col_desc[3]
        col_outline = col_desc[4].zero? ? nil : col_desc[5]

        # Row struct invariants
        cells = [
          Xlsxrb::Elements::Cell.new(row_index: idx, column_index: 0, value: "A"),
          Xlsxrb::Elements::Cell.new(row_index: idx, column_index: 2, value: "C")
        ]
        row = Xlsxrb::Elements::Row.new(
          index: idx,
          cells: cells,
          height: row_height,
          hidden: row_hidden,
          custom_height: row_custom_height,
          outline_level: row_outline
        )

        assert_true row.valid?
        assert_equal idx, row.index
        assert_equal row_height, row.height
        assert_equal row_hidden, row.hidden?
        assert_equal row_custom_height, row.custom_height?
        assert_equal row_outline, row.outline_level

        assert_equal cells[0], row[0]
        assert_nil row[1]
        assert_equal cells[1], row[2]
        assert_equal cells[0], row.cell_at(0)
        assert_nil row.cell_at(1)
        assert_equal cells[1], row.cell_at(2)

        assert_equal cells, row.to_a
        assert_equal ["A", nil, "C"], row.values
        assert_equal 2, row.each.to_a.size

        # Column struct invariants
        col = Xlsxrb::Elements::Column.new(
          index: idx % 16_384,
          width: col_width,
          hidden: col_hidden,
          custom_width: col_custom_width,
          outline_level: col_outline
        )
        assert_true col.valid?
        assert_equal idx % 16_384, col.index
        assert_equal col_width, col.width
        assert_equal col_hidden, col.hidden?
        assert_equal col_custom_width, col.custom_width?
        assert_equal col_outline, col.outline_level
      end
    end
  end

  test "property-based test: worksheet coordinate access and immutable update invariants" do
    Pbt.assert(num_runs: 30) do
      Pbt.property(
        sheet_name_generator,
        Pbt.array(Pbt.tuple(Pbt.choose(0..4), Pbt.choose(0..4), Pbt.choose(1..100)), min: 1, max: 8),
        Pbt.choose(1000..2000)
      ) do |sheet_name, cell_specs, _update_val|
        unique_cells = {}
        cell_specs.each do |r, c, v|
          unique_cells[[r, c]] = v
        end

        max_r = unique_cells.keys.map(&:first).max || 0
        wb = Xlsxrb.build do |b|
          b.sheet(sheet_name) do |s|
            (0..max_r).each do |r|
              row_vals = []
              unique_cells.each do |(cell_r, cell_c), v|
                row_vals[cell_c] = v if cell_r == r
              end
              s.row(row_vals)
            end
          end
        end

        sheet = wb.sheets.first

        # Invariant: Random coordinate access and 2D access equivalence
        unique_cells.each do |(r, c), expected_val|
          ref = "#{Xlsxrb.column_letter(c)}#{r + 1}"
          cell = sheet[ref]
          assert_not_nil cell
          assert_equal expected_val, cell.value

          cell_2d = sheet[r, c]
          assert_not_nil cell_2d
          assert_equal expected_val, cell_2d.value
        end
      end
    end
  end

  test "property-based test: style and number formatting options invariants" do
    Pbt.assert(num_runs: 25) do
      Pbt.property(
        style_options_descriptor_generator,
        Pbt.array(Pbt.tuple(Pbt.alphanumeric_string(min: 1, max: 10), Pbt.integer), min: 1, max: 4)
      ) do |style_desc, rows|
        style_opts = resolve_style_options(style_desc)
        Dir.mktmpdir do |dir|
          xlsx_path = File.join(dir, "styled_pbt.xlsx")

          Xlsxrb.write(xlsx_path) do |w|
            w.sheet("StyledSheet") do |s|
              s.style(:custom_style, **style_opts)
              rows.each do |label, num|
                s.row([label, num], cell_styles: %i[custom_style custom_style])
              end
            end
          end

          assert_true File.exist?(xlsx_path)

          # Invariant: File must load and parse valid styles without XML parse errors
          wb = Xlsxrb.load(xlsx_path)
          assert_equal 1, wb.sheets.size
          sheet = wb.sheets.first
          assert_equal rows.size, sheet.rows.size
          assert_not_nil wb.styles
        end
      end
    end
  end

  test "property-based test: sheet view, freeze panes, and protection option invariants" do
    Pbt.assert(num_runs: 25) do
      Pbt.property(
        sheet_options_descriptor_generator
      ) do |sheet_desc|
        opts = resolve_sheet_options(sheet_desc)
        Dir.mktmpdir do |dir|
          xlsx_path = File.join(dir, "options_pbt.xlsx")

          Xlsxrb.write(xlsx_path) do |w|
            w.sheet("ConfiguredSheet") do |s|
              s.freeze_pane(row: opts[:freeze_row], column: opts[:freeze_col]) if opts[:freeze_row].positive? || opts[:freeze_col].positive?
              s.sheet_view(zoom_scale: opts[:zoom_scale])
              s.sheet_properties(tab_color: opts[:tab_color])
              s.sheet_protection(sheet: opts[:protect_sheet], password: opts[:password])
              s.row(%w[Col1 Col2 Col3])
              s.row([100, 200, 300])
            end
          end

          # Invariant: Complex options roundtrip and load cleanly without XML corruption
          wb = Xlsxrb.load(xlsx_path)
          assert_equal 1, wb.sheets.size
          assert_equal "ConfiguredSheet", wb.sheets.first.name
          assert_equal 2, wb.sheets.first.rows.size
        end
      end
    end
  end

  test "property-based test: strict excel mode validates numeric extremes" do
    Pbt.assert(num_runs: 50) do
      # 1. Invalid floats
      invalid_float_arb = Pbt.constant_from(Float::NAN, Float::INFINITY, -Float::INFINITY)
      Pbt.property(invalid_float_arb) do |bad_float|
        assert_raise(ArgumentError) do
          Xlsxrb.build(strict_excel_mode: true) { |w| w.sheet("S1") { |s| s.row [bad_float] } }
        end
      end

      # 2. Invalid row heights
      invalid_height_arb = Pbt.integer.filter { |h| h.negative? || h > 409 }
      Pbt.property(invalid_height_arb) do |bad_height|
        assert_raise(ArgumentError) do
          Xlsxrb.build(strict_excel_mode: true) { |w| w.sheet("S1") { |s| s.row [1], height: bad_height } }
        end
      end

      # 3. Invalid column widths
      invalid_width_arb = Pbt.integer.filter { |w| w.negative? || w > 255 }
      Pbt.property(invalid_width_arb) do |bad_width|
        assert_raise(ArgumentError) do
          Xlsxrb.build(strict_excel_mode: true) { |w| w.sheet("S1") { |s| s.column 0, width: bad_width } }
        end
      end

      # 4. Sheet name > 31 chars
      long_sheet_arb = Pbt.alphanumeric_string(min: 32, max: 100)
      Pbt.property(long_sheet_arb) do |bad_sheet_name|
        assert_raise(ArgumentError) do
          Xlsxrb.build(strict_excel_mode: true) { |w| w.sheet(bad_sheet_name) }
        end
      end

      # 5. Invalid sheet characters
      invalid_chars_arb = Pbt.one_of(
        Pbt.constant("["), Pbt.constant("]"), Pbt.constant("*"),
        Pbt.constant("?"), Pbt.constant("/"), Pbt.constant("\\"), Pbt.constant(":")
      )
      Pbt.property(invalid_chars_arb) do |bad_char|
        assert_raise(ArgumentError) do
          # Create a sheet name that contains the invalid character
          bad_sheet_name = "Sheet#{bad_char}1"
          Xlsxrb.build(strict_excel_mode: true) { |w| w.sheet(bad_sheet_name) }
        end
      end

      # 6. String length > 32,767
      huge_string_arb = Pbt.constant("a" * 32_768)
      Pbt.property(huge_string_arb) do |bad_str|
        assert_raise(ArgumentError) do
          Xlsxrb.build(strict_excel_mode: true) { |w| w.sheet("S1") { |s| s.row [bad_str] } }
        end
      end
    end
  end

  test "property-based test: strict excel mode allows valid boundaries" do
    Pbt.assert(num_runs: 50) do
      # 1. Valid floats
      valid_float_arb = Pbt.float.filter { |f| !f.nan? && !f.infinite? }
      Pbt.property(valid_float_arb) do |good_float|
        # Should not raise ArgumentError
        Xlsxrb.build(strict_excel_mode: true) { |w| w.sheet("S1") { |s| s.row [good_float] } }
      end

      # 2. Valid row heights (0..409)
      valid_height_arb = Pbt.integer(min: 0, max: 409)
      Pbt.property(valid_height_arb) do |good_height|
        Xlsxrb.build(strict_excel_mode: true) { |w| w.sheet("S1") { |s| s.row [1], height: good_height } }
      end

      # 3. Valid column widths (0..255)
      valid_width_arb = Pbt.integer(min: 0, max: 255)
      Pbt.property(valid_width_arb) do |good_width|
        Xlsxrb.build(strict_excel_mode: true) { |w| w.sheet("S1") { |s| s.column 0, width: good_width } }
      end

      # 4. Valid sheet name lengths (1..31) and characters
      valid_sheet_arb = Pbt.printable_ascii_string(min: 1, max: 31).filter { |s| Xlsxrb::Elements::Worksheet.valid_name?(s) }
      Pbt.property(valid_sheet_arb) do |good_sheet_name|
        Xlsxrb.build(strict_excel_mode: true) { |w| w.sheet(good_sheet_name) }
      end

      # 5. Valid string length (<= 32,767)
      valid_string_arb = Pbt.printable_ascii_string(min: 0, max: 32_767)
      Pbt.property(valid_string_arb) do |good_str|
        Xlsxrb.build(strict_excel_mode: true) { |w| w.sheet("S1") { |s| s.row [good_str] } }
      end
    end
  end

  # Generator for complex passwords including ASCII, symbols, spaces, and UTF-8 multibyte strings
  def password_generator
    Pbt.tuple(
      Pbt.choose(0..4),
      Pbt.printable_ascii_string(min: 1, max: 35).filter { |s| !s.include?("\x00") && !s.empty? },
      Pbt.alphanumeric_string(min: 1, max: 25)
    )
  end

  def resolve_password(pw_tuple)
    case pw_tuple[0]
    when 0 then pw_tuple[1]
    when 1 then pw_tuple[2]
    when 2 then "パスワード🔑日本語_#{pw_tuple[2]}"
    when 3 then "P@ss w0rd!#%&*()_+=~`_#{pw_tuple[2]}"
    when 4 then "🔒Secret_Unicode_Pass_2026!_#{pw_tuple[1]}"
    end
  end

  test "property-based test: encryption roundtrip with random passwords and encryption modes" do
    Pbt.assert(num_runs: 30) do
      Pbt.property(
        password_generator,
        Pbt.choose(0..1),
        Pbt.array(Pbt.tuple(Pbt.alphanumeric_string(min: 1, max: 15), Pbt.integer(min: -10_000, max: 10_000), Pbt.boolean), min: 1, max: 4)
      ) do |pw_tuple, mode_idx, rows|
        password = resolve_password(pw_tuple)
        mode = mode_idx.zero? ? nil : :agile # nil is the default, Standard

        Dir.mktmpdir do |dir|
          xlsx_path = File.join(dir, "pbt_encrypted.xlsx")

          # 1. Write encrypted file (Streaming write)
          Xlsxrb.write(xlsx_path, password: password, encryption_mode: mode) do |wb|
            wb.sheet("PbtSheet") do |s|
              rows.each { |row| s.row(row) }
            end
          end

          assert_true File.exist?(xlsx_path)
          assert_true Xlsxrb::Ooxml::Crypto.encrypted?(File.binread(xlsx_path))

          # 2. Read with correct password (Streaming read)
          read_rows = []
          Xlsxrb.read(xlsx_path, password: password) do |sheet|
            sheet.each_row { |r| read_rows << r.cells.map(&:value) }
          end

          assert_equal rows.size, read_rows.size
          rows.each_with_index do |expected, idx|
            assert_equal expected[0], read_rows[idx][0]
            assert_equal expected[1], read_rows[idx][1]
            assert_equal expected[2], read_rows[idx][2]
          end

          # 3. Read with correct password (In-memory load)
          wb = Xlsxrb.load(xlsx_path, password: password)
          assert_equal 1, wb.sheets.size
          assert_equal "PbtSheet", wb.sheets[0].name

          # 4. Invariant: Reading without password MUST raise EncryptedFileError
          assert_raise(Xlsxrb::EncryptedFileError) do
            Xlsxrb.read(xlsx_path)
          end

          # 5. Invariant: Reading with wrong password MUST raise InvalidPasswordError
          wrong_password = "#{password}_wrong"
          assert_raise(Xlsxrb::InvalidPasswordError) do
            Xlsxrb.read(xlsx_path, password: wrong_password)
          end
        end
      end
    end
  end

  test "property-based test: encryption tampering and corrupted byte resilience" do
    Pbt.assert(num_runs: 25) do
      Pbt.property(
        password_generator,
        Pbt.integer(min: 0, max: 1000)
      ) do |pw_tuple, tamper_offset_factor|
        password = resolve_password(pw_tuple)

        Dir.mktmpdir do |dir|
          xlsx_path = File.join(dir, "pbt_tamper.xlsx")

          Xlsxrb.write(xlsx_path, password: password) do |wb|
            wb.sheet("Secure") { |s| s.row(["Sensitive", 9999]) }
          end

          raw_bytes = File.binread(xlsx_path)
          offset = (tamper_offset_factor * 17) % raw_bytes.bytesize

          # Tamper 1 to 4 bytes at random offset
          corrupted = raw_bytes.dup
          corrupted.setbyte(offset, corrupted.getbyte(offset) ^ 0xFF)

          # Invariant: Tampered payload MUST either raise an expected safe error or load safely without crashing
          begin
            wb = Xlsxrb.load(StringIO.new(corrupted), password: password)
            assert_instance_of Xlsxrb::Elements::Workbook, wb
          rescue Xlsxrb::Error, ArgumentError, OpenSSL::OpenSSLError, Zlib::Error, REXML::ParseException => e
            # Expected graceful failure on corrupted bytes with known exception types
            assert_not_nil e.message
          end
        end
      end
    end
  end

  test "property-based test: streaming vs in-memory encryption equivalence" do
    Pbt.assert(num_runs: 20) do
      Pbt.property(
        password_generator,
        Pbt.array(Pbt.tuple(Pbt.printable_ascii_string(min: 1, max: 10), Pbt.integer), min: 1, max: 5)
      ) do |pw_tuple, rows|
        password = resolve_password(pw_tuple)

        Dir.mktmpdir do |dir|
          streaming_path = File.join(dir, "streaming.xlsx")
          in_memory_path = File.join(dir, "in_memory.xlsx")

          # 1. Streaming write
          Xlsxrb.write(streaming_path, password: password) do |wb|
            wb.sheet("Data") do |s|
              rows.each { |r| s.row(r) }
            end
          end

          # 2. In-memory write
          wb = Xlsxrb.build do |b|
            b.sheet("Data") do |s|
              rows.each { |r| s.row(r) }
            end
          end
          Xlsxrb.write(in_memory_path, wb, password: password)

          # Invariant: Both methods produce equivalent decrypted data
          streaming_data = Xlsxrb.load(streaming_path, password: password).sheets[0].map(&:to_a)
          in_memory_data = Xlsxrb.load(in_memory_path, password: password).sheets[0].map(&:to_a)

          assert_equal streaming_data, in_memory_data
        end
      end
    end
  end

  # --- PBT for Pure Functions ---

  test "XmlBuilder escape and unescape roundtrip property" do
    Pbt.assert(num_runs: 100) do
      Pbt.property(Pbt.ascii_string) do |str|
        escaped = Xlsxrb::Ooxml::XmlBuilder.escape(str)
        unescaped = Xlsxrb::Ooxml::XmlBuilder.unescape(escaped)
        # escape removes the control characters that XML 1.0 cannot hold.
        assert_equal str.delete("\x00-\x08\x0B\x0C\x0E-\x1F"), unescaped
      end
    end
  end

  test "DslHelpers.absolute_range idempotency property" do
    col_idx_gen = Pbt.integer(min: 0, max: 25)
    row_generator = Pbt.integer(min: 1, max: 1_048_576)

    Pbt.assert(num_runs: 100) do
      Pbt.property(col_idx_gen, row_generator, col_idx_gen, row_generator) do |c1_idx, r1, c2_idx, r2|
        c1 = ("A".ord + c1_idx).chr
        c2 = ("A".ord + c2_idx).chr
        range = "#{c1}#{r1}:#{c2}#{r2}"
        once = Xlsxrb::DslHelpers.absolute_range(range)
        twice = Xlsxrb::DslHelpers.absolute_range(once)
        assert_equal once, twice
        assert_equal "$#{c1}$#{r1}:$#{c2}$#{r2}", once
      end
    end
  end

  test "DslHelpers.normalize_row_values length invariance property" do
    pairs_gen = Pbt.array(Pbt.tuple(Pbt.choose(0..30), Pbt.integer), max: 10)

    Pbt.assert(num_runs: 50) do
      Pbt.property(pairs_gen) do |pairs|
        h = pairs.to_h
        normalized = Xlsxrb::DslHelpers.normalize_row_values(h)
        expected_size = h.empty? ? 0 : h.keys.max + 1
        assert_equal expected_size, normalized.size
        h.each do |k, v|
          assert_equal v, normalized[k]
        end
      end
    end
  end

  # --- Stateful Property-Based Tests (Pbt.stateful v1) ---

  DualWorkbookState = Data.define(
    :sheets,
    :active_sheet_ref,
    :inactive_sheet_refs,
    :active_sheet_name,
    :styles,
    :row_cursors,
    :sheet_rows,
    :sheet_columns,
    :auto_filters,
    :merges,
    :comments
  )

  SheetPair = Data.define(:name, :builder_sheet, :stream_proxy)

  # Drives a live WorkbookBuilder (via Fiber) and StreamWriter in lockstep while
  # recording committed operations for non-destructive serialized snapshots.
  class DualWorkbookSut
    attr_reader :builder, :stream_writer

    def initialize
      @builder = Xlsxrb::WorkbookBuilder.new
      @stream_io = StringIO.new
      @stream_writer = Xlsxrb::StreamWriter.new(@stream_io)
      @active_fiber = nil
      @recorded_ops = []
    end

    def add_sheet(name)
      close_active_fiber
      builder_sheet = nil
      fiber = Fiber.new do
        @builder.sheet(name) do |sheet|
          builder_sheet = sheet
          Fiber.yield
        end
      end
      fiber.resume
      @active_fiber = fiber

      @stream_writer.sheet(name)
      stream_proxy = @stream_writer[]
      @recorded_ops << [:add_sheet, name]
      SheetPair.new(name: name, builder_sheet: builder_sheet, stream_proxy: stream_proxy)
    end

    def define_workbook_style(name, **opts)
      @builder.style(name, **opts)
      @stream_writer.style(name, **opts)
      @recorded_ops << [:define_workbook_style, name, opts]
      nil
    end

    def define_sheet_style(pair, name, **opts)
      pair.builder_sheet.style(name, **opts)
      pair.stream_proxy.style(name, **opts)
      @recorded_ops << [:define_sheet_style, name, opts]
      nil
    end

    def set_column_style(pair, col_idx, width:, style:)
      pair.builder_sheet.column(col_idx, width: width, style: style)
      pair.stream_proxy.column(col_idx, width: width, style: style)
      @recorded_ops << [:set_column_style, col_idx, width, style]
      nil
    end

    def append_row(pair, values, cell_styles:, row_style:)
      b_ret = pair.builder_sheet.row(values, cell_styles: cell_styles, row_style: row_style)
      s_ret = pair.stream_proxy.row(values, cell_styles: cell_styles, row_style: row_style)
      @recorded_ops << [:append_row, values, cell_styles, row_style]
      [b_ret, s_ret]
    end

    def chain_row(pair, values)
      b_ret = (pair.builder_sheet << values)
      s_ret = (pair.stream_proxy << values)
      @recorded_ops << [:chain_row, values]
      [b_ret.equal?(pair.builder_sheet), s_ret.equal?(pair.stream_proxy)]
    end

    def skip_rows(pair, count)
      b_ret = pair.builder_sheet.skip_rows(count)
      s_ret = pair.stream_proxy.skip_rows(count)
      @recorded_ops << [:skip_rows, count]
      [b_ret, s_ret]
    end

    def group_rows_block(pair, values)
      b_ret = pair.builder_sheet.group_rows(level: 1, collapsed: false) { |sh| sh.row(values) }
      s_ret = pair.stream_proxy.group_rows(level: 1, collapsed: false) { |sh| sh.row(values) }
      @recorded_ops << [:group_rows_block, values]
      [b_ret, s_ret]
    end

    def set_sheet_features(pair, auto_filter:, merge_range:, comment_ref:, comment_text:)
      pair.builder_sheet.auto_filter = auto_filter
      pair.stream_proxy.auto_filter = auto_filter
      if merge_range
        pair.builder_sheet.merge(merge_range)
        pair.stream_proxy.merge(merge_range)
      end
      if comment_ref
        pair.builder_sheet.comment(comment_ref, comment_text)
        pair.stream_proxy.comment(comment_ref, comment_text)
      end
      @recorded_ops << [:set_sheet_features, auto_filter, merge_range, comment_ref, comment_text]
      nil
    end

    def snapshot_bytes
      snap_builder = Xlsxrb::WorkbookBuilder.new
      snap_io = StringIO.new
      snap_stream = Xlsxrb::StreamWriter.new(snap_io)
      snap_fiber = nil
      b_sheet = nil
      s_proxy = nil

      begin
        @recorded_ops.each do |op|
          case op[0]
          when :add_sheet
            snap_fiber.resume if snap_fiber&.alive?
            snap_fiber = Fiber.new do
              snap_builder.sheet(op[1]) do |sheet|
                b_sheet = sheet
                Fiber.yield
              end
            end
            snap_fiber.resume
            snap_stream.sheet(op[1])
            s_proxy = snap_stream[]
          when :define_workbook_style
            snap_builder.style(op[1], **op[2])
            snap_stream.style(op[1], **op[2])
          when :define_sheet_style
            b_sheet.style(op[1], **op[2])
            s_proxy.style(op[1], **op[2])
          when :set_column_style
            b_sheet.column(op[1], width: op[2], style: op[3])
            s_proxy.column(op[1], width: op[2], style: op[3])
          when :append_row
            b_sheet.row(op[1], cell_styles: op[2], row_style: op[3])
            s_proxy.row(op[1], cell_styles: op[2], row_style: op[3])
          when :chain_row
            b_sheet << op[1]
            s_proxy << op[1]
          when :skip_rows
            b_sheet.skip_rows(op[1])
            s_proxy.skip_rows(op[1])
          when :group_rows_block
            b_sheet.group_rows(level: 1, collapsed: false) { |sh| sh.row(op[1]) }
            s_proxy.group_rows(level: 1, collapsed: false) { |sh| sh.row(op[1]) }
          when :set_sheet_features
            b_sheet.auto_filter = op[1]
            s_proxy.auto_filter = op[1]
            if op[2]
              b_sheet.merge(op[2])
              s_proxy.merge(op[2])
            end
            if op[3]
              b_sheet.comment(op[3], op[4])
              s_proxy.comment(op[3], op[4])
            end
          end
        end

        snap_fiber.resume if snap_fiber&.alive?
        snap_stream.close
        [Xlsxrb.write(snap_builder.build), snap_io.string.b]
      ensure
        if snap_fiber&.alive?
          begin
            snap_fiber.resume
          rescue StandardError
            nil
          end
        end
        snap_stream.cleanup!
      end
    end

    def cleanup!
      close_active_fiber
      @stream_writer.cleanup!
    end

    private

    def close_active_fiber
      return unless @active_fiber&.alive?

      @active_fiber.resume
    rescue StandardError
      nil
    ensure
      @active_fiber = nil
    end
  end

  def resolve_expected_cell_style(value, cell_style, row_style, col_info, styles_map)
    date_or_time = value.is_a?(Date) || value.is_a?(Time)
    style_name = cell_style || row_style
    style_def = if style_name
                  styles_map.fetch(style_name)
                elsif col_info
                  date_or_time ? styles_map.fetch(col_info[:style_name]) : col_info[:style_def]
                end

    bold = style_def&.dig(:bold) || false
    italic = style_def&.dig(:italic) || false
    fmt = style_def&.dig(:number_format)
    if fmt.nil?
      fmt = if value.is_a?(Time)
              "yyyy-mm-dd hh:mm:ss"
            elsif value.is_a?(Date)
              "yyyy-mm-dd"
            else
              "General"
            end
    end
    [bold, italic, fmt]
  end

  def extract_loaded_cell_style(styles, cell)
    xf = cell.style_index ? styles.cell_xf(cell.style_index) : nil
    font = xf ? styles.fonts[xf.font_id.to_i] : nil
    fmt = xf ? (styles.number_format_code(xf.num_fmt_id.to_i) || "General") : "General"
    [font&.bold? || false, font&.italic? || false, fmt]
  end

  test "stateful PBT: dual-mode workbook builder and stream writer lockstep equivalence" do
    style_name_arb = Pbt.constant_from(:s1, :s2, :s3)
    num_fmt_arb = Pbt.constant_from(nil, "#,##0.00", "0.0%", "yyyy-mm-dd")
    cell_val_arb = Pbt.one_of(
      Pbt.integer(min: -1000, max: 1000),
      Pbt.alphanumeric_string(min: 1, max: 12),
      Pbt.boolean,
      Pbt.constant(Date.new(2026, 4, 1)),
      Pbt.constant(Time.utc(2026, 4, 1, 9, 30, 0))
    )

    Pbt.assert(num_runs: 15) do
      Pbt.stateful(max_steps: 12) do |m|
        m.initial_state do
          DualWorkbookState.new(
            sheets: [],
            active_sheet_ref: nil,
            inactive_sheet_refs: [],
            active_sheet_name: nil,
            styles: {},
            row_cursors: {},
            sheet_rows: {},
            sheet_columns: {},
            auto_filters: {},
            merges: {},
            comments: {}
          )
        end

        m.sut { DualWorkbookSut.new }
        m.teardown(&:cleanup!)

        m.command :add_sheet, weight: 2 do |c|
          c.guard { |s| s.sheets.size < 3 }
          c.arguments Pbt.constant_from("Sheet1", "Sheet2", "Sheet3")
          c.precondition { |s, name| !s.sheets.include?(name) }
          c.next_state do |s, name, ref|
            inactive = s.active_sheet_ref ? [*s.inactive_sheet_refs, s.active_sheet_ref] : s.inactive_sheet_refs
            s.with(
              sheets: [*s.sheets, name],
              active_sheet_ref: ref,
              inactive_sheet_refs: inactive,
              active_sheet_name: name,
              row_cursors: s.row_cursors.merge(name => 0),
              sheet_rows: s.sheet_rows.merge(name => []),
              sheet_columns: s.sheet_columns.merge(name => {}),
              auto_filters: s.auto_filters.merge(name => nil),
              merges: s.merges.merge(name => []),
              comments: s.comments.merge(name => [])
            )
          end
          c.run { |sut, name| sut.add_sheet(name) }
        end

        m.command :define_workbook_style, weight: 2 do |c|
          c.arguments(name: style_name_arb, bold: Pbt.boolean, italic: Pbt.boolean, number_format: num_fmt_arb)
          c.next_state do |s, args|
            style_def = { bold: args[:bold], italic: args[:italic], number_format: args[:number_format] }
            s.with(styles: s.styles.merge(args[:name] => style_def))
          end
          c.run do |sut, args|
            sut.define_workbook_style(args[:name], bold: args[:bold], italic: args[:italic], number_format: args[:number_format])
          end
        end

        m.command :define_sheet_style, weight: 2 do |c|
          c.guard { |s| !s.active_sheet_ref.nil? }
          c.arguments do |s|
            Pbt.fixed_hash(
              sheet: Pbt.constant(s.active_sheet_ref),
              name: style_name_arb,
              bold: Pbt.boolean,
              italic: Pbt.boolean,
              number_format: num_fmt_arb
            )
          end
          c.next_state do |s, args|
            style_def = { bold: args[:bold], italic: args[:italic], number_format: args[:number_format] }
            s.with(styles: s.styles.merge(args[:name] => style_def))
          end
          c.run do |sut, args|
            sut.define_sheet_style(args[:sheet], args[:name], bold: args[:bold], italic: args[:italic], number_format: args[:number_format])
          end
        end

        m.command :set_column_style, weight: 2 do |c|
          c.guard { |s| !s.active_sheet_ref.nil? && !s.styles.empty? && s.row_cursors.fetch(s.active_sheet_name).zero? }
          c.arguments do |s|
            Pbt.fixed_hash(
              sheet: Pbt.constant(s.active_sheet_ref),
              col_idx: Pbt.constant_from(0, 1),
              width: Pbt.integer(min: 10, max: 30),
              style: Pbt.constant_from(*s.styles.keys)
            )
          end
          c.next_state do |s, args|
            sname = s.active_sheet_name
            col_entry = { style_name: args[:style], style_def: s.styles.fetch(args[:style]) }
            updated_cols = s.sheet_columns.fetch(sname).merge(args[:col_idx] => col_entry)
            s.with(sheet_columns: s.sheet_columns.merge(sname => updated_cols))
          end
          c.run do |sut, args|
            sut.set_column_style(args[:sheet], args[:col_idx], width: args[:width], style: args[:style])
          end
        end

        m.command :append_row, weight: 5 do |c|
          c.guard { |s| !s.active_sheet_ref.nil? }
          c.arguments do |s|
            style_choice_arb = Pbt.constant_from(nil, *s.styles.keys)
            Pbt.fixed_hash(
              sheet: Pbt.constant(s.active_sheet_ref),
              v0: cell_val_arb,
              v1: cell_val_arb,
              cs0: style_choice_arb,
              cs1: style_choice_arb,
              row_style: style_choice_arb
            )
          end
          c.next_state do |s, args|
            sname = s.active_sheet_name
            row_idx = s.row_cursors.fetch(sname)
            cols = s.sheet_columns.fetch(sname)
            expected_cells = [
              { value: args[:v0], style: resolve_expected_cell_style(args[:v0], args[:cs0], args[:row_style], cols[0], s.styles) },
              { value: args[:v1], style: resolve_expected_cell_style(args[:v1], args[:cs1], args[:row_style], cols[1], s.styles) }
            ]
            row_record = { row_index: row_idx, cells: expected_cells }
            s.with(
              row_cursors: s.row_cursors.merge(sname => row_idx + 1),
              sheet_rows: s.sheet_rows.merge(sname => [*s.sheet_rows.fetch(sname), row_record])
            )
          end
          c.run do |sut, args|
            cell_styles = args[:cs0].nil? && args[:cs1].nil? ? nil : [args[:cs0], args[:cs1]]
            sut.append_row(args[:sheet], [args[:v0], args[:v1]], cell_styles: cell_styles, row_style: args[:row_style])
          end
          c.verify { |ctx| assert_equal [nil, nil], ctx.result }
        end

        m.command :chain_row, weight: 2 do |c|
          c.guard { |s| !s.active_sheet_ref.nil? }
          c.arguments do |s|
            Pbt.fixed_hash(sheet: Pbt.constant(s.active_sheet_ref), v0: cell_val_arb, v1: cell_val_arb)
          end
          c.next_state do |s, args|
            sname = s.active_sheet_name
            row_idx = s.row_cursors.fetch(sname)
            cols = s.sheet_columns.fetch(sname)
            expected_cells = [
              { value: args[:v0], style: resolve_expected_cell_style(args[:v0], nil, nil, cols[0], s.styles) },
              { value: args[:v1], style: resolve_expected_cell_style(args[:v1], nil, nil, cols[1], s.styles) }
            ]
            row_record = { row_index: row_idx, cells: expected_cells }
            s.with(
              row_cursors: s.row_cursors.merge(sname => row_idx + 1),
              sheet_rows: s.sheet_rows.merge(sname => [*s.sheet_rows.fetch(sname), row_record])
            )
          end
          c.run { |sut, args| sut.chain_row(args[:sheet], [args[:v0], args[:v1]]) }
          c.verify { |ctx| assert_equal [true, true], ctx.result }
        end

        m.command :skip_rows, weight: 2 do |c|
          c.guard { |s| !s.active_sheet_ref.nil? }
          c.arguments { |s| Pbt.fixed_hash(sheet: Pbt.constant(s.active_sheet_ref), count: Pbt.integer(min: 1, max: 2)) }
          c.next_state do |s, args|
            sname = s.active_sheet_name
            s.with(row_cursors: s.row_cursors.merge(sname => s.row_cursors.fetch(sname) + args[:count]))
          end
          c.run { |sut, args| sut.skip_rows(args[:sheet], args[:count]) }
          c.verify { |ctx| assert_equal [nil, nil], ctx.result }
        end

        m.command :group_rows_block, weight: 2 do |c|
          c.guard { |s| !s.active_sheet_ref.nil? }
          c.arguments { |s| Pbt.fixed_hash(sheet: Pbt.constant(s.active_sheet_ref), v0: cell_val_arb, v1: cell_val_arb) }
          c.next_state do |s, args|
            sname = s.active_sheet_name
            row_idx = s.row_cursors.fetch(sname)
            cols = s.sheet_columns.fetch(sname)
            expected_cells = [
              { value: args[:v0], style: resolve_expected_cell_style(args[:v0], nil, nil, cols[0], s.styles) },
              { value: args[:v1], style: resolve_expected_cell_style(args[:v1], nil, nil, cols[1], s.styles) }
            ]
            row_record = { row_index: row_idx, cells: expected_cells }
            s.with(
              row_cursors: s.row_cursors.merge(sname => row_idx + 1),
              sheet_rows: s.sheet_rows.merge(sname => [*s.sheet_rows.fetch(sname), row_record])
            )
          end
          c.run { |sut, args| sut.group_rows_block(args[:sheet], [args[:v0], args[:v1]]) }
          c.verify { |ctx| assert_equal [nil, nil], ctx.result }
        end

        m.command :set_sheet_features, weight: 2 do |c|
          c.guard { |s| !s.active_sheet_ref.nil? }
          c.arguments do |s|
            existing_comments = s.comments.fetch(s.active_sheet_name).map { |entry| entry[:ref] }
            available_refs = %w[A1 B1 A2 B2] - existing_comments
            comment_ref_arb = available_refs.empty? ? Pbt.nil : Pbt.constant_from(*available_refs)
            Pbt.fixed_hash(
              sheet: Pbt.constant(s.active_sheet_ref),
              comment_ref: comment_ref_arb,
              comment_text: Pbt.alphanumeric_string(min: 1, max: 12)
            )
          end
          c.next_state do |s, args|
            sname = s.active_sheet_name
            existing_merges = s.merges.fetch(sname)
            new_merges = existing_merges.include?("A1:B1") ? existing_merges : [*existing_merges, "A1:B1"]
            new_comments = if args[:comment_ref]
                             [*s.comments.fetch(sname), { ref: args[:comment_ref], text: args[:comment_text] }]
                           else
                             s.comments.fetch(sname)
                           end
            s.with(
              auto_filters: s.auto_filters.merge(sname => "A1:B2"),
              merges: s.merges.merge(sname => new_merges),
              comments: s.comments.merge(sname => new_comments)
            )
          end
          c.run do |sut, args|
            sut.set_sheet_features(
              args[:sheet],
              auto_filter: "A1:B2",
              merge_range: "A1:B1",
              comment_ref: args[:comment_ref],
              comment_text: args[:comment_text]
            )
          end
        end

        m.command :write_inactive_sheet, weight: 2 do |c|
          c.guard { |s| !s.inactive_sheet_refs.empty? }
          c.arguments { |s| Pbt.constant_from(*s.inactive_sheet_refs) }
          c.run do |_sut, inactive_pair|
            b_err = begin
              inactive_pair.builder_sheet.row([1])
              nil
            rescue Xlsxrb::Error => e
              e
            end
            s_err = begin
              inactive_pair.stream_proxy.row([1])
              nil
            rescue Xlsxrb::Error => e
              e
            end
            [b_err, s_err]
          end
          c.verify do |ctx|
            assert_instance_of Xlsxrb::Error, ctx.result[0]
            assert_instance_of Xlsxrb::Error, ctx.result[1]
          end
        end

        m.invariant(:live_and_serialized_parity) do |state, sut|
          assert_equal state.sheets, sut.builder.sheet_names
          assert_equal state.sheets, sut.stream_writer.sheet_names
          next if state.sheets.empty?

          built_bytes, streamed_bytes = sut.snapshot_bytes
          assert_empty ExcelRules.check(built_bytes).map(&:to_s)
          assert_empty ExcelRules.check(streamed_bytes).map(&:to_s)

          built_wb = Xlsxrb.load(StringIO.new(built_bytes))
          streamed_wb = Xlsxrb.load(StringIO.new(streamed_bytes))

          assert_equal state.sheets, built_wb.sheets.map(&:name)
          assert_equal state.sheets, streamed_wb.sheets.map(&:name)

          state.sheets.each_with_index do |sname, idx|
            b_sheet = built_wb.sheet(idx)
            s_sheet = streamed_wb.sheet(idx)
            expected_rows = state.sheet_rows.fetch(sname)

            assert_equal expected_rows.map { |r| r[:row_index] }, b_sheet.rows.map(&:index)
            assert_equal expected_rows.map { |r| r[:row_index] }, s_sheet.rows.map(&:index)
            assert_equal state.auto_filters.fetch(sname), b_sheet.auto_filter
            assert_equal state.auto_filters.fetch(sname), s_sheet.auto_filter
            assert_equal state.merges.fetch(sname), b_sheet.merged_cells
            assert_equal state.merges.fetch(sname), s_sheet.merged_cells
            assert_equal(state.comments.fetch(sname).map { |c| [c[:ref], c[:text]] },
                         b_sheet.comments.map { |c| [c.ref, c.text] })
            assert_equal(state.comments.fetch(sname).map { |c| [c[:ref], c[:text]] },
                         s_sheet.comments.map { |c| [c.ref, c.text] })

            expected_rows.each_with_index do |exp_row, r_idx|
              b_row = b_sheet.rows[r_idx]
              s_row = s_sheet.rows[r_idx]
              assert_equal exp_row[:cells].size, b_row.cells.size
              assert_equal exp_row[:cells].size, s_row.cells.size

              exp_row[:cells].each_with_index do |exp_cell, c_idx|
                b_cell = b_row.cells[c_idx]
                s_cell = s_row.cells[c_idx]
                exp_val = exp_cell[:value]

                case exp_val
                when Date
                  assert_equal exp_val, b_cell.to_date
                  assert_equal exp_val, s_cell.to_date
                when Time
                  assert_in_delta exp_val.to_i, b_cell.to_time.to_i, 1
                  assert_in_delta exp_val.to_i, s_cell.to_time.to_i, 1
                else
                  assert_equal exp_val, b_cell.value
                  assert_equal exp_val, s_cell.value
                end

                b_style = extract_loaded_cell_style(built_wb.styles, b_cell)
                s_style = extract_loaded_cell_style(streamed_wb.styles, s_cell)
                assert_equal exp_cell[:style], b_style
                assert_equal exp_cell[:style], s_style
              end
            end
          end
        end
      end
    end
  end

  TemporalStyleState = Data.define(:clock, :style_history, :rows)

  # Drives a live WorkbookBuilder sheet via Fiber and replays ops for non-destructive snapshot checks.
  class TemporalStyleSut
    def initialize
      @builder = Xlsxrb::WorkbookBuilder.new
      @sheet = nil
      @fiber = Fiber.new do
        @builder.sheet("S1") do |sheet|
          @sheet = sheet
          Fiber.yield
        end
      end
      @fiber.resume
      @ops = []
    end

    def write_row(value, style_name)
      @sheet.row([value], cell_styles: [style_name])
      @ops << [:row, value, style_name]
      nil
    end

    def define_style(name, **opts)
      @sheet.style(name, **opts)
      @ops << [:style, name, opts]
      nil
    end

    def snapshot_bytes
      wb = Xlsxrb::WorkbookBuilder.new
      wb.sheet("S1") do |sheet|
        @ops.each do |op|
          if op[0] == :row
            sheet.row([op[1]], cell_styles: [op[2]])
          else
            sheet.style(op[1], **op[2])
          end
        end
      end
      Xlsxrb.write(wb.build)
    end

    def cleanup!
      return unless @fiber&.alive?

      @fiber.resume
    rescue StandardError
      nil
    ensure
      @fiber = nil
    end
  end

  test "stateful PBT: WorkbookBuilder pre-definition style reference and StyleLog temporal resolution" do
    style_name_arb = Pbt.constant_from(:s1, :s2)
    num_fmt_arb = Pbt.constant_from(nil, "0.00", "yyyy/mm/dd")
    val_arb = Pbt.one_of(
      Pbt.integer(min: -100, max: 100),
      Pbt.alphanumeric_string(min: 1, max: 8),
      Pbt.constant(Date.new(2026, 5, 1)),
      Pbt.constant(Time.utc(2026, 5, 1, 12, 0, 0))
    )

    Pbt.assert(num_runs: 25) do
      Pbt.stateful(max_steps: 15) do |m|
        m.initial_state { TemporalStyleState.new(clock: 0, style_history: {}, rows: []) }
        m.sut { TemporalStyleSut.new }
        m.teardown(&:cleanup!)

        m.command :write_row, weight: 4 do |c|
          c.arguments(value: val_arb, style_name: style_name_arb)
          c.next_state do |s, args|
            entry = { row_index: s.rows.size, value: args[:value], style_name: args[:style_name], clock: s.clock }
            s.with(rows: [*s.rows, entry])
          end
          c.run { |sut, args| sut.write_row(args[:value], args[:style_name]) }
        end

        m.command :define_style, weight: 3 do |c|
          c.arguments(name: style_name_arb, bold: Pbt.boolean, italic: Pbt.boolean, number_format: num_fmt_arb)
          c.next_state do |s, args|
            new_clock = s.clock + 1
            entry = { clock: new_clock, bold: args[:bold], italic: args[:italic], number_format: args[:number_format] }
            history = s.style_history[args[:name]] || []
            s.with(clock: new_clock, style_history: s.style_history.merge(args[:name] => [*history, entry]))
          end
          c.run do |sut, args|
            sut.define_style(args[:name], bold: args[:bold], italic: args[:italic], number_format: args[:number_format])
          end
        end

        m.invariant(:temporal_style_resolution) do |state, sut|
          next if state.rows.empty?

          referenced = state.rows.map { |r| r[:style_name] }.uniq
          next unless referenced.all? { |name| state.style_history.key?(name) }

          bytes = sut.snapshot_bytes
          assert_empty ExcelRules.check(bytes).map(&:to_s)

          wb = Xlsxrb.load(StringIO.new(bytes))
          loaded_rows = wb.sheet(0).rows
          assert_equal state.rows.size, loaded_rows.size

          state.rows.each_with_index do |exp_row, idx|
            defs = state.style_history.fetch(exp_row[:style_name])
            prior_defs = defs.select { |d| d[:clock] <= exp_row[:clock] }
            resolved_def = prior_defs.empty? ? defs.first : prior_defs.last

            exp_fmt = resolved_def[:number_format]
            if exp_fmt.nil?
              exp_fmt = if exp_row[:value].is_a?(Time)
                          "yyyy-mm-dd hh:mm:ss"
                        elsif exp_row[:value].is_a?(Date)
                          "yyyy-mm-dd"
                        else
                          "General"
                        end
            end
            expected_style = [resolved_def[:bold] ? true : false, resolved_def[:italic] ? true : false, exp_fmt]
            actual_style = extract_loaded_cell_style(wb.styles, loaded_rows[idx].cells[0])
            assert_equal expected_style, actual_style
          end
        end
      end
    end
  end

  ZipModelState = Data.define(
    :seekable,
    :compression_level,
    :zip64_limit,
    :in_stream,
    :current_entry,
    :current_chunks,
    :entries,
    :next_id
  )

  # Minimal write-only sink without seek/tell/pos to exercise unseekable ZipWriter paths.
  class UnseekableZipSink
    attr_reader :buffer

    def initialize
      @buffer = +"".b
    end

    def write(data)
      @buffer << data.b
      data.bytesize
    end
  end

  # Drives a live Ooxml::ZipWriter while recording operations for non-destructive archive checks.
  class ZipWriterSut
    def initialize(initial_state)
      @seekable = initial_state.seekable
      @compression_level = initial_state.compression_level
      @zip64_limit = initial_state.zip64_limit
      @io = @seekable ? StringIO.new : UnseekableZipSink.new
      @writer = Xlsxrb::Ooxml::ZipWriter.new(@io, compression_level: @compression_level)
      @writer.instance_variable_set(:@zip64_limit, @zip64_limit)
      @next_id = 0
      @recorded_ops = []
    end

    def add_entry(payload)
      name = "entry_#{@next_id}.xml"
      @next_id += 1
      @writer.add_entry(name, payload)
      @recorded_ops << [:add_entry, name, payload]
      nil
    end

    def add_binary_entry(payload)
      name = "bin_#{@next_id}.bin"
      @next_id += 1
      @writer.add_binary_entry(name, payload)
      @recorded_ops << [:add_binary_entry, name, payload]
      nil
    end

    def start_stream_entry
      name = "stream_#{@next_id}.xml"
      @next_id += 1
      @writer.start_entry(name)
      @recorded_ops << [:start_entry, name]
      nil
    end

    def write_stream_data(chunk)
      @writer.write_data(chunk)
      @recorded_ops << [:write_data, chunk]
      nil
    end

    def finish_stream_entry
      @writer.finish_entry
      @recorded_ops << [:finish_entry]
      nil
    end

    def write_data_without_entry(chunk)
      @writer.write_data(chunk)
    end

    def snapshot_bytes
      snap_io = @seekable ? StringIO.new : UnseekableZipSink.new
      snap_writer = Xlsxrb::Ooxml::ZipWriter.new(snap_io, compression_level: @compression_level)
      snap_writer.instance_variable_set(:@zip64_limit, @zip64_limit)

      @recorded_ops.each do |op|
        case op[0]
        when :add_entry
          snap_writer.add_entry(op[1], op[2])
        when :add_binary_entry
          snap_writer.add_binary_entry(op[1], op[2])
        when :start_entry
          snap_writer.start_entry(op[1])
        when :write_data
          snap_writer.write_data(op[1])
        when :finish_entry
          snap_writer.finish_entry
        end
      end

      snap_writer.close
      @seekable ? snap_io.string.b : snap_io.buffer
    end

    def cleanup!
      @writer.close
    rescue StandardError
      nil
    end
  end

  test "stateful PBT: Ooxml::ZipWriter streaming, raw entries, and ZIP64 promotion state machine" do
    initial_arb = Pbt.fixed_hash(
      seekable: Pbt.boolean,
      compression_level: Pbt.constant_from(0, 1, 6),
      zip64_limit: Pbt.integer(min: 64, max: 256)
    ).map(
      lambda do |h|
        ZipModelState.new(
          seekable: h[:seekable],
          compression_level: h[:compression_level],
          zip64_limit: h[:zip64_limit],
          in_stream: false,
          current_entry: nil,
          current_chunks: [],
          entries: {},
          next_id: 0
        )
      end,
      lambda do |s|
        {
          seekable: s.seekable,
          compression_level: s.compression_level,
          zip64_limit: s.zip64_limit
        }
      end
    )

    binary_arb = Pbt.array(Pbt.integer(min: 0, max: 255), min: 0, max: 300).map(
      ->(bytes) { bytes.pack("C*") },
      :bytes.to_proc
    )

    Pbt.assert(num_runs: 30) do
      Pbt.stateful(max_steps: 15) do |m|
        m.initial_state initial_arb
        m.sut { |init| ZipWriterSut.new(init) }
        m.teardown(&:cleanup!)

        m.command :add_entry, weight: 3 do |c|
          c.guard { |s| !s.in_stream }
          c.arguments Pbt.alphanumeric_string(min: 0, max: 300)
          c.next_state do |s, payload|
            name = "entry_#{s.next_id}.xml"
            s.with(entries: s.entries.merge(name => payload.b), next_id: s.next_id + 1)
          end
          c.run { |sut, payload| sut.add_entry(payload) }
        end

        m.command :add_binary_entry, weight: 3 do |c|
          c.guard { |s| !s.in_stream }
          c.arguments binary_arb
          c.next_state do |s, payload|
            name = "bin_#{s.next_id}.bin"
            s.with(entries: s.entries.merge(name => payload.b), next_id: s.next_id + 1)
          end
          c.run { |sut, payload| sut.add_binary_entry(payload) }
        end

        m.command :start_stream_entry, weight: 3 do |c|
          c.guard { |s| !s.in_stream }
          c.next_state do |s|
            name = "stream_#{s.next_id}.xml"
            s.with(in_stream: true, current_entry: name, current_chunks: [], next_id: s.next_id + 1)
          end
          c.run { |sut, _args| sut.start_stream_entry }
        end

        m.command :write_stream_data, weight: 4 do |c|
          c.guard(&:in_stream)
          c.arguments Pbt.alphanumeric_string(min: 0, max: 200)
          c.next_state { |s, chunk| s.with(current_chunks: [*s.current_chunks, chunk.b]) }
          c.run { |sut, chunk| sut.write_stream_data(chunk) }
        end

        m.command :finish_stream_entry, weight: 3 do |c|
          c.guard(&:in_stream)
          c.next_state do |s|
            payload = s.current_chunks.join
            s.with(
              in_stream: false,
              current_entry: nil,
              current_chunks: [],
              entries: s.entries.merge(s.current_entry => payload)
            )
          end
          c.run { |sut, _args| sut.finish_stream_entry }
        end

        m.command :write_data_without_entry, weight: 1 do |c|
          c.guard { |s| !s.in_stream }
          c.run do |sut|
            sut.write_data_without_entry("x")
            nil
          rescue RuntimeError => e
            e
          end
          c.verify do |ctx|
            assert_instance_of RuntimeError, ctx.result
            assert_match(/No entry started/, ctx.result.message)
          end
        end

        m.invariant(:zip_archive_integrity) do |state, sut|
          next if state.in_stream || state.entries.empty?

          bytes = sut.snapshot_bytes
          assert_empty ExcelRules.duplicate_entries(bytes).map(&:to_s)

          reader = Xlsxrb::Ooxml::ZipReader.new(StringIO.new(bytes))
          assert_equal state.entries.keys, reader.entry_names

          state.entries.each do |name, expected_bytes|
            assert_equal expected_bytes, reader.read_entry(name).b
            assert_equal expected_bytes.bytesize, reader.entry_catalog.fetch(name).uncompressed_size
          end

          eocd_pos = bytes.rindex("PK\x05\x06".b)
          assert_not_nil eocd_pos
          _d, _cdd, _de, _te, eocd_cd_size, eocd_cd_offset, = bytes.byteslice(eocd_pos + 4, 18).unpack("vvvvVVv")
          zip64_eocd_pos = bytes.index("PK\x06\x06".b)
          zip64_locator_pos = bytes.index("PK\x06\x07".b)

          if zip64_eocd_pos
            actual_cd_size, actual_cd_offset = bytes.byteslice(zip64_eocd_pos + 40, 16).unpack("Q<Q<")
            assert_not_nil zip64_locator_pos
            assert_true actual_cd_offset >= state.zip64_limit || actual_cd_size >= state.zip64_limit || state.entries.size >= 0xFFFF
            assert_equal 0xFFFF_FFFF, eocd_cd_size if actual_cd_size >= state.zip64_limit
            assert_equal 0xFFFF_FFFF, eocd_cd_offset if actual_cd_offset >= state.zip64_limit
          else
            assert_nil zip64_locator_pos
            assert_operator eocd_cd_offset, :<, state.zip64_limit
            assert_operator eocd_cd_size, :<, state.zip64_limit
            assert_operator state.entries.size, :<, 0xFFFF
          end
        end
      end
    end
  end

  MultiReaderState = Data.define(
    :row_cursor,
    :rows,
    :columns,
    :auto_filter,
    :merges,
    :comments,
    :hyperlinks,
    :print_area,
    :freeze_row
  )

  # Records sheet authoring operations and serializes a single-sheet XLSX for multi-reader parity checks.
  class MultiReaderSut
    def initialize
      @ops = []
    end

    def append_row(values)
      @ops << [:row, values]
      nil
    end

    def skip_rows(count)
      @ops << [:skip_rows, count]
      nil
    end

    def set_column(col_idx, width:, hidden:, outline_level:)
      @ops << [:column, col_idx, width, hidden, outline_level]
      nil
    end

    def apply_freeze_pane(row)
      @ops << [:freeze_pane, row]
      nil
    end

    def add_merge_and_filter(auto_filter:, merge_range:)
      @ops << [:merge_filter, auto_filter, merge_range]
      nil
    end

    def add_hyperlink_and_comment(ref, url, comment_text)
      @ops << [:link_comment, ref, url, comment_text]
      nil
    end

    def apply_print_area(range)
      @ops << [:print_area, range]
      nil
    end

    def serialize_bytes
      io = StringIO.new
      Xlsxrb.write(io) do |wb|
        wb.sheet("Data") do |sheet|
          @ops.each do |op|
            case op[0]
            when :row
              sheet.row(op[1])
            when :skip_rows
              sheet.skip_rows(op[1])
            when :column
              sheet.column(op[1], width: op[2], hidden: op[3], outline_level: op[4])
            when :freeze_pane
              sheet.freeze_pane(row: op[1], column: 0)
            when :merge_filter
              sheet.auto_filter = op[1]
              sheet.merge(op[2])
            when :link_comment
              sheet.hyperlink(op[1], op[2])
              sheet.comment(op[1], op[3])
            when :print_area
              sheet.print_area = op[1]
            end
          end
        end
      end
      io.string.b
    end
  end

  test "stateful PBT: Xlsxrb.load, Xlsxrb.read (StreamSheet), and StreamSheet#load equivalence across stateful sheet authoring" do
    val_arb = Pbt.one_of(
      Pbt.integer(min: -500, max: 500),
      Pbt.alphanumeric_string(min: 1, max: 10),
      Pbt.boolean,
      Pbt.constant(Date.new(2026, 6, 15))
    )

    Pbt.assert(num_runs: 15) do
      Pbt.stateful(max_steps: 14) do |m|
        m.initial_state do
          MultiReaderState.new(
            row_cursor: 0,
            rows: [],
            columns: {},
            auto_filter: nil,
            merges: [],
            comments: {},
            hyperlinks: {},
            print_area: nil,
            freeze_row: nil
          )
        end

        m.sut { MultiReaderSut.new }

        m.command :append_row, weight: 4 do |c|
          c.arguments(v0: val_arb, v1: val_arb)
          c.next_state do |s, args|
            expected = [args[:v0], args[:v1]].map { |v| v.is_a?(Date) ? Xlsxrb.date_to_serial(v) : v }
            row_record = { row_index: s.row_cursor, values: expected }
            s.with(row_cursor: s.row_cursor + 1, rows: [*s.rows, row_record])
          end
          c.run { |sut, args| sut.append_row([args[:v0], args[:v1]]) }
        end

        m.command :skip_rows, weight: 2 do |c|
          c.arguments Pbt.integer(min: 1, max: 2)
          c.next_state { |s, count| s.with(row_cursor: s.row_cursor + count) }
          c.run { |sut, count| sut.skip_rows(count) }
        end

        m.command :set_column, weight: 2 do |c|
          c.guard { |s| s.row_cursor.zero? }
          c.arguments(
            col_idx: Pbt.constant_from(0, 1),
            width: Pbt.integer(min: 10, max: 30),
            hidden: Pbt.boolean,
            outline_level: Pbt.constant_from(nil, 1, 2)
          )
          c.next_state do |s, args|
            col_spec = { width: args[:width].to_f, hidden: args[:hidden], outline_level: args[:outline_level] }
            s.with(columns: s.columns.merge(args[:col_idx] => col_spec))
          end
          c.run do |sut, args|
            sut.set_column(args[:col_idx], width: args[:width], hidden: args[:hidden], outline_level: args[:outline_level])
          end
        end

        m.command :set_freeze_pane, weight: 1 do |c|
          c.guard { |s| s.row_cursor.zero? }
          c.next_state { |s| s.with(freeze_row: 1) }
          c.run { |sut, _args| sut.apply_freeze_pane(1) }
        end

        m.command :add_merge_and_filter, weight: 2 do |c|
          c.next_state do |s|
            new_merges = s.merges.include?("A1:B1") ? s.merges : [*s.merges, "A1:B1"]
            s.with(auto_filter: "A1:B2", merges: new_merges)
          end
          c.run { |sut, _args| sut.add_merge_and_filter(auto_filter: "A1:B2", merge_range: "A1:B1") }
        end

        m.command :add_hyperlink_and_comment, weight: 2 do |c|
          c.guard { |s| s.comments.size < 3 }
          c.arguments do |s|
            available_refs = %w[A1 B1 A2] - s.comments.keys
            Pbt.fixed_hash(
              ref: Pbt.constant_from(*available_refs),
              comment_text: Pbt.alphanumeric_string(min: 1, max: 12)
            )
          end
          c.next_state do |s, args|
            ref = args[:ref]
            s.with(
              comments: s.comments.merge(ref => args[:comment_text]),
              hyperlinks: s.hyperlinks.merge(ref => "https://example.com/#{ref}")
            )
          end
          c.run do |sut, args|
            ref = args[:ref]
            sut.add_hyperlink_and_comment(ref, "https://example.com/#{ref}", args[:comment_text])
          end
        end

        m.command :set_print_area, weight: 1 do |c|
          c.next_state { |s| s.with(print_area: "A1:B4") }
          c.run { |sut, _args| sut.apply_print_area("A1:B4") }
        end

        m.invariant(:three_way_reader_equivalence) do |state, sut|
          next if state.rows.empty?

          bytes = sut.serialize_bytes
          assert_empty ExcelRules.check(bytes).map(&:to_s)

          loaded_ws = Xlsxrb.load(StringIO.new(bytes)).sheet(0)
          expected_rows = state.rows.map { |r| [r[:row_index], r[:values]] }
          loaded_rows = loaded_ws.rows.map { |r| [r.index, r.cells.map(&:value)] }
          assert_equal expected_rows, loaded_rows
          assert_equal state.auto_filter, loaded_ws.auto_filter
          assert_equal state.merges, loaded_ws.merged_cells
          assert_equal(state.comments.to_a, loaded_ws.comments.map { |c| [c.ref, c.text] })
          assert_equal(state.hyperlinks, loaded_ws.hyperlinks.transform_values(&:url))
          assert_equal state.print_area, loaded_ws.print_area
          assert_equal state.freeze_row, loaded_ws.freeze_pane&.row

          expected_cols = state.columns.sort_by(&:first).map do |idx, col|
            [idx, col[:width], col[:hidden], col[:outline_level]]
          end
          loaded_cols = loaded_ws.columns.map { |c| [c.index, c.width, c.hidden?, c.outline_level] }
          assert_equal expected_cols, loaded_cols

          stream_count = 0
          Xlsxrb.read(StringIO.new(bytes)) do |stream_sheet|
            stream_count += 1
            assert_equal loaded_ws.name, stream_sheet.name
            assert_equal loaded_ws.auto_filter, stream_sheet.auto_filter
            assert_equal loaded_ws.merged_cells, stream_sheet.merged_cells
            assert_equal(loaded_ws.comments.map { |c| [c.ref, c.text] }, stream_sheet.comments.map { |c| [c.ref, c.text] })
            assert_equal(loaded_ws.hyperlinks.transform_values(&:url), stream_sheet.hyperlinks.transform_values(&:url))
            assert_equal loaded_ws.print_area, stream_sheet.print_area
            assert_equal loaded_ws.freeze_pane&.row, stream_sheet.freeze_pane&.row
            assert_equal(loaded_cols, stream_sheet.columns.map { |c| [c.index, c.width, c.hidden?, c.outline_level] })
            assert_equal(loaded_rows, stream_sheet.each_row.map { |r| [r.index, r.values] })

            mat_ws = stream_sheet.load
            assert_equal loaded_ws.name, mat_ws.name
            assert_equal loaded_ws.auto_filter, mat_ws.auto_filter
            assert_equal loaded_ws.merged_cells, mat_ws.merged_cells
            assert_equal(loaded_ws.comments.map { |c| [c.ref, c.text] }, mat_ws.comments.map { |c| [c.ref, c.text] })
            assert_equal(loaded_ws.hyperlinks.transform_values(&:url), mat_ws.hyperlinks.transform_values(&:url))
            assert_equal loaded_ws.print_area, mat_ws.print_area
            assert_equal loaded_ws.freeze_pane&.row, mat_ws.freeze_pane&.row
            assert_equal(loaded_cols, mat_ws.columns.map { |c| [c.index, c.width, c.hidden?, c.outline_level] })
            assert_equal(loaded_rows, mat_ws.rows.map { |r| [r.index, r.cells.map(&:value)] })
          end
          assert_equal 1, stream_count

          padded_indices = []
          Xlsxrb.read(StringIO.new(bytes), pad_empty_rows: true) do |stream_sheet|
            padded_indices = stream_sheet.each_row.map(&:index)
          end
          assert_equal((0..state.rows.last[:row_index]).to_a, padded_indices)
        end
      end
    end
  end
end
