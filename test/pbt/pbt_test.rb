# frozen_string_literal: true

require "test_helper"
require "pbt"
require "tempfile"
require "bigdecimal"
require "date"
require "time"
require_relative "../support/excel_rules"
require_relative "../support/xsd_validator"

class PbtTest < Test::Unit::TestCase
  # Duck-typed Time replacement simulating ActiveSupport::TimeWithZone.
  class FakeTimeWithZone
    attr_reader :year, :month, :day, :hour, :min, :sec, :subsec, :utc_offset

    def initialize(year, month, day, hour, min, sec, subsec = 0, utc_offset = 0)
      @year = year
      @month = month
      @day = day
      @hour = hour
      @min = min
      @sec = sec
      @subsec = subsec
      @utc_offset = utc_offset
    end

    def is_a?(klass)
      klass == Time || super
    end

    def to_time
      Time.new(@year, @month, @day, @hour, @min, @sec + @subsec, @utc_offset)
    end

    def ==(other)
      other.is_a?(FakeTimeWithZone) &&
        year == other.year && month == other.month && day == other.day &&
        hour == other.hour && min == other.min && sec == other.sec &&
        subsec == other.subsec && utc_offset == other.utc_offset
    end
  end

  def setup
    omit "Skip property-based tests under RBS runtime testing to prevent OOM / slowdown" if defined?(RBS::Test)
  end

  # --- PBT Arbitraries & Generators for Union Types and Structures ---

  # Generates sheet names (1-31 chars, alphanumeric)
  def sheet_name_generator
    Pbt.alphanumeric_string(min: 1, max: 31)
  end

  # Generates raw descriptors for the entire union type domain:
  # [type_code (0..13), int_val, float_val, bool_val, str_val, sub_choice (0..7)]
  def cell_value_descriptor_generator
    Pbt.tuple(
      Pbt.choose(0..13),
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
    when 12 # BigDecimal
      BigDecimal(["123.45", "-0.005", "0"][sub_choice % 3])
    when 13 # TimeLike (duck-typed Time replacement)
      year = 1900 + (int_val.abs % 150)
      month = 1 + (sub_choice % 12)
      day = 1 + (int_val.abs % 28)
      hour = int_val.abs % 24
      min = (sub_choice * 7) % 60
      sec = int_val.abs % 60
      offset = [0, 32_400, -18_000][sub_choice % 3]
      FakeTimeWithZone.new(year, month, day, hour, min, sec, 0, offset)
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
                when FakeTimeWithZone
                  actual_time = actual_cell.to_time
                  assert_not_nil actual_time
                  assert_equal(
                    [expected_val.year, expected_val.month, expected_val.day, expected_val.hour, expected_val.min, expected_val.sec],
                    [actual_time.year, actual_time.month, actual_time.day, actual_time.hour, actual_time.min, actual_time.sec]
                  )
                when Time
                  actual_time = actual_cell.to_time
                  assert_not_nil actual_time
                  assert_in_delta expected_val.to_i, actual_time.to_i, 2
                when BigDecimal
                  assert_in_delta expected_val.to_f, actual_val.to_f, 0.0001
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
        Pbt.array(Pbt.array(cell_value_descriptor_generator, min: 1, max: 5), min: 1, max: 5),
        Pbt.constant_from(nil, "Arial", { name: "Calibri", size: 12 })
      ) do |sheet_name, raw_rows_data, default_font|
        rows_data = raw_rows_data.map do |row|
          row.map { |cell_desc| resolve_cell_value(cell_desc) }
        end

        Dir.mktmpdir do |dir|
          streaming_path = File.join(dir, "streaming_union.xlsx")
          in_memory_path = File.join(dir, "in_memory_union.xlsx")

          # 1. Streaming write
          Xlsxrb.write(streaming_path, default_font: default_font) do |wb|
            wb.sheet(sheet_name) do |s|
              rows_data.each { |r| s.row(r) }
            end
          end

          # 2. In-memory write
          wb = Xlsxrb.build(default_font: default_font) do |b|
            b.sheet(sheet_name) do |s|
              rows_data.each { |r| s.row(r) }
            end
          end
          Xlsxrb.write(in_memory_path, wb)

          streaming_bytes = File.binread(streaming_path)
          in_memory_bytes = File.binread(in_memory_path)
          assert_empty ExcelRules.check(streaming_bytes).map(&:to_s)
          assert_empty ExcelRules.check(in_memory_bytes).map(&:to_s)
          assert_empty XsdValidator.check(streaming_bytes)
          assert_empty XsdValidator.check(in_memory_bytes)

          # Invariant: Both write modes yield identical sheets, styles, formatted values, and cell values
          streaming_wb = Xlsxrb.load(streaming_path)
          in_memory_wb = Xlsxrb.load(in_memory_path)

          assert_equal streaming_wb.styles.fonts.first&.name, in_memory_wb.styles.fonts.first&.name
          assert_equal streaming_wb.styles.fonts.first&.size, in_memory_wb.styles.fonts.first&.size

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
            assert_equal s_row.formatted_values, m_row.formatted_values

            s_row.cells.each_with_index do |s_cell, c_idx|
              m_cell = m_row.cells[c_idx]
              assert_equal s_cell.column_index, m_cell.column_index
              assert_equal s_cell.ref, m_cell.ref
              assert_equal s_cell.formula, m_cell.formula
              assert_equal s_cell.formatted_value, m_cell.formatted_value

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
        val = val.to_time if val.is_a?(FakeTimeWithZone)
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
              s.sheet_protection(sheet: opts[:protect_sheet])
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

    Pbt.assert(num_runs: 3) do
      Pbt.property(Pbt.alphanumeric_string(min: 1, max: 12)) do |pw|
        io = StringIO.new
        Xlsxrb.write(io) do |w|
          w.sheet("Protected") do |s|
            s.sheet_protection(sheet: true, password: pw)
            s.row([1])
          end
        end

        prot = Xlsxrb.load(StringIO.new(io.string)).sheet(0).sheet_protection
        assert_not_nil prot
        assert_true prot.sheet?
        assert_equal "SHA-512", prot.algorithm_name
        assert_equal 100_000, prot.spin_count
        assert_not_nil prot.hash_value
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
    Pbt.assert(num_runs: 8) do
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
    Pbt.assert(num_runs: 8) do
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
    Pbt.assert(num_runs: 6) do
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

  # Shared fixture helper for extended worksheet features tested across builders and readers.
  def self.apply_extended_sheet_features(sheet)
    sheet.data_validation("A1:B2", type: :whole, operator: :between, formula1: "0", formula2: "1000")
    sheet.conditional_format("A1:B2", type: :cell_is, operator: :greater_than, formula: "10", fill_color: "FF00FF00")
    sheet.page_setup(orientation: :landscape, paper_size: 9)
    sheet.row_break(1)
  end

  def assert_extended_sheet_features(sheet)
    assert_equal ["A1:B2"], sheet.data_validations.map(&:sqref)
    assert_equal ["A1:B2"], sheet.conditional_formats.map(&:sqref)
    assert_equal "landscape", sheet.page_setup&.orientation
    assert_equal 9, sheet.page_setup&.paper_size
    assert_equal [1], sheet.row_breaks.map(&:id)
  end

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
    :comments,
    :print_titles,
    :extended_features
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

    def set_column_style(pair, col_idx, width:, style:, collapsed: false, best_fit: false)
      pair.builder_sheet.column(col_idx, width: width, style: style, collapsed: collapsed, best_fit: best_fit)
      pair.stream_proxy.column(col_idx, width: width, style: style, collapsed: collapsed, best_fit: best_fit)
      @recorded_ops << [:set_column_style, col_idx, width, style, collapsed, best_fit]
      [
        pair.builder_sheet.columns.map { |c| [c.index, c.collapsed?, c.best_fit?] },
        pair.stream_proxy.columns.map { |c| [c.index, c.collapsed?, c.best_fit?] }
      ]
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

    def set_sheet_features(pair, auto_filter:, merge_range:, comment_ref:, comment_text:, repeat_row: nil, repeat_col: nil)
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
      if repeat_row
        pair.builder_sheet.repeat_rows(repeat_row)
        pair.stream_proxy.repeat_rows(repeat_row)
      end
      if repeat_col
        pair.builder_sheet.repeat_columns(repeat_col)
        pair.stream_proxy.repeat_columns(repeat_col)
      end
      @recorded_ops << [:set_sheet_features, auto_filter, merge_range, comment_ref, comment_text, repeat_row, repeat_col]
      [pair.builder_sheet.print_titles, pair.stream_proxy.print_titles]
    end

    def apply_validation_and_page_setup(pair)
      apply_extended_features_to(pair.builder_sheet)
      apply_extended_features_to(pair.stream_proxy)
      @recorded_ops << [:extended_features]
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
            b_sheet.column(op[1], width: op[2], style: op[3], collapsed: op[4], best_fit: op[5])
            s_proxy.column(op[1], width: op[2], style: op[3], collapsed: op[4], best_fit: op[5])
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
            if op[5]
              b_sheet.repeat_rows(op[5])
              s_proxy.repeat_rows(op[5])
            end
            if op[6]
              b_sheet.repeat_columns(op[6])
              s_proxy.repeat_columns(op[6])
            end
          when :extended_features
            apply_extended_features_to(b_sheet)
            apply_extended_features_to(s_proxy)
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

    def apply_extended_features_to(sheet)
      PbtTest.apply_extended_sheet_features(sheet)
    end

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
      Pbt.constant(BigDecimal("123.45")),
      Pbt.constant(Date.new(2026, 4, 1)),
      Pbt.constant(Time.utc(2026, 4, 1, 9, 30, 0)),
      Pbt.constant(FakeTimeWithZone.new(2026, 4, 1, 9, 30, 0, 0, 32_400))
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
            comments: {},
            print_titles: {},
            extended_features: {}
          )
        end

        m.sut { DualWorkbookSut.new }
        m.teardown(&:cleanup!)

        m.command :add_sheet, weight: 2 do |c|
          c.guard { |s| s.sheets.size < 3 }
          c.arguments Pbt.constant_from("Sheet1", "Sheet2", "Master's Data")
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
              comments: s.comments.merge(name => []),
              print_titles: s.print_titles.merge(name => { rows: nil, columns: nil }),
              extended_features: s.extended_features.merge(name => false)
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
              style: Pbt.constant_from(*s.styles.keys),
              collapsed: Pbt.boolean,
              best_fit: Pbt.boolean
            )
          end
          c.next_state do |s, args|
            sname = s.active_sheet_name
            prev_col = s.sheet_columns.fetch(sname)[args[:col_idx]]
            col_entry = {
              style_name: args[:style],
              style_def: s.styles.fetch(args[:style]),
              collapsed: args[:collapsed] || prev_col&.dig(:collapsed) || false,
              best_fit: args[:best_fit] || prev_col&.dig(:best_fit) || false
            }
            updated_cols = s.sheet_columns.fetch(sname).merge(args[:col_idx] => col_entry)
            s.with(sheet_columns: s.sheet_columns.merge(sname => updated_cols))
          end
          c.run do |sut, args|
            sut.set_column_style(
              args[:sheet],
              args[:col_idx],
              width: args[:width],
              style: args[:style],
              collapsed: args[:collapsed],
              best_fit: args[:best_fit]
            )
          end
          c.verify { |ctx| assert_equal ctx.result[0], ctx.result[1] }
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
              comment_text: Pbt.alphanumeric_string(min: 1, max: 12),
              repeat_row: Pbt.constant_from(nil, 0, 1),
              repeat_col: Pbt.constant_from(nil, 0, 1)
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
            cur_pt = s.print_titles.fetch(sname)
            next_rows = args[:repeat_row] ? "#{args[:repeat_row] + 1}:#{args[:repeat_row] + 1}" : cur_pt[:rows]
            next_cols = if args[:repeat_col]
                          col_l = Xlsxrb::Elements::Cell._column_letter(args[:repeat_col])
                          "#{col_l}:#{col_l}"
                        else
                          cur_pt[:columns]
                        end
            s.with(
              auto_filters: s.auto_filters.merge(sname => "A1:B2"),
              merges: s.merges.merge(sname => new_merges),
              comments: s.comments.merge(sname => new_comments),
              print_titles: s.print_titles.merge(sname => { rows: next_rows, columns: next_cols })
            )
          end
          c.run do |sut, args|
            sut.set_sheet_features(
              args[:sheet],
              auto_filter: "A1:B2",
              merge_range: "A1:B1",
              comment_ref: args[:comment_ref],
              comment_text: args[:comment_text],
              repeat_row: args[:repeat_row],
              repeat_col: args[:repeat_col]
            )
          end
          c.verify { |ctx| assert_equal ctx.result[0], ctx.result[1] }
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

        m.command :set_validation_and_page_setup, weight: 2 do |c|
          c.guard { |s| !s.active_sheet_ref.nil? && !s.extended_features.fetch(s.active_sheet_name) }
          c.arguments { |s| Pbt.constant(s.active_sheet_ref) }
          c.next_state do |s, _sheet_ref|
            s.with(extended_features: s.extended_features.merge(s.active_sheet_name => true))
          end
          c.run { |sut, sheet_pair| sut.apply_validation_and_page_setup(sheet_pair) }
          c.verify { |ctx| assert_nil ctx.result }
        end

        m.invariant(:live_and_serialized_parity) do |state, sut|
          assert_equal state.sheets, sut.builder.sheet_names
          assert_equal state.sheets, sut.stream_writer.sheet_names
          next if state.sheets.empty?

          built_bytes, streamed_bytes = sut.snapshot_bytes
          assert_empty ExcelRules.check(built_bytes).map(&:to_s)
          assert_empty ExcelRules.check(streamed_bytes).map(&:to_s)
          assert_empty XsdValidator.check(built_bytes)
          assert_empty XsdValidator.check(streamed_bytes)

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

            exp_cols = state.sheet_columns.fetch(sname).sort_by(&:first).map do |idx, col|
              [idx, col[:collapsed], col[:best_fit]]
            end
            assert_equal(exp_cols, b_sheet.columns.map { |c| [c.index, c.collapsed?, c.best_fit?] })
            assert_equal(exp_cols, s_sheet.columns.map { |c| [c.index, c.collapsed?, c.best_fit?] })
            exp_pt = state.print_titles.fetch(sname)
            if exp_pt[:rows] || exp_pt[:columns]
              assert_equal exp_pt[:rows], b_sheet.print_titles&.rows
              assert_equal exp_pt[:columns], b_sheet.print_titles&.columns
              assert_equal exp_pt[:rows], s_sheet.print_titles&.rows
              assert_equal exp_pt[:columns], s_sheet.print_titles&.columns
            else
              assert_nil b_sheet.print_titles
              assert_nil s_sheet.print_titles
            end

            if state.extended_features.fetch(sname)
              assert_extended_sheet_features(b_sheet)
              assert_extended_sheet_features(s_sheet)
              assert_equal b_sheet.conditional_formats.map(&:dxf_id), s_sheet.conditional_formats.map(&:dxf_id)
            end

            expected_rows.each_with_index do |exp_row, r_idx|
              b_row = b_sheet.rows[r_idx]
              s_row = s_sheet.rows[r_idx]
              assert_equal exp_row[:cells].size, b_row.cells.size
              assert_equal exp_row[:cells].size, s_row.cells.size
              assert_equal b_row.formatted_values, s_row.formatted_values

              exp_row[:cells].each_with_index do |exp_cell, c_idx|
                b_cell = b_row.cells[c_idx]
                s_cell = s_row.cells[c_idx]
                exp_val = exp_cell[:value]

                case exp_val
                when Date
                  assert_equal exp_val, b_cell.to_date
                  assert_equal exp_val, s_cell.to_date
                when FakeTimeWithZone
                  bt = b_cell.to_time
                  st = s_cell.to_time
                  exp_wall = [exp_val.year, exp_val.month, exp_val.day, exp_val.hour, exp_val.min, exp_val.sec]
                  assert_equal exp_wall, [bt.year, bt.month, bt.day, bt.hour, bt.min, bt.sec]
                  assert_equal exp_wall, [st.year, st.month, st.day, st.hour, st.min, st.sec]
                when Time
                  assert_in_delta exp_val.to_i, b_cell.to_time.to_i, 1
                  assert_in_delta exp_val.to_i, s_cell.to_time.to_i, 1
                when BigDecimal
                  assert_in_delta exp_val.to_f, b_cell.value.to_f, 0.0001
                  assert_in_delta exp_val.to_f, s_cell.value.to_f, 0.0001
                else
                  assert_equal exp_val, b_cell.value
                  assert_equal exp_val, s_cell.value
                end

                assert_equal b_cell.formatted_value, s_cell.formatted_value
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
          assert_empty XsdValidator.check(bytes)

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
    :formulas,
    :columns,
    :auto_filter,
    :merges,
    :comments,
    :hyperlinks,
    :print_area,
    :freeze_row,
    :extended_features
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

    def set_column(col_idx, width:, hidden:, outline_level:, collapsed: false, best_fit: false)
      @ops << [:column, col_idx, width, hidden, outline_level, collapsed, best_fit]
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

    def apply_extended_features
      @ops << [:extended_features]
      nil
    end

    def serialize_bytes
      io = StringIO.new
      Xlsxrb.write(io) do |wb|
        wb.sheet("Data") do |sheet|
          @ops.each do |op|
            case op[0]
            when :row
              row_vals = op[1].map do |v|
                if v.is_a?(Hash) && v.key?(:rich_text)
                  Xlsxrb::Elements::RichText.new(runs: [{ text: v[:rich_text][0], font: { bold: true } }, { text: v[:rich_text][1] }])
                else
                  v
                end
              end
              sheet.row(row_vals)
            when :skip_rows
              sheet.skip_rows(op[1])
            when :column
              sheet.column(op[1], width: op[2], hidden: op[3], outline_level: op[4], collapsed: op[5], best_fit: op[6])
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
            when :extended_features
              PbtTest.apply_extended_sheet_features(sheet)
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
      Pbt.constant(Date.new(2026, 6, 15)),
      Pbt.constant({ formula: "1+2", value: 3 }),
      Pbt.constant({ rich_text: %w[Bold Plain] })
    )

    Pbt.assert(num_runs: 15) do
      Pbt.stateful(max_steps: 14) do |m|
        m.initial_state do
          MultiReaderState.new(
            row_cursor: 0,
            rows: [],
            formulas: [],
            columns: {},
            auto_filter: nil,
            merges: [],
            comments: {},
            hyperlinks: {},
            print_area: nil,
            freeze_row: nil,
            extended_features: false
          )
        end

        m.sut { MultiReaderSut.new }

        m.command :append_row, weight: 4 do |c|
          c.arguments(v0: val_arb, v1: val_arb)
          c.next_state do |s, args|
            pair = [args[:v0], args[:v1]]
            expected = pair.map do |v|
              case v
              when Date then Xlsxrb.date_to_serial(v)
              when Hash then v.key?(:rich_text) ? v[:rich_text].join : v[:value]
              else v
              end
            end
            exp_formulas = pair.map { |v| v.is_a?(Hash) && v.key?(:formula) ? v[:formula] : nil }
            row_record = { row_index: s.row_cursor, values: expected }
            formula_record = { row_index: s.row_cursor, formulas: exp_formulas }
            s.with(row_cursor: s.row_cursor + 1, rows: [*s.rows, row_record], formulas: [*s.formulas, formula_record])
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
            outline_level: Pbt.constant_from(nil, 1, 2),
            collapsed: Pbt.boolean,
            best_fit: Pbt.boolean
          )
          c.next_state do |s, args|
            prev_col = s.columns[args[:col_idx]]
            col_spec = {
              width: args[:width].to_f,
              hidden: args[:hidden] || prev_col&.dig(:hidden) || false,
              outline_level: args[:outline_level].nil? ? prev_col&.dig(:outline_level) : args[:outline_level],
              collapsed: args[:collapsed] || prev_col&.dig(:collapsed) || false,
              best_fit: args[:best_fit] || prev_col&.dig(:best_fit) || false
            }
            s.with(columns: s.columns.merge(args[:col_idx] => col_spec))
          end
          c.run do |sut, args|
            sut.set_column(
              args[:col_idx],
              width: args[:width],
              hidden: args[:hidden],
              outline_level: args[:outline_level],
              collapsed: args[:collapsed],
              best_fit: args[:best_fit]
            )
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

        m.command :set_validation_and_page_setup, weight: 2 do |c|
          c.guard { |s| !s.extended_features }
          c.next_state { |s| s.with(extended_features: true) }
          c.run { |sut, _args| sut.apply_extended_features }
        end

        m.invariant(:three_way_reader_equivalence) do |state, sut|
          next if state.rows.empty?

          bytes = sut.serialize_bytes
          assert_empty ExcelRules.check(bytes).map(&:to_s)
          assert_empty XsdValidator.check(bytes)

          loaded_ws = Xlsxrb.load(StringIO.new(bytes)).sheet(0)
          expected_rows = state.rows.map { |r| [r[:row_index], r[:values]] }
          expected_formulas = state.formulas.map { |r| [r[:row_index], r[:formulas]] }
          loaded_rows = loaded_ws.rows.map { |r| [r.index, r.cells.map(&:value)] }
          loaded_formulas = loaded_ws.rows.map { |r| [r.index, r.cells.map { |c| c.formula&.expression }] }
          assert_equal expected_rows, loaded_rows
          assert_equal expected_formulas, loaded_formulas
          assert_equal state.auto_filter, loaded_ws.auto_filter
          assert_equal state.merges, loaded_ws.merged_cells
          assert_equal(state.comments.to_a, loaded_ws.comments.map { |c| [c.ref, c.text] })
          assert_equal(state.hyperlinks, loaded_ws.hyperlinks.transform_values(&:url))
          assert_equal state.print_area, loaded_ws.print_area
          assert_equal state.freeze_row, loaded_ws.freeze_pane&.row
          assert_extended_sheet_features(loaded_ws) if state.extended_features

          expected_cols = state.columns.sort_by(&:first).map do |idx, col|
            [idx, col[:width], col[:hidden], col[:outline_level], col[:collapsed], col[:best_fit]]
          end
          loaded_cols = loaded_ws.columns.map { |c| [c.index, c.width, c.hidden?, c.outline_level, c.collapsed?, c.best_fit?] }
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
            assert_extended_sheet_features(stream_sheet) if state.extended_features
            assert_equal(loaded_cols, stream_sheet.columns.map { |c| [c.index, c.width, c.hidden?, c.outline_level, c.collapsed?, c.best_fit?] })
            stream_rows = stream_sheet.each_row.to_a
            assert_equal(loaded_rows, stream_rows.map { |r| [r.index, r.values] })
            assert_equal(loaded_formulas, stream_rows.map { |r| [r.index, r.cells.map { |c| c.formula&.expression }] })

            mat_ws = stream_sheet.load
            assert_equal loaded_ws.name, mat_ws.name
            assert_equal loaded_ws.auto_filter, mat_ws.auto_filter
            assert_equal loaded_ws.merged_cells, mat_ws.merged_cells
            assert_equal(loaded_ws.comments.map { |c| [c.ref, c.text] }, mat_ws.comments.map { |c| [c.ref, c.text] })
            assert_equal(loaded_ws.hyperlinks.transform_values(&:url), mat_ws.hyperlinks.transform_values(&:url))
            assert_equal loaded_ws.print_area, mat_ws.print_area
            assert_equal loaded_ws.freeze_pane&.row, mat_ws.freeze_pane&.row
            assert_extended_sheet_features(mat_ws) if state.extended_features
            assert_equal(loaded_cols, mat_ws.columns.map { |c| [c.index, c.width, c.hidden?, c.outline_level, c.collapsed?, c.best_fit?] })
            assert_equal(loaded_rows, mat_ws.rows.map { |r| [r.index, r.cells.map(&:value)] })
            assert_equal(loaded_formulas, mat_ws.rows.map { |r| [r.index, r.cells.map { |c| c.formula&.expression }] })
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

  test "XmlBuilder.encode_xstring, escape, and unescape roundtrip property across control chars and _xHHHH_ literals" do
    segment_arb = Pbt.frequency(
      [3, Pbt.alphanumeric_string(min: 0, max: 12)],
      [2, Pbt.constant_from("\x00", "\x01", "\x08", "\x1F", "\uFFFE")],
      [2, Pbt.constant_from("_x0041_", "_x005F_", "_xdead_", "_x005F_x0041_")],
      [2, Pbt.constant_from("&", "<", ">", "\"", "'", "&#60;", "&amp;lt;")]
    )

    Pbt.assert(num_runs: 100) do
      Pbt.property(Pbt.array(segment_arb, min: 0, max: 12)) do |segments|
        str = segments.join
        encoded = Xlsxrb::Ooxml::XmlBuilder.encode_xstring(str)
        escaped = Xlsxrb::Ooxml::XmlBuilder.escape(encoded)
        unescaped = Xlsxrb::Ooxml::XmlBuilder.unescape(escaped)
        assert_equal str, unescaped
      end
    end
  end

  test "Xlsxrb.truncate_cell_text UTF-16 code unit boundary and character integrity property" do
    unit_arb = Pbt.constant_from("a", "漢", "🎉")

    Pbt.assert(num_runs: 80) do
      Pbt.property(
        unit_arb,
        Pbt.integer(min: 32_758, max: 32_774),
        Pbt.array(unit_arb, min: 1, max: 6)
      ) do |fill_unit, base_units, tail_units|
        fill_unit_size = Xlsxrb::Ooxml::Utils.utf16_length(fill_unit)
        repeat_count = base_units / fill_unit_size
        str = (fill_unit * repeat_count) + tail_units.join

        truncated = Xlsxrb.truncate_cell_text(str)
        assert_true truncated.valid_encoding?
        assert_operator Xlsxrb::Ooxml::Utils.utf16_length(truncated), :<=, 32_767
        assert_true str.start_with?(truncated)

        assert_equal str, truncated if Xlsxrb::Ooxml::Utils.utf16_length(str) <= 32_767
      end
    end
  end

  test "Xlsxrb.format_value determinism and safety across numeric, temporal, and edge values" do
    all_formats = [nil, "General", "0.00", "#,##0", "0.0%", "yyyy-mm-dd", "hh:mm:ss", "@"]

    val_arb = Pbt.one_of(
      Pbt.integer(min: -100_000, max: 100_000),
      Pbt.float,
      Pbt.constant_from(Float::NAN, Float::INFINITY, -Float::INFINITY),
      Pbt.date(base_date: Date.new(2026, 6, 15), past_offset_days: -36_500, future_offset_days: 36_500),
      Pbt.constant(Time.utc(2026, 6, 15, 10, 20, 30)),
      Pbt.boolean,
      Pbt.alphanumeric_string(min: 0, max: 15),
      Pbt.nil
    )

    Pbt.assert(num_runs: 60) do
      Pbt.property(val_arb, Pbt.constant_from(*all_formats)) do |val, fmt|
        formatted = Xlsxrb.format_value(val, fmt)
        assert_instance_of String, formatted
        assert_includes [Encoding::UTF_8, Encoding::US_ASCII], formatted.encoding
        assert_true formatted.valid_encoding?
        assert_equal formatted, Xlsxrb.format_value(val, fmt)
      end
    end
  end

  test "Ooxml::Cfb::Writer and Cfb::Reader roundtrip across mini-stream, multi-sector directory, and multi-FAT boundaries" do
    stream_spec_arb = Pbt.tuple(
      Pbt.alphanumeric_string(min: 1, max: 14),
      Pbt.choose(0..3),
      Pbt.ascii_string(min: 1, max: 300)
    )

    Pbt.assert(num_runs: 30) do
      Pbt.property(Pbt.array(stream_spec_arb, min: 1, max: 6)) do |specs|
        streams_by_lower = {}
        specs.each do |name, kind, seed_chunk|
          key = name.downcase
          next if streams_by_lower.key?(key)

          payload = case kind
                    when 0 then "".b
                    when 1 then seed_chunk.b
                    when 2 then (("R" * 4096) + seed_chunk).b
                    else (("L" * 66_000) + seed_chunk).b
                    end
          streams_by_lower[key] = [name, payload]
        end
        input = streams_by_lower.values.to_h

        cfb_bytes = Xlsxrb::Ooxml::Cfb::Writer.write(input)
        assert_true Xlsxrb::Ooxml::Cfb::Reader.cfb?(cfb_bytes)

        reader = Xlsxrb::Ooxml::Cfb::Reader.new(cfb_bytes)
        assert_equal input.keys, reader.stream_names
        input.each do |name, expected|
          assert_equal expected, reader.read_stream(name)
        end
      end
    end
  end

  test "Ooxml::SharedStringsParser.parse and SharedStringsListener equivalence across synthetic and rich-text SST XML" do
    si_spec_arb = Pbt.tuple(
      Pbt.choose(0..5),
      Pbt.alphanumeric_string(min: 0, max: 12),
      Pbt.alphanumeric_string(min: 1, max: 8),
      Pbt.boolean
    )

    Pbt.assert(num_runs: 80) do
      Pbt.property(Pbt.boolean, Pbt.array(si_spec_arb, min: 1, max: 12)) do |use_ns, specs|
        p = use_ns ? "x:" : ""
        xmlns = use_ns ? 'xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' : 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        sis = specs.map do |kind, t1, t2, bold|
          esc1 = Xlsxrb::Ooxml::XmlBuilder.escape(Xlsxrb::Ooxml::XmlBuilder.encode_xstring(t1))
          esc2 = Xlsxrb::Ooxml::XmlBuilder.escape(Xlsxrb::Ooxml::XmlBuilder.encode_xstring(t2))
          case kind
          when 0 then "<#{p}si/>"
          when 1 then "<#{p}si><#{p}t/></#{p}si>"
          when 2 then %(<#{p}si><#{p}t xml:space="preserve">#{esc1}</#{p}t></#{p}si>)
          when 3 then %(<#{p}si><#{p}t>#{esc1}</#{p}t><#{p}rPh sb="0" eb="1"><#{p}t>#{esc2}</#{p}t></#{p}rPh></#{p}si>)
          when 4 then "<#{p}si><#{p}r><#{p}rPr>#{"<#{p}b/>" if bold}</#{p}rPr><#{p}t>#{esc1}</#{p}t></#{p}r><#{p}r><#{p}t>#{esc2}</#{p}t></#{p}r></#{p}si>"
          else %(<#{p}si><#{p}r><#{p}t/></#{p}r><#{p}r><#{p}t>#{esc2}</#{p}t></#{p}r><#{p}rPh sb="0" eb="1"><#{p}t>フリガナ</#{p}t></#{p}rPh></#{p}si>)
          end
        end.join
        xml = %(<?xml version="1.0" encoding="UTF-8"?><#{p}sst #{xmlns} count="#{specs.size}" uniqueCount="#{specs.size}">#{sis}</#{p}sst>)

        fast_parsed = Xlsxrb::Ooxml::SharedStringsParser.parse(xml)
        listener = Xlsxrb::Ooxml::Reader::SharedStringsListener.new
        Xlsxrb::Ooxml::XmlParser.parse(xml, listener)
        assert_equal listener.strings.map(&:to_s), fast_parsed
      end
    end
  end

  test "Ooxml::WorksheetParser.scan_rows_stream chunk boundary invariance against fast_scan_rows_direct" do
    row_spec_arb = Pbt.tuple(
      Pbt.boolean, # self-closing row
      Pbt.integer(min: -1000, max: 1000),
      Pbt.boolean,
      Pbt.alphanumeric_string(min: 0, max: 12)
    )

    Pbt.assert(num_runs: 80) do
      Pbt.property(Pbt.boolean, Pbt.integer(min: 1, max: 64), Pbt.array(row_spec_arb, min: 1, max: 6)) do |use_ns, chunk_size, row_specs|
        p = use_ns ? "x:" : ""
        xmlns = use_ns ? 'xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' : 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        row_num = 1
        rows_xml = row_specs.map do |self_close, int_val, flag, txt|
          r = row_num
          row_num += flag ? 2 : 1
          if self_close
            %(<#{p}row r="#{r}"/>)
          else
            esc = Xlsxrb::Ooxml::XmlBuilder.escape("#{txt} 漢字 🎉")
            attrs = flag ? 'ht="20.5" customHeight="1" hidden="1" collapsed="1" outlineLevel="2"' : ""
            %(<#{p}row r="#{r}" #{attrs}><#{p}c r="A#{r}" t="inlineStr"><#{p}is><#{p}t>#{esc}</#{p}t></#{p}is></#{p}c><#{p}c r="B#{r}"><#{p}v>#{int_val}</#{p}v></#{p}c><#{p}c r="C#{r}" t="b"><#{p}v>#{flag ? 1 : 0}</#{p}v></#{p}c><#{p}c r="D#{r}" t="s"><#{p}v>0</#{p}v></#{p}c></#{p}row>)
          end
        end.join
        xml = %(<?xml version="1.0" encoding="UTF-8"?><#{p}worksheet #{xmlns}><#{p}sheetData>#{rows_xml}</#{p}sheetData><#{p}rowBreaks count="1"><#{p}brk id="1" max="16383" man="1"/></#{p}rowBreaks></#{p}worksheet>)

        shared = ["Shared0"]
        direct_rows = []
        Xlsxrb::Ooxml::WorksheetParser.send(:fast_scan_rows_direct, xml, shared, "sheet1.xml") do |row|
          direct_rows << [row.index, row.values, row.attributes]
        end

        stream_rows = []
        supplier = proc do |&blk|
          bytes = xml.b
          pos = 0
          while pos < bytes.bytesize
            blk.call(bytes.byteslice(pos, chunk_size))
            pos += chunk_size
          end
        end
        Xlsxrb::Ooxml::WorksheetParser.send(:scan_rows_stream, supplier, shared, "sheet1.xml") do |row|
          stream_rows << [row.index, row.values, row.attributes]
        end

        assert_equal direct_rows, stream_rows
      end
    end
  end

  test "NumberFormatter.format multi-section, scientific, percentage, and date serial invariants" do
    num_arb = Pbt.one_of(
      Pbt.integer(min: -1_000_000, max: 1_000_000),
      Pbt.float.filter(&:finite?)
    )
    fmt_arb = Pbt.constant_from(
      '0.00;"neg "0.00;"zero";"txt "@',
      '#,##0.00;(#,##0.00);"-"',
      "[Red]#,##0;[Blue]-#,##0;0",
      "0.00E+00",
      "##0.0E+0",
      "0.00%",
      "$#,##0.00"
    )

    Pbt.assert(num_runs: 100) do
      Pbt.property(num_arb, fmt_arb) do |val, fmt|
        formatted = Xlsxrb::NumberFormatter.format(val, fmt)
        assert_false formatted.empty?
        assert_equal Encoding::UTF_8, formatted.encoding
        assert_true formatted.valid_encoding?
        assert_equal formatted, Xlsxrb::NumberFormatter.format(val, fmt)
      end

      Pbt.property(Pbt.integer(min: 1, max: 2_958_465), Pbt.constant_from("yyyy-mm-dd", "yyyy-mm-dd hh:mm:ss")) do |serial, dfmt|
        formatted = Xlsxrb::NumberFormatter.format(serial, dfmt)
        assert_equal Encoding::UTF_8, formatted.encoding
        assert_match(/\A\d{4}-\d{2}-\d{2}/, formatted)
        assert_equal formatted, Xlsxrb::NumberFormatter.format(serial, dfmt)
      end
    end
  end

  MINIMAL_PNG_BYTES = [
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
    0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
    0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
    0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
    0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
    0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC,
    0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
    0x44, 0xAE, 0x42, 0x60, 0x82
  ].pack("C*").freeze

  # --- Stateful PBT #1: Multi-sheet compound parts & dual-builder parity ---

  MultiPartsState = Data.define(
    :date1904,
    :shared_strings,
    :sheets,
    :active_sheet,
    :col_grouped,
    :row_counts,
    :rows_data,
    :charts,
    :shapes,
    :images,
    :tables,
    :filters,
    :comments,
    :defined_names
  )

  # Drives live WorkbookBuilder and StreamWriter across multi-sheet compound parts and replays ops for snapshots.
  class MultiPartsSut
    def initialize(date1904:, shared_strings:)
      @date1904 = date1904
      @shared_strings = shared_strings
      @builder = Xlsxrb::WorkbookBuilder.new
      @builder.workbook_properties(date1904: true) if @date1904
      @stream_io = StringIO.new
      @stream_writer = Xlsxrb::StreamWriter.new(@stream_io, shared_strings: @shared_strings)
      @stream_writer.workbook_properties(date1904: true) if @date1904
      @active_fiber = nil
      @b_sheet = nil
      @s_proxy = nil
      @current_row_count = 0
      @ops = []
    end

    def add_sheet(name)
      close_active_fiber
      fiber = Fiber.new do
        @builder.sheet(name) do |sh|
          @b_sheet = sh
          Fiber.yield
        end
      end
      fiber.resume
      @active_fiber = fiber
      @stream_writer.sheet(name)
      @s_proxy = @stream_writer[]
      @current_row_count = 0
      @ops << [:add_sheet, name]
      nil
    end

    def group_columns(collapsed:)
      @b_sheet.group_columns(0..1, collapsed: collapsed)
      @s_proxy.group_columns(0..1, collapsed: collapsed)
      @ops << [:group_columns, collapsed]
      nil
    end

    def append_next_row(label, val, day_off)
      values = @current_row_count.zero? ? %w[Cat Val When] : [label, val, Date.new(2026, 6, 15) + day_off]
      @b_sheet.row(values)
      @s_proxy.row(values)
      @current_row_count += 1
      @ops << [:append_row, values]
      nil
    end

    def add_chart(spec)
      @b_sheet.chart(**spec)
      @s_proxy.chart(**spec)
      @ops << [:add_chart, spec]
      nil
    end

    def add_shape(spec)
      @b_sheet.shape(**spec)
      @s_proxy.shape(**spec)
      @ops << [:add_shape, spec]
      nil
    end

    def add_image(spec)
      @b_sheet.image(MINIMAL_PNG_BYTES, **spec)
      @s_proxy.image(MINIMAL_PNG_BYTES, **spec)
      @ops << [:add_image, spec]
      nil
    end

    def add_table(ref, name:, columns:)
      @b_sheet.table(ref, name: name, columns: columns)
      @s_proxy.table(ref, name: name, columns: columns)
      @ops << [:add_table, ref, name, columns]
      nil
    end

    def add_filter_and_sort(auto_filter:, filter_col:, filter_val:, sort_ref:, cond_ref:)
      @b_sheet.auto_filter = auto_filter
      @s_proxy.auto_filter = auto_filter
      @b_sheet.filter_column(filter_col, { type: :filters, values: [filter_val] })
      @s_proxy.filter_column(filter_col, { type: :filters, values: [filter_val] })
      @b_sheet.sort_state(sort_ref, [{ ref: cond_ref, descending: true }])
      @s_proxy.sort_state(sort_ref, [{ ref: cond_ref, descending: true }])
      @ops << [:add_filter_and_sort, auto_filter, filter_col, filter_val, sort_ref, cond_ref]
      nil
    end

    def add_comment(ref, text)
      @b_sheet.comment(ref, text)
      @s_proxy.comment(ref, text)
      @ops << [:add_comment, ref, text]
      nil
    end

    def add_defined_name(name, range:, sheet:, scope:)
      @builder.defined_name(name, range: range, sheet: sheet, scope: scope)
      @stream_writer.defined_name(name, range: range, sheet: sheet, scope: scope)
      @ops << [:add_defined_name, name, range, sheet, scope]
      nil
    end

    def snapshot_bytes
      snap_builder = Xlsxrb::WorkbookBuilder.new
      snap_builder.workbook_properties(date1904: true) if @date1904
      snap_io = StringIO.new
      snap_stream = Xlsxrb::StreamWriter.new(snap_io, shared_strings: @shared_strings)
      snap_stream.workbook_properties(date1904: true) if @date1904
      snap_fiber = nil
      b_sh = nil
      s_pr = nil

      begin
        @ops.each do |op|
          case op[0]
          when :add_sheet
            snap_fiber.resume if snap_fiber&.alive?
            snap_fiber = Fiber.new do
              snap_builder.sheet(op[1]) do |sh|
                b_sh = sh
                Fiber.yield
              end
            end
            snap_fiber.resume
            snap_stream.sheet(op[1])
            s_pr = snap_stream[]
          when :group_columns
            b_sh.group_columns(0..1, collapsed: op[1])
            s_pr.group_columns(0..1, collapsed: op[1])
          when :append_row
            b_sh.row(op[1])
            s_pr.row(op[1])
          when :add_chart
            b_sh.chart(**op[1])
            s_pr.chart(**op[1])
          when :add_shape
            b_sh.shape(**op[1])
            s_pr.shape(**op[1])
          when :add_image
            b_sh.image(MINIMAL_PNG_BYTES, **op[1])
            s_pr.image(MINIMAL_PNG_BYTES, **op[1])
          when :add_table
            b_sh.table(op[1], name: op[2], columns: op[3])
            s_pr.table(op[1], name: op[2], columns: op[3])
          when :add_filter_and_sort
            b_sh.auto_filter = op[1]
            s_pr.auto_filter = op[1]
            b_sh.filter_column(op[2], { type: :filters, values: [op[3]] })
            s_pr.filter_column(op[2], { type: :filters, values: [op[3]] })
            b_sh.sort_state(op[4], [{ ref: op[5], descending: true }])
            s_pr.sort_state(op[4], [{ ref: op[5], descending: true }])
          when :add_comment
            b_sh.comment(op[1], op[2])
            s_pr.comment(op[1], op[2])
          when :add_defined_name
            snap_builder.defined_name(op[1], range: op[2], sheet: op[3], scope: op[4])
            snap_stream.defined_name(op[1], range: op[2], sheet: op[3], scope: op[4])
          end
        end

        snap_fiber.resume if snap_fiber&.alive?
        snap_stream.close
        mem_bytes = Xlsxrb.write(snap_builder.build, shared_strings: @shared_strings)
        [mem_bytes, snap_io.string.b]
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

  test "stateful PBT: multi-sheet compound parts numbering, relationships, and dual-builder parity" do
    init_arb = Pbt.fixed_hash(date1904: Pbt.boolean, shared_strings: Pbt.boolean).map(
      lambda do |h|
        MultiPartsState.new(
          date1904: h[:date1904],
          shared_strings: h[:shared_strings],
          sheets: [],
          active_sheet: nil,
          col_grouped: {},
          row_counts: {},
          rows_data: {},
          charts: {},
          shapes: {},
          images: {},
          tables: {},
          filters: {},
          comments: {},
          defined_names: []
        )
      end,
      ->(s) { { date1904: s.date1904, shared_strings: s.shared_strings } }
    )

    Pbt.assert(num_runs: 10) do
      Pbt.stateful(max_steps: 14) do |m|
        m.initial_state init_arb
        m.sut { |init| MultiPartsSut.new(date1904: init.date1904, shared_strings: init.shared_strings) }
        m.teardown(&:cleanup!)

        m.command :add_sheet, weight: 3 do |c|
          c.guard { |s| s.sheets.size < 3 }
          c.arguments Pbt.constant_from("Part1", "Data's", "Summary")
          c.precondition { |s, name| !s.sheets.include?(name) }
          c.next_state do |s, name|
            s.with(
              sheets: [*s.sheets, name],
              active_sheet: name,
              col_grouped: s.col_grouped.merge(name => false),
              row_counts: s.row_counts.merge(name => 0),
              rows_data: s.rows_data.merge(name => []),
              charts: s.charts.merge(name => []),
              shapes: s.shapes.merge(name => []),
              images: s.images.merge(name => []),
              tables: s.tables.merge(name => []),
              filters: s.filters.merge(name => false),
              comments: s.comments.merge(name => [])
            )
          end
          c.run { |sut, name| sut.add_sheet(name) }
        end

        m.command :group_columns, weight: 1 do |c|
          c.guard { |s| !s.active_sheet.nil? && s.row_counts.fetch(s.active_sheet).zero? && !s.col_grouped.fetch(s.active_sheet) }
          c.arguments Pbt.boolean
          c.next_state { |s, collapsed| s.with(col_grouped: s.col_grouped.merge(s.active_sheet => collapsed)) }
          c.run { |sut, collapsed| sut.group_columns(collapsed: collapsed) }
        end

        m.command :append_row, weight: 4 do |c|
          c.guard { |s| !s.active_sheet.nil? }
          c.arguments(label: Pbt.constant_from("A", "B", "C"), val: Pbt.integer(min: 1, max: 500), day_off: Pbt.integer(min: 0, max: 30))
          c.next_state do |s, args|
            sname = s.active_sheet
            count = s.row_counts.fetch(sname)
            row_vals = count.zero? ? %w[Cat Val When] : [args[:label], args[:val], Date.new(2026, 6, 15) + args[:day_off]]
            s.with(
              row_counts: s.row_counts.merge(sname => count + 1),
              rows_data: s.rows_data.merge(sname => [*s.rows_data.fetch(sname), row_vals])
            )
          end
          c.run do |sut, args|
            sut.append_next_row(args[:label], args[:val], args[:day_off])
          end
        end

        m.command :add_chart, weight: 4 do |c|
          c.guard { |s| !s.active_sheet.nil? && s.row_counts.fetch(s.active_sheet) >= 2 && s.charts.fetch(s.active_sheet).size < 2 }
          c.arguments do |s|
            sname = s.active_sheet
            idx = s.charts.fetch(sname).size
            last_r = s.row_counts.fetch(sname)
            quoted = Xlsxrb::DslHelpers.quote_sheet_name(sname)
            Pbt.fixed_hash(
              type: Pbt.constant_from(:bar, :line),
              title: Pbt.constant("Chart_#{sname}_#{idx + 1}"),
              from_column: Pbt.constant(3 + (idx * 6)),
              from_row: Pbt.constant(0),
              to_column: Pbt.constant(8 + (idx * 6)),
              to_row: Pbt.constant(9),
              series: Pbt.constant([{ name: "S#{idx + 1}", categories: "#{quoted}!$A$2:$A$#{last_r}", values: "#{quoted}!$B$2:$B$#{last_r}" }])
            )
          end
          c.next_state do |s, spec|
            sname = s.active_sheet
            entry = { type: spec[:type] == :line ? "lineChart" : "barChart", title: spec[:title] }
            s.with(charts: s.charts.merge(sname => [*s.charts.fetch(sname), entry]))
          end
          c.run { |sut, spec| sut.add_chart(spec) }
        end

        m.command :add_shape, weight: 2 do |c|
          c.guard { |s| !s.active_sheet.nil? && s.shapes.fetch(s.active_sheet).size < 2 }
          c.arguments do |s|
            sname = s.active_sheet
            idx = s.shapes.fetch(sname).size
            Pbt.fixed_hash(
              preset: Pbt.constant_from("rect", "ellipse"),
              text: Pbt.constant("Box_#{sname}_#{idx + 1}"),
              from_column: Pbt.constant(3),
              from_row: Pbt.constant(11 + (idx * 5)),
              to_column: Pbt.constant(6),
              to_row: Pbt.constant(15 + (idx * 5))
            )
          end
          c.next_state do |s, spec|
            sname = s.active_sheet
            entry = { preset: spec[:preset], text: spec[:text] }
            s.with(shapes: s.shapes.merge(sname => [*s.shapes.fetch(sname), entry]))
          end
          c.run { |sut, spec| sut.add_shape(spec) }
        end

        m.command :add_image, weight: 3 do |c|
          c.guard { |s| !s.active_sheet.nil? && s.images.fetch(s.active_sheet).size < 2 }
          c.arguments do |s|
            idx = s.images.fetch(s.active_sheet).size
            Pbt.fixed_hash(
              ext: Pbt.constant("png"),
              from_column: Pbt.constant(0),
              from_row: Pbt.constant(10 + (idx * 5)),
              to_column: Pbt.constant(2),
              to_row: Pbt.constant(14 + (idx * 5))
            )
          end
          c.next_state do |s, spec|
            sname = s.active_sheet
            s.with(images: s.images.merge(sname => [*s.images.fetch(sname), spec]))
          end
          c.run { |sut, spec| sut.add_image(spec) }
        end

        m.command :add_table, weight: 2 do |c|
          c.guard { |s| !s.active_sheet.nil? && s.row_counts.fetch(s.active_sheet) >= 2 && s.tables.fetch(s.active_sheet).empty? }
          c.arguments do |s|
            sname = s.active_sheet
            sidx = s.sheets.index(sname) + 1
            Pbt.fixed_hash(
              ref: Pbt.constant("A1:C#{s.row_counts.fetch(sname)}"),
              name: Pbt.constant("TblSheet#{sidx}"),
              columns: Pbt.constant(%w[Cat Val When])
            )
          end
          c.next_state do |s, args|
            sname = s.active_sheet
            s.with(tables: s.tables.merge(sname => [{ name: args[:name], ref: args[:ref] }]))
          end
          c.run { |sut, args| sut.add_table(args[:ref], name: args[:name], columns: args[:columns]) }
        end

        m.command :add_filter_and_sort, weight: 2 do |c|
          c.guard { |s| !s.active_sheet.nil? && s.row_counts.fetch(s.active_sheet) >= 2 && !s.filters.fetch(s.active_sheet) }
          c.arguments do |s|
            last_r = s.row_counts.fetch(s.active_sheet)
            Pbt.fixed_hash(
              auto_filter: Pbt.constant("A1:C#{last_r}"),
              filter_col: Pbt.constant(0),
              filter_val: Pbt.constant("A"),
              sort_ref: Pbt.constant("A2:C#{last_r}"),
              cond_ref: Pbt.constant("B2:B#{last_r}")
            )
          end
          c.next_state { |s, _args| s.with(filters: s.filters.merge(s.active_sheet => true)) }
          c.run do |sut, args|
            sut.add_filter_and_sort(
              auto_filter: args[:auto_filter],
              filter_col: args[:filter_col],
              filter_val: args[:filter_val],
              sort_ref: args[:sort_ref],
              cond_ref: args[:cond_ref]
            )
          end
        end

        m.command :add_comment, weight: 2 do |c|
          c.guard { |s| !s.active_sheet.nil? && s.comments.fetch(s.active_sheet).size < 2 }
          c.arguments do |s|
            used = s.comments.fetch(s.active_sheet).map { |cm| cm[:ref] }
            avail = %w[A1 B1] - used
            Pbt.fixed_hash(ref: Pbt.constant_from(*avail), text: Pbt.alphanumeric_string(min: 1, max: 12))
          end
          c.next_state do |s, args|
            sname = s.active_sheet
            s.with(comments: s.comments.merge(sname => [*s.comments.fetch(sname), { ref: args[:ref], text: args[:text] }]))
          end
          c.run { |sut, args| sut.add_comment(args[:ref], args[:text]) }
        end

        m.command :add_defined_name, weight: 2 do |c|
          c.guard { |s| !s.active_sheet.nil? && s.defined_names.size < 4 }
          c.arguments do |s|
            idx = s.defined_names.size + 1
            Pbt.fixed_hash(
              name: Pbt.constant("NamedRange#{idx}"),
              range: Pbt.constant_from("A1:B2", "B1:C2"),
              sheet: Pbt.constant(s.active_sheet),
              scoped: Pbt.boolean
            )
          end
          c.next_state do |s, args|
            sname = args[:sheet]
            val = "#{Xlsxrb::DslHelpers.quote_sheet_name(sname)}!#{Xlsxrb::DslHelpers.absolute_range(args[:range])}"
            local_idx = args[:scoped] ? s.sheets.index(sname) : nil
            entry = { name: args[:name], value: val, local_sheet_id: local_idx }
            s.with(defined_names: [*s.defined_names, entry])
          end
          c.run do |sut, args|
            scope = args[:scoped] ? args[:sheet] : nil
            sut.add_defined_name(args[:name], range: args[:range], sheet: args[:sheet], scope: scope)
          end
        end

        m.invariant(:compound_parts_and_parity) do |state, sut|
          next if state.sheets.empty?

          mem_bytes, stream_bytes = sut.snapshot_bytes
          assert_empty ExcelRules.check(mem_bytes).map(&:to_s), "WorkbookBuilder violated ExcelRules"
          assert_empty ExcelRules.check(stream_bytes).map(&:to_s), "StreamWriter violated ExcelRules"
          assert_empty XsdValidator.check(mem_bytes)
          assert_empty XsdValidator.check(stream_bytes)

          expected_total_charts = state.charts.values.sum(&:size)
          [mem_bytes, stream_bytes].each do |bytes|
            entries = Xlsxrb::Ooxml::ZipReader.new(StringIO.new(bytes)).entry_names
            (1..expected_total_charts).each do |ci|
              assert_include entries, "xl/charts/chart#{ci}.xml"
            end
          end

          m_wb = Xlsxrb.load(StringIO.new(mem_bytes))
          s_wb = Xlsxrb.load(StringIO.new(stream_bytes))

          assert_equal state.date1904, m_wb.date1904?
          assert_equal state.date1904, s_wb.date1904?
          assert_equal state.sheets, m_wb.sheet_names
          assert_equal state.sheets, s_wb.sheet_names

          exp_dns = state.defined_names.map { |d| [d[:name], d[:value], d[:local_sheet_id]] }
          assert_equal(exp_dns, m_wb.defined_names.map { |d| [d.name, d.value, d.local_sheet_id] })
          assert_equal(exp_dns, s_wb.defined_names.map { |d| [d.name, d.value, d.local_sheet_id] })

          state.sheets.each_with_index do |sname, sidx|
            m_ws = m_wb.sheet(sidx)
            s_ws = s_wb.sheet(sidx)

            exp_charts = state.charts.fetch(sname).map { |ch| [ch[:type], ch[:title]] }
            assert_equal(exp_charts, m_ws.charts.map { |ch| [ch.type, ch.title] })
            assert_equal(exp_charts, s_ws.charts.map { |ch| [ch.type, ch.title] })

            exp_shapes = state.shapes.fetch(sname).map { |sh| [sh[:preset], sh[:text]] }
            assert_equal(exp_shapes, m_ws.shapes.map { |sh| [sh.preset, sh.text] })
            assert_equal(exp_shapes, s_ws.shapes.map { |sh| [sh.preset, sh.text] })

            assert_equal state.images.fetch(sname).size, m_ws.images.size
            assert_equal state.images.fetch(sname).size, s_ws.images.size
            assert_equal(m_ws.images.map(&:data), s_ws.images.map(&:data))

            exp_tables = state.tables.fetch(sname).map { |tb| [tb[:name], tb[:ref]] }
            assert_equal(exp_tables, m_ws.tables.map { |tb| [tb.name, tb.ref] })
            assert_equal(exp_tables, s_ws.tables.map { |tb| [tb.name, tb.ref] })

            exp_cmts = state.comments.fetch(sname).map { |cm| [cm[:ref], cm[:text]] }
            assert_equal(exp_cmts, m_ws.comments.map { |cm| [cm.ref, cm.text] })
            assert_equal(exp_cmts, s_ws.comments.map { |cm| [cm.ref, cm.text] })

            assert_equal m_ws.filter_columns, s_ws.filter_columns
            assert_equal m_ws.sort_state, s_ws.sort_state
            assert_equal(m_ws.columns.map { |col| [col.index, col.outline_level, col.collapsed?] },
                         s_ws.columns.map { |col| [col.index, col.outline_level, col.collapsed?] })

            exp_rows = state.rows_data.fetch(sname)
            assert_equal exp_rows.size, m_ws.rows.size
            assert_equal exp_rows.size, s_ws.rows.size
            exp_rows.each_with_index do |exp_r, r_idx|
              assert_equal m_ws.rows[r_idx].formatted_values, s_ws.rows[r_idx].formatted_values
              exp_r.each_with_index do |exp_v, c_idx|
                mc = m_ws[r_idx, c_idx]
                sc = s_ws[r_idx, c_idx]
                if exp_v.is_a?(Date)
                  assert_equal exp_v, mc.to_date
                  assert_equal exp_v, sc.to_date
                  assert_equal exp_v.strftime("%Y-%m-%d"), mc.formatted_value
                  assert_equal exp_v.strftime("%Y-%m-%d"), sc.formatted_value
                else
                  assert_equal exp_v, mc.value
                  assert_equal exp_v, sc.value
                  assert_equal mc.formatted_value, sc.formatted_value
                end
              end
            end
          end
        end
      end
    end
  end

  # Helper for MultiPartsSut#append_row inside command run block
  # --- Stateful PBT #2: Load -> #with immutable edit -> re-save -> reload lifecycle ---

  EditLifecycleState = Data.define(
    :use_stream,
    :rows_vals,
    :merged_cells,
    :auto_filter,
    :print_area,
    :print_titles_rows,
    :freeze_row,
    :zoom_scale,
    :tab_color,
    :comments,
    :hyperlinks,
    :row_breaks,
    :column_breaks,
    :core_title,
    :user_defined_names
  )

  # Holds a loaded Workbook and applies immutable #with updates and save/reload cycles.
  class EditLifecycleSut
    attr_accessor :wb

    def initialize(use_stream:)
      build_init = lambda do |w|
        w.core_properties(title: "InitTitle")
        w.defined_name("UserRange", range: "A1:B2", sheet: "EditSheet")
        w.sheet("EditSheet") do |s|
          s.freeze_pane(row: 1, column: 0)
          s.sheet_view(zoom_scale: 120)
          s.sheet_properties(tab_color: "FFFF0000")
          s.row(%w[H1 H2])
          s.row([10, 20])
          s.merge("A1:B1")
          s.auto_filter = "A1:B2"
          s.print_area = "A1:B2"
          s.print_titles(rows: "1:1")
          s.comment("A1", "init comment")
          s.hyperlink("B1", "https://example.com/init")
          s.row_break(0)
          s.column_break(0)
        end
      end
      bytes = use_stream ? Xlsxrb.write(&build_init) : Xlsxrb.write(Xlsxrb.build(&build_init))
      @wb = Xlsxrb.load(StringIO.new(bytes))
    end

    def update_sheet(**changes)
      ws = @wb.sheet(0).with(**changes)
      @wb = @wb.with(sheets: [ws])
      nil
    end

    def update_workbook(**changes)
      @wb = @wb.with(**changes)
      nil
    end

    def save_and_reload!
      bytes = @wb.to_data
      @wb = Xlsxrb.load(StringIO.new(bytes))
      bytes
    end
  end

  test "stateful PBT: load, immutable edit via #with, re-save, and reload fixed-point lifecycle" do
    init_arb = Pbt.boolean.map(
      lambda do |use_stream|
        EditLifecycleState.new(
          use_stream: use_stream,
          rows_vals: [%w[H1 H2], [10, 20]],
          merged_cells: ["A1:B1"],
          auto_filter: "A1:B2",
          print_area: "A1:B2",
          print_titles_rows: "1:1",
          freeze_row: 1,
          zoom_scale: 120,
          tab_color: "FFFF0000",
          comments: [["A1", "init comment"]],
          hyperlinks: { "B1" => "https://example.com/init" },
          row_breaks: [0],
          column_breaks: [0],
          core_title: "InitTitle",
          user_defined_names: [%w[UserRange 'EditSheet'!$A$1:$B$2]]
        )
      end,
      :use_stream.to_proc
    )

    Pbt.assert(num_runs: 15) do
      Pbt.stateful(max_steps: 14) do |m|
        m.initial_state init_arb
        m.sut { |init| EditLifecycleSut.new(use_stream: init.use_stream) }

        m.command :update_rows, weight: 3 do |c|
          c.arguments(v0: Pbt.integer(min: 1, max: 999), v1: Pbt.alphanumeric_string(min: 1, max: 8))
          c.next_state { |s, args| s.with(rows_vals: [%w[H1 H2], [args[:v0], args[:v1]]]) }
          c.run do |sut, args|
            new_rows = [
              Xlsxrb::Elements::Row.new(index: 0, cells: [
                                          Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: "H1"),
                                          Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 1, value: "H2")
                                        ]),
              Xlsxrb::Elements::Row.new(index: 1, cells: [
                                          Xlsxrb::Elements::Cell.new(row_index: 1, column_index: 0, value: args[:v0]),
                                          Xlsxrb::Elements::Cell.new(row_index: 1, column_index: 1, value: args[:v1])
                                        ])
            ]
            sut.update_sheet(rows: new_rows)
          end
        end

        m.command :edit_merges, weight: 2 do |c|
          c.arguments Pbt.constant_from([], ["A1:B1"], ["A2:B2"])
          c.next_state { |s, merges| s.with(merged_cells: merges) }
          c.run { |sut, merges| sut.update_sheet(merged_cells: merges) }
        end

        m.command :edit_auto_filter, weight: 2 do |c|
          c.arguments Pbt.constant_from(nil, "A1:B2")
          c.next_state { |s, af| s.with(auto_filter: af) }
          c.run { |sut, af| sut.update_sheet(auto_filter: af) }
        end

        m.command :edit_print_settings, weight: 2 do |c|
          c.arguments(
            print_area: Pbt.constant_from(nil, "A1:B2", "A1:A2"),
            print_titles_rows: Pbt.constant_from(nil, "1:1", "2:2")
          )
          c.next_state { |s, args| s.with(print_area: args[:print_area], print_titles_rows: args[:print_titles_rows]) }
          c.run do |sut, args|
            pt = args[:print_titles_rows] ? { rows: args[:print_titles_rows] } : nil
            sut.update_sheet(print_area: args[:print_area], print_titles: pt)
          end
        end

        m.command :edit_freeze_and_view, weight: 2 do |c|
          c.arguments(freeze_row: Pbt.constant_from(nil, 1), zoom_scale: Pbt.constant_from(80, 120, 150))
          c.next_state { |s, args| s.with(freeze_row: args[:freeze_row], zoom_scale: args[:zoom_scale]) }
          c.run do |sut, args|
            fp = args[:freeze_row] ? { row: args[:freeze_row], column: 0 } : nil
            sut.update_sheet(freeze_pane: fp, sheet_view: { zoom_scale: args[:zoom_scale] })
          end
        end

        m.command :edit_sheet_properties, weight: 2 do |c|
          c.arguments Pbt.constant_from("FFFF0000", "FF00FF00", "FF0000FF")
          c.next_state { |s, color| s.with(tab_color: color) }
          c.run { |sut, color| sut.update_sheet(sheet_properties: { tab_color: color }) }
        end

        m.command :edit_comments_and_links, weight: 2 do |c|
          c.arguments(clear: Pbt.boolean, comment_text: Pbt.alphanumeric_string(min: 1, max: 10))
          c.next_state do |s, args|
            if args[:clear]
              s.with(comments: [], hyperlinks: {})
            else
              s.with(comments: [["A1", args[:comment_text]]], hyperlinks: { "B1" => "https://example.com/#{args[:comment_text]}" })
            end
          end
          c.run do |sut, args|
            if args[:clear]
              sut.update_sheet(comments: [], hyperlinks: {})
            else
              sut.update_sheet(
                comments: [{ ref: "A1", text: args[:comment_text] }],
                hyperlinks: { "B1" => "https://example.com/#{args[:comment_text]}" }
              )
            end
          end
        end

        m.command :edit_breaks, weight: 2 do |c|
          c.arguments Pbt.boolean
          c.next_state do |s, clear|
            clear ? s.with(row_breaks: [], column_breaks: []) : s.with(row_breaks: [0], column_breaks: [0])
          end
          c.run do |sut, clear|
            if clear
              sut.update_sheet(row_breaks: [], column_breaks: [])
            else
              sut.update_sheet(row_breaks: [0], column_breaks: [0])
            end
          end
        end

        m.command :edit_workbook_meta, weight: 2 do |c|
          c.arguments(title: Pbt.alphanumeric_string(min: 1, max: 12), clear_dn: Pbt.boolean)
          c.next_state do |s, args|
            dns = args[:clear_dn] ? [] : [%w[UserRange 'EditSheet'!$A$1:$B$2]]
            s.with(core_title: args[:title], user_defined_names: dns)
          end
          c.run do |sut, args|
            dns = if args[:clear_dn]
                    []
                  else
                    [{ name: "UserRange", value: "'EditSheet'!$A$1:$B$2", local_sheet_id: 0 }]
                  end
            sut.update_workbook(core_properties: { title: args[:title] }, defined_names: dns)
          end
        end

        m.command :save_and_reload, weight: 4 do |c|
          c.run { |sut, _args| sut.save_and_reload! }
          c.verify { |ctx| assert_empty ExcelRules.check(ctx.result).map(&:to_s) }
        end

        m.invariant(:edit_lifecycle_and_fixed_point) do |state, sut|
          bytes1 = sut.wb.to_data
          assert_empty ExcelRules.check(bytes1).map(&:to_s)
          assert_empty XsdValidator.check(bytes1)
          wb1 = Xlsxrb.load(StringIO.new(bytes1))
          ws1 = wb1.sheet(0)

          assert_equal(state.rows_vals, ws1.rows.map { |r| r.cells.map(&:value) })
          assert_equal state.merged_cells, ws1.merged_cells
          assert_equal state.auto_filter, ws1.auto_filter
          assert_equal state.print_area, ws1.print_area
          assert_equal state.print_titles_rows, ws1.print_titles&.rows
          assert_equal state.freeze_row, ws1.freeze_pane&.row
          assert_equal state.zoom_scale, ws1.sheet_view.zoom_scale
          assert_equal state.tab_color, ws1.tab_color
          assert_equal(state.comments, ws1.comments.map { |cm| [cm.ref, cm.text] })
          assert_equal state.hyperlinks, ws1.hyperlinks.transform_values(&:url)
          assert_equal state.row_breaks, ws1.row_breaks.map(&:id)
          assert_equal state.column_breaks, ws1.column_breaks.map(&:id)
          assert_equal state.core_title, wb1.core_properties.title

          actual_user_dns = wb1.defined_names.reject { |dn| dn.name.start_with?("_xlnm.") }.map { |dn| [dn.name, dn.value] }
          assert_equal state.user_defined_names, actual_user_dns

          bytes2 = wb1.to_data
          assert_empty ExcelRules.check(bytes2).map(&:to_s)
          assert_empty XsdValidator.check(bytes2)
          wb2 = Xlsxrb.load(StringIO.new(bytes2))
          ws2 = wb2.sheet(0)

          assert_equal ws1, ws2
          assert_equal ws1.print_area, ws2.print_area
          assert_equal ws1.print_titles, ws2.print_titles
          assert_equal wb1.core_properties, wb2.core_properties
        end
      end
    end
  end

  # --- Stateful PBT #3: StreamWriter 64KB flush boundary & style precedence ---

  FlushStyleState = Data.define(
    :shared_strings,
    :default_font,
    :flushed,
    :row_cursor,
    :styles,
    :col_styles,
    :tracked_rows,
    :has_post_features
  )

  # Drives WorkbookBuilder and StreamWriter across the 64KB row-buffer flush boundary.
  class FlushStyleSut
    FLUSH_ROW_COUNT = 160

    attr_reader :stream_writer

    def initialize(shared_strings:, default_font: nil)
      @shared_strings = shared_strings
      @default_font = default_font
      @builder = Xlsxrb::WorkbookBuilder.new(default_font: @default_font)
      @stream_io = StringIO.new
      @stream_writer = Xlsxrb::StreamWriter.new(@stream_io, shared_strings: @shared_strings, default_font: @default_font)
      @b_sheet = nil
      @fiber = Fiber.new do
        @builder.sheet("FlushSheet") do |sh|
          @b_sheet = sh
          Fiber.yield
        end
      end
      @fiber.resume
      @stream_writer.sheet("FlushSheet")
      @s_proxy = @stream_writer[]
      @ops = []
      @row_idx = 0
    end

    def set_column_style(col_idx, style:)
      @b_sheet.column(col_idx, style: style)
      @s_proxy.column(col_idx, style: style)
      @ops << [:col, col_idx, style]
      nil
    end

    def define_style(name, **opts)
      @builder.style(name, **opts)
      @stream_writer.style(name, **opts)
      @ops << [:style, name, opts]
      nil
    end

    def flush_past_64kb
      FLUSH_ROW_COUNT.times do |i|
        vals = Array.new(20) { |c| @row_idx + i + c }
        @b_sheet.row(vals)
        @s_proxy.row(vals)
      end
      @row_idx += FLUSH_ROW_COUNT
      @ops << [:flush_rows, FLUSH_ROW_COUNT]
      @stream_writer.instance_variable_get(:@sheet_entry_started)
    end

    def append_styled_row(values, cell_styles:, row_style:)
      @b_sheet.row(values, cell_styles: cell_styles, row_style: row_style)
      @s_proxy.row(values, cell_styles: cell_styles, row_style: row_style)
      @row_idx += 1
      @ops << [:styled_row, values, cell_styles, row_style]
      nil
    end

    def add_post_features
      @b_sheet.merge("A1:B1")
      @s_proxy.merge("A1:B1")
      @b_sheet.auto_filter = "A1:C2"
      @s_proxy.auto_filter = "A1:C2"
      @b_sheet.comment("A1", "post_flush")
      @s_proxy.comment("A1", "post_flush")
      @ops << [:post_features]
      nil
    end

    def snapshot_bytes
      snap_builder = Xlsxrb::WorkbookBuilder.new(default_font: @default_font)
      snap_io = StringIO.new
      snap_stream = Xlsxrb::StreamWriter.new(snap_io, shared_strings: @shared_strings, default_font: @default_font)
      snap_b_sheet = nil
      snap_fiber = Fiber.new do
        snap_builder.sheet("FlushSheet") do |sh|
          snap_b_sheet = sh
          Fiber.yield
        end
      end
      snap_fiber.resume
      snap_stream.sheet("FlushSheet")
      snap_s_proxy = snap_stream[]
      r_idx = 0

      begin
        @ops.each do |op|
          case op[0]
          when :col
            snap_b_sheet.column(op[1], style: op[2])
            snap_s_proxy.column(op[1], style: op[2])
          when :style
            snap_builder.style(op[1], **op[2])
            snap_stream.style(op[1], **op[2])
          when :flush_rows
            op[1].times do |i|
              vals = Array.new(20) { |c| r_idx + i + c }
              snap_b_sheet.row(vals)
              snap_s_proxy.row(vals)
            end
            r_idx += op[1]
          when :styled_row
            snap_b_sheet.row(op[1], cell_styles: op[2], row_style: op[3])
            snap_s_proxy.row(op[1], cell_styles: op[2], row_style: op[3])
            r_idx += 1
          when :post_features
            snap_b_sheet.merge("A1:B1")
            snap_s_proxy.merge("A1:B1")
            snap_b_sheet.auto_filter = "A1:C2"
            snap_s_proxy.auto_filter = "A1:C2"
            snap_b_sheet.comment("A1", "post_flush")
            snap_s_proxy.comment("A1", "post_flush")
          end
        end

        snap_fiber.resume if snap_fiber.alive?
        snap_stream.close
        [Xlsxrb.write(snap_builder.build, shared_strings: @shared_strings), snap_io.string.b]
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
      if @fiber&.alive?
        begin
          @fiber.resume
        rescue StandardError
          nil
        end
      end
      @stream_writer.cleanup!
    end
  end

  def resolve_flush_cell_style(value, cell_style, row_style, col_style_spec, styles_map)
    spec = cell_style || row_style
    style_def = if spec.is_a?(Hash)
                  spec
                elsif spec
                  styles_map.fetch(spec)
                elsif col_style_spec
                  col_style_spec.is_a?(Hash) ? col_style_spec : styles_map.fetch(col_style_spec)
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

  test "stateful PBT: StreamWriter 64KB flush boundary with style redefinition, precedence, and date variants" do
    init_arb = Pbt.fixed_hash(
      shared_strings: Pbt.boolean,
      default_font: Pbt.constant_from(nil, { name: "Arial", size: 12 })
    ).map(
      lambda do |h|
        FlushStyleState.new(
          shared_strings: h[:shared_strings],
          default_font: h[:default_font],
          flushed: false,
          row_cursor: 0,
          styles: {
            s_a: { bold: true, italic: false, number_format: nil },
            s_b: { bold: false, italic: true, number_format: "#,##0" }
          },
          col_styles: {},
          tracked_rows: [],
          has_post_features: false
        )
      end,
      ->(s) { { shared_strings: s.shared_strings, default_font: s.default_font } }
    )

    Pbt.assert(num_runs: 10) do
      Pbt.stateful(max_steps: 10) do |m|
        m.initial_state init_arb
        m.sut do |init|
          sut = FlushStyleSut.new(shared_strings: init.shared_strings, default_font: init.default_font)
          init.styles.each { |name, opts| sut.define_style(name, **opts) }
          sut
        end
        m.teardown(&:cleanup!)

        m.command :set_column_style, weight: 2 do |c|
          c.guard { |s| s.row_cursor.zero? }
          c.arguments(col_idx: Pbt.constant_from(0, 1), style: Pbt.constant_from(:s_a, :s_b, { bold: true, italic: true }))
          c.next_state { |s, args| s.with(col_styles: s.col_styles.merge(args[:col_idx] => args[:style])) }
          c.run { |sut, args| sut.set_column_style(args[:col_idx], style: args[:style]) }
        end

        m.command :define_style, weight: 3 do |c|
          c.arguments(
            name: Pbt.constant_from(:s_a, :s_b),
            bold: Pbt.boolean,
            italic: Pbt.boolean,
            number_format: Pbt.constant_from(nil, "#,##0", "0.00")
          )
          c.next_state do |s, args|
            defn = { bold: args[:bold], italic: args[:italic], number_format: args[:number_format] }
            s.with(styles: s.styles.merge(args[:name] => defn))
          end
          c.run { |sut, args| sut.define_style(args[:name], bold: args[:bold], italic: args[:italic], number_format: args[:number_format]) }
        end

        m.command :flush_past_64kb, weight: 3 do |c|
          c.guard { |s| !s.flushed }
          c.next_state do |s|
            s.with(flushed: true, row_cursor: s.row_cursor + FlushStyleSut::FLUSH_ROW_COUNT)
          end
          c.run { |sut, _args| sut.flush_past_64kb }
          c.verify { |ctx| assert_true ctx.result }
        end

        m.command :append_styled_row, weight: 4 do |c|
          c.arguments(
            v0: Pbt.one_of(Pbt.integer(min: 1, max: 500), Pbt.alphanumeric_string(min: 1, max: 8), Pbt.constant(BigDecimal("-12.34"))),
            day_off: Pbt.integer(min: 0, max: 20),
            use_hash_cell_styles: Pbt.boolean,
            cs0: Pbt.constant_from(nil, :s_a, :s_b, { bold: true }),
            cs1: Pbt.constant_from(nil, :s_a, :s_b, { italic: true }),
            cs2: Pbt.constant_from(nil, :s_a, :s_b),
            row_style: Pbt.constant_from(nil, :s_a, :s_b, { italic: true })
          )
          c.next_state do |s, args|
            vals = [
              args[:v0],
              Date.new(2026, 6, 15) + args[:day_off],
              Time.utc(2026, 6, 15, 9, 30, 0),
              FakeTimeWithZone.new(2026, 6, 15, 9, 30, 0, 0, 32_400)
            ]
            styles_for_cells = [args[:cs0], args[:cs1], args[:cs2], args[:cs2]]
            exp_cells = vals.each_with_index.map do |v, ci|
              [v, resolve_flush_cell_style(v, styles_for_cells[ci], args[:row_style], s.col_styles[ci], s.styles)]
            end
            record = { row_index: s.row_cursor, cells: exp_cells }
            s.with(row_cursor: s.row_cursor + 1, tracked_rows: [*s.tracked_rows, record])
          end
          c.run do |sut, args|
            vals = [
              args[:v0],
              Date.new(2026, 6, 15) + args[:day_off],
              Time.utc(2026, 6, 15, 9, 30, 0),
              FakeTimeWithZone.new(2026, 6, 15, 9, 30, 0, 0, 32_400)
            ]
            cell_styles = if args[:use_hash_cell_styles]
                            h = {}
                            h[0] = args[:cs0] if args[:cs0]
                            h["B"] = args[:cs1] if args[:cs1]
                            h["C"] = args[:cs2] if args[:cs2]
                            h["D"] = args[:cs2] if args[:cs2]
                            h.empty? ? nil : h
                          elsif args[:cs0].nil? && args[:cs1].nil? && args[:cs2].nil?
                            nil
                          else
                            [args[:cs0], args[:cs1], args[:cs2], args[:cs2]]
                          end
            sut.append_styled_row(vals, cell_styles: cell_styles, row_style: args[:row_style])
          end
        end

        m.command :add_post_features, weight: 2 do |c|
          c.guard { |s| s.row_cursor >= 2 && !s.has_post_features }
          c.next_state { |s| s.with(has_post_features: true) }
          c.run { |sut, _args| sut.add_post_features }
        end

        m.invariant(:flush_boundary_style_parity) do |state, sut|
          next if state.tracked_rows.empty?

          mem_bytes, stream_bytes = sut.snapshot_bytes
          assert_empty ExcelRules.check(mem_bytes).map(&:to_s)
          assert_empty ExcelRules.check(stream_bytes).map(&:to_s)
          assert_empty XsdValidator.check(mem_bytes)
          assert_empty XsdValidator.check(stream_bytes)

          m_wb = Xlsxrb.load(StringIO.new(mem_bytes))
          s_wb = Xlsxrb.load(StringIO.new(stream_bytes))
          m_ws = m_wb.sheet(0)
          s_ws = s_wb.sheet(0)

          assert_equal state.row_cursor, m_ws.rows.size
          assert_equal state.row_cursor, s_ws.rows.size

          if state.has_post_features
            assert_equal ["A1:B1"], m_ws.merged_cells
            assert_equal ["A1:B1"], s_ws.merged_cells
            assert_equal "A1:C2", m_ws.auto_filter
            assert_equal "A1:C2", s_ws.auto_filter
            assert_equal([%w[A1 post_flush]], m_ws.comments.map { |cm| [cm.ref, cm.text] })
            assert_equal([%w[A1 post_flush]], s_ws.comments.map { |cm| [cm.ref, cm.text] })
          end

          assert_equal m_wb.styles.fonts.first&.name, s_wb.styles.fonts.first&.name
          assert_equal m_wb.styles.fonts.first&.size, s_wb.styles.fonts.first&.size

          state.tracked_rows.each do |rec|
            r_idx = rec[:row_index]
            assert_equal m_ws.rows[r_idx].formatted_values, s_ws.rows[r_idx].formatted_values
            rec[:cells].each_with_index do |(exp_val, exp_style), c_idx|
              mc = m_ws[r_idx, c_idx]
              sc = s_ws[r_idx, c_idx]
              case exp_val
              when Date
                assert_equal exp_val, mc.to_date
                assert_equal exp_val, sc.to_date
              when FakeTimeWithZone
                mt = mc.to_time
                st = sc.to_time
                exp_wall = [exp_val.year, exp_val.month, exp_val.day, exp_val.hour, exp_val.min, exp_val.sec]
                assert_equal exp_wall, [mt.year, mt.month, mt.day, mt.hour, mt.min, mt.sec]
                assert_equal exp_wall, [st.year, st.month, st.day, st.hour, st.min, st.sec]
              when Time
                assert_in_delta exp_val.to_i, mc.to_time.to_i, 1
                assert_in_delta exp_val.to_i, sc.to_time.to_i, 1
              when BigDecimal
                assert_in_delta exp_val.to_f, mc.value.to_f, 0.0001
                assert_in_delta exp_val.to_f, sc.value.to_f, 0.0001
              else
                assert_equal exp_val, mc.value
                assert_equal exp_val, sc.value
              end
              assert_equal mc.formatted_value, sc.formatted_value
              assert_equal exp_style, extract_loaded_cell_style(m_wb.styles, mc)
              assert_equal exp_style, extract_loaded_cell_style(s_wb.styles, sc)
            end
          end
        end
      end
    end
  end

  # --- Stateful PBT #4: Ooxml::ZipReader & EntryStreamIO state machine ---

  ZipReaderState = Data.define(:entries, :next_stream_id, :open_streams, :closed_streams)

  # Drives a single Ooxml::ZipReader and multiple interleaved EntryStreamIO handles.
  class ZipReaderSut
    attr_reader :reader, :expected_bytes

    def initialize(entries_spec)
      io = StringIO.new
      zw = Xlsxrb::Ooxml::ZipWriter.new(io, compression_level: 6)
      @expected_bytes = {}

      zw.instance_variable_set(:@compression_level, 0)
      zw.add_entry("small_stored.xml", entries_spec[:small])
      @expected_bytes["small_stored.xml"] = entries_spec[:small].b

      zw.instance_variable_set(:@compression_level, 6)
      zw.add_entry("large_deflated.xml", entries_spec[:large])
      @expected_bytes["large_deflated.xml"] = entries_spec[:large].b

      utf8_bom = ("\xEF\xBB\xBF".b + entries_spec[:bom8].b).b
      zw.add_binary_entry("bom_utf8.xml", utf8_bom)
      @expected_bytes["bom_utf8.xml"] = entries_spec[:bom8].b

      utf16_le = ("\xFF\xFE".b + entries_spec[:bom16].encode("UTF-16LE").b).b
      zw.add_binary_entry("bom_utf16le.xml", utf16_le)
      @expected_bytes["bom_utf16le.xml"] = entries_spec[:bom16].encode("UTF-8").b

      zw.close

      @reader = Xlsxrb::Ooxml::ZipReader.new(StringIO.new(io.string))
      @streams = {}
      @next_id = 0
    end

    def open_stream(name)
      id = @next_id
      @next_id += 1
      @streams[id] = @reader.open_entry_io(name)
      id
    end

    def read_chunk(stream_id, len)
      @streams.fetch(stream_id).read(len)
    end

    def direct_read(name)
      @reader.read_entry(name).b
    end

    def copy_entry(name)
      out = StringIO.new
      Xlsxrb::Ooxml::ZipWriter.open(out) { |zw| @reader.copy_to_writer(name, zw) }
      Xlsxrb::Ooxml::ZipReader.new(StringIO.new(out.string)).read_entry(name).b
    end

    def close_stream(stream_id)
      st = @streams.fetch(stream_id)
      st.close
      { closed: st.closed? }
    end

    def read_closed(stream_id)
      @streams.fetch(stream_id).read(4)
      nil
    rescue IOError => e
      e
    end

    def stream(stream_id)
      @streams.fetch(stream_id)
    end

    def cleanup!
      @streams.each_value do |st|
        st.close unless st.closed?
      rescue StandardError
        nil
      end
      @reader.close
    rescue StandardError
      nil
    end
  end

  test "stateful PBT: Ooxml::ZipReader and EntryStreamIO interleaved streaming, BOM decoding, and close transitions" do
    init_arb = Pbt.fixed_hash(
      small_txt: Pbt.alphanumeric_string(min: 20, max: 300),
      large_pad: Pbt.integer(min: 66_000, max: 72_000),
      bom8_txt: Pbt.alphanumeric_string(min: 10, max: 200),
      bom16_txt: Pbt.alphanumeric_string(min: 10, max: 200)
    ).map(
      lambda do |h|
        small_xml = "<root>#{h[:small_txt]}_日本語</root>"
        large_xml = "<root>#{"L" * h[:large_pad]}_end</root>"
        bom8_xml = "<bom8>#{h[:bom8_txt]}_漢字</bom8>"
        bom16_xml = "<bom16>#{h[:bom16_txt]}_絵文字</bom16>"
        ZipReaderState.new(
          entries: {
            "small_stored.xml" => small_xml.b,
            "large_deflated.xml" => large_xml.b,
            "bom_utf8.xml" => bom8_xml.b,
            "bom_utf16le.xml" => bom16_xml.b
          },
          next_stream_id: 0,
          open_streams: {},
          closed_streams: []
        )
      end,
      ->(_s) { { small_txt: "x", large_pad: 66_000, bom8_txt: "y", bom16_txt: "z" } }
    )

    Pbt.assert(num_runs: 20) do
      Pbt.stateful(max_steps: 18) do |m|
        m.initial_state init_arb
        m.sut do |init|
          ZipReaderSut.new(
            small: init.entries.fetch("small_stored.xml").dup.force_encoding("UTF-8"),
            large: init.entries.fetch("large_deflated.xml").dup.force_encoding("UTF-8"),
            bom8: init.entries.fetch("bom_utf8.xml").dup.force_encoding("UTF-8"),
            bom16: init.entries.fetch("bom_utf16le.xml").dup.force_encoding("UTF-8")
          )
        end
        m.teardown(&:cleanup!)

        m.command :open_stream, weight: 3 do |c|
          c.guard { |s| s.open_streams.size + s.closed_streams.size < 5 }
          c.arguments { |s| Pbt.constant_from(*s.entries.keys) }
          c.next_state do |s, name|
            sid = s.next_stream_id
            info = { name: name, pos: 0, eof_seen: false }
            s.with(next_stream_id: sid + 1, open_streams: s.open_streams.merge(sid => info))
          end
          c.run { |sut, name| sut.open_stream(name) }
        end

        m.command :read_chunk, weight: 6 do |c|
          c.guard { |s| !s.open_streams.empty? }
          c.arguments do |s|
            Pbt.fixed_hash(
              stream_id: Pbt.constant_from(*s.open_streams.keys),
              len: Pbt.constant_from(0, 1, 2, 7, 257, 4096, 70_000, nil)
            )
          end
          c.next_state do |s, args|
            sid = args[:stream_id]
            len = args[:len]
            info = s.open_streams.fetch(sid)
            total = s.entries.fetch(info[:name]).bytesize
            if len.nil?
              new_pos = total
              eof_seen = true
            elsif len.zero?
              new_pos = info[:pos]
              eof_seen = info[:eof_seen]
            else
              avail = total - info[:pos]
              if avail <= 0
                new_pos = info[:pos]
                eof_seen = true
              else
                read_len = [len, avail].min
                new_pos = info[:pos] + read_len
                eof_seen = info[:eof_seen]
              end
            end
            s.with(open_streams: s.open_streams.merge(sid => info.merge(pos: new_pos, eof_seen: eof_seen)))
          end
          c.run { |sut, args| [args[:stream_id], args[:len], sut.read_chunk(args[:stream_id], args[:len])] }
          c.verify do |ctx|
            sid, len, chunk = ctx.result
            old_info = ctx.before.open_streams.fetch(sid)
            expected = ctx.before.entries.fetch(old_info[:name])
            pos = old_info[:pos]
            if len.nil?
              exp = expected.byteslice(pos, expected.bytesize - pos) || "".b
              assert_equal exp, chunk.b
            elsif len.zero?
              assert_equal "".b, chunk.b
            elsif pos >= expected.bytesize
              assert_nil chunk
            else
              exp = expected.byteslice(pos, [len, expected.bytesize - pos].min)
              assert_equal exp, chunk.b
            end
          end
        end

        m.command :direct_read, weight: 2 do |c|
          c.arguments { |s| Pbt.constant_from(*s.entries.keys) }
          c.run { |sut, name| [name, sut.direct_read(name)] }
          c.verify { |ctx| assert_equal ctx.before.entries.fetch(ctx.result[0]), ctx.result[1] }
        end

        m.command :copy_to_writer, weight: 2 do |c|
          c.arguments { |s| Pbt.constant_from(*s.entries.keys) }
          c.run { |sut, name| [name, sut.copy_entry(name)] }
          c.verify { |ctx| assert_equal ctx.before.entries.fetch(ctx.result[0]), ctx.result[1] }
        end

        m.command :close_stream, weight: 2 do |c|
          c.guard { |s| !s.open_streams.empty? }
          c.arguments { |s| Pbt.constant_from(*s.open_streams.keys) }
          c.next_state do |s, sid|
            rest = s.open_streams.reject { |k, _| k == sid }
            s.with(open_streams: rest, closed_streams: [*s.closed_streams, sid])
          end
          c.run(&:close_stream)
          c.verify { |ctx| assert_true ctx.result[:closed] }
        end

        m.command :read_closed_stream, weight: 1 do |c|
          c.guard { |s| !s.closed_streams.empty? }
          c.arguments { |s| Pbt.constant_from(*s.closed_streams) }
          c.run(&:read_closed)
          c.verify { |ctx| assert_instance_of IOError, ctx.result }
        end

        m.invariant(:stream_states) do |state, sut|
          state.open_streams.each do |sid, info|
            st = sut.stream(sid)
            assert_false st.closed?
            total = state.entries.fetch(info[:name]).bytesize
            if info[:pos] < total
              assert_false st.eof?
            elsif info[:eof_seen]
              assert_true st.eof?
            end
          end

          state.closed_streams.each do |sid|
            assert_true sut.stream(sid).closed?
          end
        end
      end
    end
  end

  # --- Stateful PBT #5: StreamSheet lazy metadata & option switching ---

  StreamSheetState = Data.define(:use_stream, :lead_skip, :data_rows, :mid_skip, :tail_empty, :verified_ops)

  # Holds an open Xlsxrb.read Fiber yielding StreamSheet alongside a reference Worksheet from Xlsxrb.load.
  class StreamSheetSut
    attr_reader :stream_sheet, :loaded_ws

    def initialize(spec)
      build_proc = lambda do |wb|
        wb.sheet("Sparse") do |s|
          s.column(0, width: 16)
          s.column(1, width: 22, best_fit: true)
          s.skip_rows(spec.lead_skip) if spec.lead_skip.positive?
          spec.data_rows.each_with_index do |r, idx|
            s.row(r)
            s.skip_rows(spec.mid_skip) if idx.zero? && spec.mid_skip.positive?
          end
          spec.tail_empty.times { s.row([nil, nil]) }
          s.merge("A1:B1")
          s.auto_filter = "A1:C4"
          s.print_area = "A1:C4"
          s.comment("A1", "lazy_comment")
          s.hyperlink("B1", "https://example.com/lazy")
        end
      end

      bytes = spec.use_stream ? Xlsxrb.write(&build_proc) : Xlsxrb.write(Xlsxrb.build(&build_proc))
      @loaded_ws = Xlsxrb.load(StringIO.new(bytes)).sheet(0)
      @stream_sheet = nil
      @fiber = Fiber.new do
        Xlsxrb.read(StringIO.new(bytes)) do |ss|
          @stream_sheet = ss
          Fiber.yield
        end
      end
      @fiber.resume
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

  test "stateful PBT: StreamSheet lazy metadata access, partial iteration, and row option switching" do
    init_arb = Pbt.fixed_hash(
      use_stream: Pbt.boolean,
      lead_skip: Pbt.integer(min: 0, max: 2),
      mid_skip: Pbt.integer(min: 0, max: 2),
      tail_empty: Pbt.integer(min: 0, max: 2),
      n1: Pbt.integer(min: -50, max: 500),
      str1: Pbt.alphanumeric_string(min: 1, max: 8)
    ).map(
      lambda do |h|
        rows = [
          [nil, h[:n1], Date.new(2026, 6, 15)],
          [h[:str1], nil, true]
        ]
        StreamSheetState.new(
          use_stream: h[:use_stream],
          lead_skip: h[:lead_skip],
          data_rows: rows,
          mid_skip: h[:mid_skip],
          tail_empty: h[:tail_empty],
          verified_ops: 0
        )
      end,
      ->(s) { { use_stream: s.use_stream, lead_skip: s.lead_skip, mid_skip: s.mid_skip, tail_empty: s.tail_empty, n1: 1, str1: "a" } }
    )

    Pbt.assert(num_runs: 15) do
      Pbt.stateful(max_steps: 14) do |m|
        m.initial_state init_arb
        m.sut { |init| StreamSheetSut.new(init) }
        m.teardown(&:cleanup!)

        m.command :iterate_full, weight: 4 do |c|
          c.arguments(trim: Pbt.boolean, pad_r: Pbt.boolean, pad_c: Pbt.boolean)
          c.next_state { |s, _| s.with(verified_ops: s.verified_ops + 1) }
          c.run do |sut, args|
            opts = { trim_empty_rows: args[:trim], pad_empty_rows: args[:pad_r], pad_empty_cells: args[:pad_c] }
            s_rows = sut.stream_sheet.each_row(**opts).map { |r| [r.index, r.cells.map { |cell| [cell.column_index, cell.value] }] }
            l_rows = sut.loaded_ws.each_row(**opts).map { |r| [r.index, r.cells.map { |cell| [cell.column_index, cell.value] }] }
            [s_rows, l_rows]
          end
          c.verify { |ctx| assert_equal ctx.result[1], ctx.result[0] }
        end

        m.command :iterate_values, weight: 3 do |c|
          c.arguments(type_cast: Pbt.boolean, trim: Pbt.boolean, pad_r: Pbt.boolean, pad_c: Pbt.boolean)
          c.next_state { |s, _| s.with(verified_ops: s.verified_ops + 1) }
          c.run do |sut, args|
            opts = { type_cast: args[:type_cast], trim_empty_rows: args[:trim], pad_empty_rows: args[:pad_r], pad_empty_cells: args[:pad_c] }
            [sut.stream_sheet.each_row_values(**opts).to_a, sut.loaded_ws.each_row_values(**opts).to_a]
          end
          c.verify { |ctx| assert_equal ctx.result[1], ctx.result[0] }
        end

        m.command :iterate_partial_break, weight: 3 do |c|
          c.arguments Pbt.integer(min: 0, max: 2)
          c.next_state { |s, _| s.with(verified_ops: s.verified_ops + 1) }
          c.run do |sut, take_n|
            partial = []
            sut.stream_sheet.each_row do |r|
              break if partial.size >= take_n

              partial << [r.index, r.values]
            end
            expected = sut.loaded_ws.each_row.first(take_n).map { |r| [r.index, r.values] }
            [partial, expected]
          end
          c.verify { |ctx| assert_equal ctx.result[1], ctx.result[0] }
        end

        m.command :access_metadata, weight: 4 do |c|
          c.arguments Pbt.constant_from(:columns, :comments, :hyperlinks, :auto_filter, :merged_cells, :print_area)
          c.next_state { |s, _| s.with(verified_ops: s.verified_ops + 1) }
          c.run do |sut, field|
            case field
            when :columns
              [
                sut.stream_sheet.columns.map { |col| [col.index, col.width, col.best_fit?] },
                sut.loaded_ws.columns.map { |col| [col.index, col.width, col.best_fit?] }
              ]
            when :comments
              [
                sut.stream_sheet.comments.map { |cm| [cm.ref, cm.text] },
                sut.loaded_ws.comments.map { |cm| [cm.ref, cm.text] }
              ]
            when :hyperlinks
              [
                sut.stream_sheet.hyperlinks.transform_values(&:url),
                sut.loaded_ws.hyperlinks.transform_values(&:url)
              ]
            when :auto_filter
              [sut.stream_sheet.auto_filter, sut.loaded_ws.auto_filter]
            when :merged_cells
              [sut.stream_sheet.merged_cells, sut.loaded_ws.merged_cells]
            when :print_area
              [sut.stream_sheet.print_area, sut.loaded_ws.print_area]
            end
          end
          c.verify { |ctx| assert_equal ctx.result[1], ctx.result[0] }
        end

        m.command :materialize_load, weight: 2 do |c|
          c.next_state { |s| s.with(verified_ops: s.verified_ops + 1) }
          c.run do |sut|
            mat = sut.stream_sheet.load
            [
              [mat.rows.map { |r| [r.index, r.values] }, mat.columns, mat.comments, mat.hyperlinks, mat.auto_filter, mat.merged_cells, mat.print_area],
              [sut.loaded_ws.rows.map { |r| [r.index, r.values] }, sut.loaded_ws.columns, sut.loaded_ws.comments, sut.loaded_ws.hyperlinks, sut.loaded_ws.auto_filter, sut.loaded_ws.merged_cells, sut.loaded_ws.print_area]
            ]
          end
          c.verify { |ctx| assert_equal ctx.result[1], ctx.result[0] }
        end

        m.invariant(:stream_sheet_default_parity) do |_state, sut|
          assert_equal sut.loaded_ws.name, sut.stream_sheet.name
          assert_equal(sut.loaded_ws.each_row.map { |r| [r.index, r.values] },
                       sut.stream_sheet.each_row.map { |r| [r.index, r.values] })
        end
      end
    end
  end

  # --- Stateful PBT #6: StyleBuilder fluent mutation, copy isolation, and StylesXml roundtrip ---

  StyleBuilderState = Data.define(:builders, :next_id)

  # Tracks live StyleBuilder instances and serializes them through StylesXml and StylesParser.
  class StyleBuilderSut
    attr_reader :builders

    def initialize
      @builders = { b0: Xlsxrb::StyleBuilder.new(:b0) }
    end

    def mutate_font(id, args)
      sb = @builders.fetch(id)
      sb.bold = args[:bold]
      sb.italic = args[:italic]
      sb.strike = args[:strike]
      sb.underline = args[:underline]
      case args[:va]
      when :super then sb.superscript = true
      when :sub then sb.subscript = true
      else sb.vert_align = :baseline
      end
      sb.font_size = args[:font_size]
      sb.font_name = args[:font_name]
      sb.font_color = args[:font_color]
      nil
    end

    def mutate_fill(id, mode, color1, color2)
      sb = @builders.fetch(id)
      case mode
      when :solid
        sb.fill_color = color1
      when :pattern
        sb.fill_pattern("gray125", fg_color: color1, bg_color: color2)
      when :gradient
        sb.fill_gradient(type: "linear", degree: 90, stops: [{ position: 0, color: color1 }, { position: 1, color: color2 }])
      end
      nil
    end

    def mutate_border(id, all_style, all_color, diag_style, diag_up)
      sb = @builders.fetch(id)
      sb.border_all(style: all_style, color: all_color)
      sb.border_diagonal(style: diag_style, color: all_color, up: diag_up, down: !diag_up)
      nil
    end

    def mutate_alignment_and_prot(id, args)
      sb = @builders.fetch(id)
      sb.align_horizontal = args[:h]
      sb.align_vertical = args[:v]
      sb.wrap_text = args[:wrap]
      sb.text_rotation = args[:rot]
      sb.indent = args[:indent]
      sb.shrink_to_fit = args[:shrink]
      sb.locked = args[:locked]
      sb.hidden = args[:hidden]
      sb.number_format = args[:num_fmt]
      nil
    end

    def fork_copy(src_id, new_id, with_fmt:)
      src = @builders.fetch(src_id)
      copy = with_fmt ? src.with_number_format(new_id.to_s, with_fmt) : src.dup
      @builders[new_id] = copy
      nil
    end

    def roundtrip_styles
      writer = Xlsxrb::Ooxml::Writer.new
      xf_by_id = @builders.transform_values { |sb| sb.register_with(writer) }
      xml = writer.send(:generate_styles_xml)
      styles = Xlsxrb::Ooxml::StylesParser.parse(xml)
      [xf_by_id, styles]
    end
  end

  test "stateful PBT: StyleBuilder fluent mutation, copy isolation, and StylesXml roundtrip" do
    init_model = {
      bold: false, italic: false, strike: false, underline: nil, vert_align: nil,
      font_size: nil, font_name: nil, font_color: nil,
      fill_mode: nil, fill_color: nil, fill_pattern: nil, fill_bg: nil, fill_grad_stops: nil,
      border_style: nil, border_color: nil,
      align_h: nil, align_v: nil, wrap_text: false, text_rotation: nil, indent: nil, shrink_to_fit: false,
      locked: true, hidden: false, number_format: nil
    }

    Pbt.assert(num_runs: 20) do
      Pbt.stateful(max_steps: 14) do |m|
        m.initial_state { StyleBuilderState.new(builders: { b0: init_model }, next_id: 1) }
        m.sut { StyleBuilderSut.new }

        m.command :mutate_font, weight: 3 do |c|
          c.arguments do |s|
            Pbt.fixed_hash(
              id: Pbt.constant_from(*s.builders.keys),
              bold: Pbt.boolean,
              italic: Pbt.boolean,
              strike: Pbt.boolean,
              underline: Pbt.constant_from(false, true, :double, :single_accounting),
              va: Pbt.constant_from(:super, :sub, :baseline),
              font_size: Pbt.constant_from(10, 11.5, 14),
              font_name: Pbt.constant_from("Arial", "Calibri"),
              font_color: Pbt.constant_from(:red, "#00FF00", "#00F")
            )
          end
          c.next_state do |s, args|
            cur = s.builders.fetch(args[:id])
            exp_u = case args[:underline]
                    when false then nil
                    when true then "single"
                    when :single_accounting then "singleAccounting"
                    else args[:underline].to_s
                    end
            exp_va = case args[:va]
                     when :super then "superscript"
                     when :sub then "subscript"
                     else "baseline"
                     end
            exp_fc = Xlsxrb::StyleBuilder.resolve_color(args[:font_color])
            updated = cur.merge(
              bold: args[:bold],
              italic: args[:italic],
              strike: args[:strike],
              underline: exp_u,
              vert_align: exp_va,
              font_size: args[:font_size],
              font_name: args[:font_name],
              font_color: exp_fc
            )
            s.with(builders: s.builders.merge(args[:id] => updated))
          end
          c.run { |sut, args| sut.mutate_font(args[:id], args) }
        end

        m.command :mutate_fill, weight: 3 do |c|
          c.arguments do |s|
            Pbt.fixed_hash(
              id: Pbt.constant_from(*s.builders.keys),
              mode: Pbt.constant_from(:solid, :pattern, :gradient),
              color1: Pbt.constant_from("FFFF0000", "FF00FF00"),
              color2: Pbt.constant_from("FF0000FF", "FFFFFF00")
            )
          end
          c.next_state do |s, args|
            cur = s.builders.fetch(args[:id])
            updated = case args[:mode]
                      when :solid
                        cur.merge(fill_mode: :solid, fill_color: args[:color1], fill_pattern: "solid", fill_bg: nil, fill_grad_stops: nil)
                      when :pattern
                        cur.merge(fill_mode: :pattern, fill_color: args[:color1], fill_pattern: "gray125", fill_bg: args[:color2], fill_grad_stops: nil)
                      when :gradient
                        cur.merge(fill_mode: :gradient, fill_color: nil, fill_pattern: nil, fill_bg: nil, fill_grad_stops: [args[:color1], args[:color2]])
                      end
            s.with(builders: s.builders.merge(args[:id] => updated))
          end
          c.run { |sut, args| sut.mutate_fill(args[:id], args[:mode], args[:color1], args[:color2]) }
        end

        m.command :mutate_border, weight: 3 do |c|
          c.arguments do |s|
            Pbt.fixed_hash(
              id: Pbt.constant_from(*s.builders.keys),
              all_style: Pbt.constant_from(:thin, :medium_dashed, 2, "double"),
              all_color: Pbt.constant_from(:navy, "#112233"),
              diag_style: Pbt.constant_from(:thin, :dotted),
              diag_up: Pbt.boolean
            )
          end
          c.next_state do |s, args|
            cur = s.builders.fetch(args[:id])
            norm_s = Xlsxrb::StyleBuilder.border_style!(args[:all_style])
            norm_c = Xlsxrb::StyleBuilder.resolve_color(args[:all_color])
            s.with(builders: s.builders.merge(args[:id] => cur.merge(border_style: norm_s, border_color: norm_c)))
          end
          c.run { |sut, args| sut.mutate_border(args[:id], args[:all_style], args[:all_color], args[:diag_style], args[:diag_up]) }
        end

        m.command :mutate_alignment_and_prot, weight: 2 do |c|
          c.arguments do |s|
            Pbt.fixed_hash(
              id: Pbt.constant_from(*s.builders.keys),
              h: Pbt.constant_from("left", "center", "right"),
              v: Pbt.constant_from("top", "center", "bottom"),
              wrap: Pbt.boolean,
              rot: Pbt.constant_from(0, 45, 90),
              indent: Pbt.integer(min: 0, max: 3),
              shrink: Pbt.boolean,
              locked: Pbt.boolean,
              hidden: Pbt.boolean,
              num_fmt: Pbt.constant_from("#,##0.00", "0.0%", "yyyy-mm-dd")
            )
          end
          c.next_state do |s, args|
            cur = s.builders.fetch(args[:id])
            updated = cur.merge(
              align_h: args[:h], align_v: args[:v], wrap_text: args[:wrap],
              text_rotation: args[:rot], indent: args[:indent], shrink_to_fit: args[:shrink],
              locked: args[:locked], hidden: args[:hidden], number_format: args[:num_fmt]
            )
            s.with(builders: s.builders.merge(args[:id] => updated))
          end
          c.run { |sut, args| sut.mutate_alignment_and_prot(args[:id], args) }
        end

        m.command :fork_copy, weight: 3 do |c|
          c.guard { |s| s.builders.size < 4 }
          c.arguments do |s|
            Pbt.fixed_hash(
              src_id: Pbt.constant_from(*s.builders.keys),
              new_id: Pbt.constant(:"b#{s.next_id}"),
              with_fmt: Pbt.constant_from(nil, "0.000", "yyyy/mm/dd")
            )
          end
          c.next_state do |s, args|
            src = s.builders.fetch(args[:src_id])
            copy = args[:with_fmt] ? src.merge(number_format: args[:with_fmt]) : src.dup
            s.with(builders: s.builders.merge(args[:new_id] => copy), next_id: s.next_id + 1)
          end
          c.run { |sut, args| sut.fork_copy(args[:src_id], args[:new_id], with_fmt: args[:with_fmt]) }
        end

        m.invariant(:builder_isolation_and_styles_xml_roundtrip) do |state, sut|
          assert_equal state.builders.keys, sut.builders.keys
          state.builders.each do |id, exp|
            sb = sut.builders.fetch(id)
            assert_equal exp[:bold], sb.bold?
            assert_equal exp[:italic], sb.italic?
            assert_equal exp[:strike], sb.strike?
            assert_equal exp[:underline], sb.underline
            assert_equal exp[:vert_align], sb.vert_align
            assert_equal exp[:font_size], sb.font_size
            assert_equal exp[:font_name], sb.font_name
            assert_equal exp[:font_color], sb.font_color
            assert_equal exp[:align_h], sb.align_horizontal
            assert_equal exp[:align_v], sb.align_vertical
            assert_equal exp[:wrap_text], sb.wrap_text?
            assert_equal exp[:text_rotation], sb.text_rotation
            assert_equal exp[:indent], sb.indent
            assert_equal exp[:shrink_to_fit], sb.shrink_to_fit?
            assert_equal exp[:locked], sb.locked?
            assert_equal exp[:hidden], sb.hidden?
            assert_equal exp[:number_format], sb.number_format
            assert_equal exp[:fill_pattern], sb.fill_pattern if exp[:fill_mode] && exp[:fill_mode] != :gradient
            assert_equal exp[:fill_color], sb.fill_color if exp[:fill_mode] == :solid
          end

          xf_by_id, styles = sut.roundtrip_styles
          state.builders.each do |id, exp|
            xf_idx = xf_by_id.fetch(id)
            xf = styles.cell_xf(xf_idx)
            assert_not_nil xf

            if exp[:font_name]
              font = styles.font(xf.font_id)
              assert_equal exp[:bold], font.bold?
              assert_equal exp[:italic], font.italic?
              assert_equal exp[:strike], font.strike?
              assert_equal exp[:underline], font.underline
              assert_equal exp[:vert_align], font.vert_align
              assert_equal exp[:font_size].to_f, font.size
              assert_equal exp[:font_name], font.name
              assert_equal({ rgb: exp[:font_color] }, font.color)
            end

            if exp[:fill_mode]
              fill = styles.fill(xf.fill_id)
              case exp[:fill_mode]
              when :solid
                assert_nil fill.gradient
                assert_equal "solid", fill.pattern_type
                assert_equal({ rgb: exp[:fill_color] }, fill.fg_color)
              when :pattern
                assert_nil fill.gradient
                assert_equal exp[:fill_pattern], fill.pattern_type
                assert_equal({ rgb: exp[:fill_color] }, fill.fg_color)
                assert_equal({ rgb: exp[:fill_bg] }, fill.bg_color)
              when :gradient
                assert_not_nil fill.gradient
                assert_equal(exp[:fill_grad_stops].map { |c| { rgb: c } }, fill.gradient[:stops].map { |st| st[:color] })
              end
            end

            if exp[:border_style]
              border = styles.border(xf.border_id)
              assert_equal({ style: exp[:border_style], color: { rgb: exp[:border_color] } }, border.left)
              assert_equal({ style: exp[:border_style], color: { rgb: exp[:border_color] } }, border.bottom)
            end

            next unless exp[:align_h]

            align = xf.alignment
            assert_equal exp[:align_h], align.horizontal
            assert_equal exp[:align_v], align.vertical
            assert_equal exp[:wrap_text], align.wrap_text?
            assert_equal exp[:text_rotation], align.text_rotation
            assert_equal exp[:indent], align.indent
            assert_equal exp[:shrink_to_fit], align.shrink_to_fit?
            assert_equal exp[:locked], xf.protection.locked?
            assert_equal exp[:hidden], xf.protection.hidden?
            assert_equal exp[:number_format], styles.number_format(xf_idx)
          end
        end
      end
    end
  end
end
