# frozen_string_literal: true

require "test_helper"
require "pbt"
require "tempfile"
require "date"
require "time"

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
      sz: sz,
      color: colors[color_idx],
      fill_color: colors[fill_idx],
      horizontal: horizontals[h_idx],
      vertical: verticals[v_idx],
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

          read_wb = Xlsxrb.read(tmp.path).load

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
                  assert_equal expected_expr, actual_formula
                  assert_equal expected_val.cached_value, actual_val unless expected_val.cached_value.nil?
                when Hash
                  expected_expr = expected_val[:formula].delete_prefix("=")
                  assert_equal expected_expr, actual_formula
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
                  if expected_val.start_with?("=")
                    expected_expr = expected_val.delete_prefix("=")
                    assert_equal expected_expr, actual_formula
                  else
                    assert_equal expected_val, actual_val.to_s
                  end
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
          streaming_wb = Xlsxrb.read(streaming_path).load
          in_memory_wb = Xlsxrb.read(in_memory_path).load

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
        letter = Xlsxrb::Elements::Cell.column_letter(col_idx)
        assert_equal col_idx, Xlsxrb::Elements::Cell.column_index(letter)
        assert_equal col_idx, Xlsxrb::Elements::Cell.column_index(letter.downcase)
        assert_equal col_idx, Xlsxrb::Elements::Cell.column_index(letter.to_sym)

        # 2. Reference bijection: parse_ref <-> ref
        ref = "#{letter}#{row_idx + 1}"
        parsed = Xlsxrb::Elements::Cell.parse_ref(ref)
        assert_equal [row_idx, col_idx], parsed

        parsed_lower = Xlsxrb::Elements::Cell.parse_ref(ref.downcase)
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
        assert_equal val, cell.content
        assert_equal val, cell[:value]
        assert_equal cell.ref, cell[:ref]
        assert_equal row_idx, cell[:row_index]
        assert_equal col_idx, cell[:column_index]

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
        when Xlsxrb::Elements::RichText
          assert_equal val.to_s, cell.to_s
        when true, false
          assert_equal "b", cell[:type]
          assert_equal val.to_s, cell.to_s
        when String
          assert_equal "s", cell[:type]
          assert_equal val, cell.to_s
        when nil
          assert_nil cell[:type]
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
        assert_equal row_hidden, row.hidden
        assert_equal row_custom_height, row.custom_height
        assert_equal row_outline, row.outline_level
        assert_equal row_height, row[:height]
        assert_equal row_hidden, row[:hidden]

        assert_equal cells[0], row[0]
        assert_equal cells[1], row[1]
        assert_nil row[2]
        assert_equal cells[0], row.cell_at(0)
        assert_nil row.cell_at(1)
        assert_equal cells[1], row.cell_at(2)

        assert_equal ["A", nil, "C"], row.to_a
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
        assert_equal col_hidden, col.hidden
        assert_equal col_custom_width, col.custom_width
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
      ) do |sheet_name, cell_specs, update_val|
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

        # Invariant 1: Random coordinate access
        unique_cells.each do |(r, c), expected_val|
          ref = "#{Xlsxrb::Elements::Cell.column_letter(c)}#{r + 1}"
          cell = sheet[ref]
          assert_not_nil cell
          assert_equal expected_val, cell.value
        end

        # Invariant 2: Immutable update_cell
        target_r, target_c = unique_cells.keys.first
        target_ref = "#{Xlsxrb::Elements::Cell.column_letter(target_c)}#{target_r + 1}"
        original_val = unique_cells[[target_r, target_c]]

        updated_sheet = sheet.update_cell(target_ref, value: update_val)

        # Original worksheet is not modified
        assert_equal original_val, sheet[target_ref].value
        # Updated worksheet has new value
        assert_equal update_val, updated_sheet[target_ref].value

        # Other cells remain identical
        unique_cells.each do |(r, c), expected_val|
          next if r == target_r && c == target_c

          ref = "#{Xlsxrb::Elements::Cell.column_letter(c)}#{r + 1}"
          assert_equal expected_val, updated_sheet[ref].value
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
                s.row([label, num], styles: %i[custom_style custom_style])
              end
            end
          end

          assert_true File.exist?(xlsx_path)

          # Invariant: File must load and parse valid styles without XML parse errors
          wb = Xlsxrb.read(xlsx_path).load
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
              s.freeze_pane(row: opts[:freeze_row], col: opts[:freeze_col]) if opts[:freeze_row].positive? || opts[:freeze_col].positive?
              s.sheet_view(:zoom_scale, opts[:zoom_scale])
              s.sheet_properties(:tab_color, opts[:tab_color])
              s.protect_sheet(sheet: opts[:protect_sheet], password: opts[:password])
              s.row(%w[Col1 Col2 Col3])
              s.row([100, 200, 300])
            end
          end

          # Invariant: Complex options roundtrip and load cleanly without XML corruption
          wb = Xlsxrb.read(xlsx_path).load
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
      invalid_float_arb = Pbt.one_of(Float::NAN, Float::INFINITY, -Float::INFINITY)
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
        Pbt.constant("?"), Pbt.constant("/"), Pbt.constant("\\")
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
      # rubocop:disable Style/SelectByRegexp
      valid_sheet_arb = Pbt.printable_ascii_string(min: 1, max: 31).filter { |s| !s.match?(%r{[\[\]*?/\\]}) }
      # rubocop:enable Style/SelectByRegexp
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
        mode = mode_idx.zero? ? :standard : :agile

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
          wb = Xlsxrb.read(xlsx_path, password: password).load
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
            wb = Xlsxrb.read(corrupted, password: password).load
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
          streaming_data = Xlsxrb.read(streaming_path, password: password).load.sheets[0].map(&:to_a)
          in_memory_data = Xlsxrb.read(in_memory_path, password: password).load.sheets[0].map(&:to_a)

          assert_equal streaming_data, in_memory_data
        end
      end
    end
  end
end
