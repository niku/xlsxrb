# frozen_string_literal: true

require "test_helper"
require "pbt"
require "stringio"
require "date"
require_relative "../support/excel_rules"
require_relative "../support/xsd_validator"

class DslDrawingStylesPbtTest < Test::Unit::TestCase
  def setup
    omit "Skip property-based tests under RBS runtime testing to prevent OOM / slowdown" if defined?(RBS::Test)
  end

  # --- Item 4: Conditional Formatting Stateful PBT ---

  CfState = Data.define(:sheets, :active_sheet_idx, :dirty)

  # Drives WorkbookBuilder and StreamWriter with conditional formatting rules.
  class CfSut
    CF_TARGETS = {
      sqref_a: { sqref: "A2:C5" }.freeze,
      sqref_b: { sqref: "B2:D5" }.freeze,
      coords: { columns: 0..1, rows: 1..5, except_rows: [2] }.freeze
    }.freeze

    VISUAL_RULES = {
      color_scale_two: { sqref: "A2:A6", type: :color_scale, colors: %w[F8696B 63BE7B] }.freeze,
      color_scale_three: { sqref: "B2:B6", type: :color_scale, colors: %w[F8696B FFEB84 63BE7B] }.freeze,
      data_bar: { sqref: "C2:C6", type: :data_bar, color: "FF5A8DD4" }.freeze,
      icon_set_three: { sqref: "D2:D6", type: :icon_set, icon_style: :three_arrows }.freeze,
      icon_set_four: { sqref: "A2:D2", type: :icon_set, icon_style: :four_rating }.freeze,
      icon_set_five: { sqref: "A3:D3", type: :icon_set, icon_style: :five_quarters }.freeze
    }.freeze

    def self.build_dxf_rule_spec(args)
      base = case args[:rule_kind]
             when :cell_is_gt then { type: :cell_is, operator: :greater_than, formula: "10" }
             when :cell_is_between then { type: :cell_is, operator: :between, formula: %w[5 50] }
             when :expression then { type: :expression, formula: "A2>10" }
             when :top10 then { type: :top10, rank: 5, percent: args[:flag_a], bottom: args[:flag_b] }
             when :above_average then { type: :above_average, above_average: args[:flag_a], equal_average: args[:flag_b] }
             when :time_period then { type: :time_period, time_period: (args[:flag_a] ? :today : :this_month) }
             when :contains_text then { type: :contains_text, text: "Alert" }
             when :duplicate_values then { type: :duplicate_values }
             end
      style_opts = case args[:style_mode]
                   when :flat_a
                     { fill_color: "FFCCCC", font_color: "990000", bold: true }
                   when :flat_b
                     { fill_color: "CCFFCC", italic: true }
                   when :dxf_hash
                     { dxf: { font: { bold: true, color: "FF0000FF" }, fill: { pattern: "solid", fg_color: "FFFFFF99" } } }
                   when :block_dsl
                     { fill_color: "FFCCCC", bold: true, stop_if_true: args[:flag_a], _block: true }
                   end
      base.merge(style_opts).merge(CF_TARGETS.fetch(args[:target_mode]))
    end

    def self.build_visual_rule_spec(args)
      VISUAL_RULES.fetch(args[:kind])
    end

    def initialize
      @sheets = [{ name: "Sheet1", rules: [] }]
      @active_sheet_idx = 0
    end

    def add_sheet
      @sheets << { name: "Sheet#{@sheets.size + 1}", rules: [] }
      @active_sheet_idx = @sheets.size - 1
      nil
    end

    def add_rule(spec)
      @sheets[@active_sheet_idx][:rules] << spec
      nil
    end

    def snapshot_bytes
      wb = Xlsxrb.build do |b|
        @sheets.each { |sh| b.sheet(sh[:name]) { |s| populate_cf_sheet(s, sh[:rules]) } }
      end
      mem_bytes = Xlsxrb.write(wb)

      stream_bytes = Xlsxrb.write do |sw|
        @sheets.each { |sh| sw.sheet(sh[:name]) { |s| populate_cf_sheet(s, sh[:rules]) } }
      end
      [mem_bytes, stream_bytes]
    end

    private

    def populate_cf_sheet(sheet, rules)
      sheet.row(%w[A B C D])
      5.times { |i| sheet.row([i * 5, "Alert_#{i}", 10 + i, 100 - i]) }

      rules.each do |spec|
        if spec[:_block]
          clean = spec.except(:_block, :sqref, :columns, :rows, :except_rows)
          coord_opts = spec.slice(:columns, :rows, :except_rows)
          sheet.conditional_format(spec[:sqref], **coord_opts) do |cf|
            cf.type = clean[:type]
            cf.operator = clean[:operator] if clean[:operator]
            cf.formula = clean[:formula] if clean[:formula]
            cf.rank = clean[:rank] if clean[:rank]
            cf.percent = clean[:percent] unless clean[:percent].nil?
            cf.bottom = clean[:bottom] unless clean[:bottom].nil?
            cf.above_average = clean[:above_average] unless clean[:above_average].nil?
            cf.equal_average = clean[:equal_average] unless clean[:equal_average].nil?
            cf.time_period = clean[:time_period] if clean[:time_period]
            cf.text = clean[:text] if clean[:text]
            cf.fill_color = clean[:fill_color]
            cf.bold = clean[:bold]
            cf.stop_if_true = clean[:stop_if_true]
          end
        else
          sqref = spec[:sqref]
          opts = spec.except(:sqref)
          sheet.conditional_format(sqref, **opts)
        end
      end
    end
  end

  def normalize_cf_dxf(styles, dxf_id)
    return nil if dxf_id.nil?

    dxf = styles.dxf(dxf_id)
    return :missing if dxf.nil?

    font_sig = ([dxf.font.bold?, dxf.font.italic?, dxf.font.color&.[](:rgb)] if dxf.font)
    fill_sig = ([dxf.fill.pattern_type, dxf.fill.fg_color&.[](:rgb), dxf.fill.bg_color&.[](:rgb)] if dxf.fill)
    [font_sig, fill_sig]
  end

  def normalize_cf_list(worksheet, styles)
    entries = worksheet.conditional_formats.map do |cf|
      [
        cf.sqref,
        cf.type,
        cf.operator,
        cf.formula,
        cf.rank,
        cf.percent?,
        cf.bottom?,
        cf.equal_average?,
        cf.time_period,
        cf.text,
        cf.stop_if_true?,
        !cf.color_scale.nil?,
        !cf.data_bar.nil?,
        cf.icon_set&.[](:icon_set),
        normalize_cf_dxf(styles, cf.dxf_id)
      ]
    end
    entries.sort_by { |entry| [entry[0].to_s, entry[1].to_s, entry[2].to_s, entry[4].to_i, entry[9].to_s, entry[8].to_s] }
  end

  test "stateful property: conditional formatting rules, block DSL, dxf dedup, and coordinate targeting" do
    init_state = CfState.new(sheets: [{ name: "Sheet1", rules: [] }], active_sheet_idx: 0, dirty: false)

    Pbt.assert(num_runs: 12) do
      Pbt.stateful(max_steps: 9) do |m|
        m.initial_state Pbt.constant(init_state)
        m.sut { |_init| CfSut.new }

        m.command :add_sheet, weight: 2 do |c|
          c.guard { |s| s.sheets.size < 2 }
          c.next_state do |s|
            new_sheet = { name: "Sheet#{s.sheets.size + 1}", rules: [] }
            s.with(sheets: [*s.sheets, new_sheet], active_sheet_idx: s.sheets.size, dirty: true)
          end
          c.run { |sut, _args| sut.add_sheet }
        end

        m.command :add_dxf_rule, weight: 4 do |c|
          c.guard { |s| s.sheets[s.active_sheet_idx][:rules].size < 5 }
          c.arguments(
            rule_kind: Pbt.constant_from(:cell_is_gt, :cell_is_between, :expression, :top10, :above_average, :time_period, :contains_text, :duplicate_values),
            style_mode: Pbt.constant_from(:flat_a, :flat_b, :dxf_hash, :block_dsl),
            target_mode: Pbt.constant_from(:sqref_a, :sqref_b, :coords),
            flag_a: Pbt.boolean,
            flag_b: Pbt.boolean
          )
          c.next_state do |s, args|
            spec = CfSut.build_dxf_rule_spec(args)
            sheets = s.sheets.each_with_index.map do |sh, i|
              i == s.active_sheet_idx ? sh.merge(rules: [*sh[:rules], spec]) : sh
            end
            s.with(sheets: sheets, dirty: true)
          end
          c.run { |sut, args| sut.add_rule(CfSut.build_dxf_rule_spec(args)) }
        end

        m.command :add_visual_rule, weight: 3 do |c|
          c.guard { |s| s.sheets[s.active_sheet_idx][:rules].size < 5 }
          c.arguments(
            kind: Pbt.constant_from(:color_scale_two, :color_scale_three, :data_bar, :icon_set_three, :icon_set_four, :icon_set_five)
          )
          c.next_state do |s, args|
            spec = CfSut.build_visual_rule_spec(args)
            sheets = s.sheets.each_with_index.map do |sh, i|
              i == s.active_sheet_idx ? sh.merge(rules: [*sh[:rules], spec]) : sh
            end
            s.with(sheets: sheets, dirty: true)
          end
          c.run { |sut, args| sut.add_rule(CfSut.build_visual_rule_spec(args)) }
        end

        m.invariant(:cf_parity_and_roundtrip) do |state, sut|
          next unless state.dirty

          mem_bytes, stream_bytes = sut.snapshot_bytes
          assert_empty(ExcelRules.check(mem_bytes).map(&:to_s))
          assert_empty(ExcelRules.check(stream_bytes).map(&:to_s))
          assert_empty(XsdValidator.check(mem_bytes))
          assert_empty(XsdValidator.check(stream_bytes))

          m_wb = Xlsxrb.load(StringIO.new(mem_bytes))
          s_wb = Xlsxrb.load(StringIO.new(stream_bytes))

          resaved = m_wb.to_data
          assert_empty(ExcelRules.check(resaved).map(&:to_s))
          assert_empty(XsdValidator.check(resaved))
          r_wb = Xlsxrb.load(StringIO.new(resaved))

          assert_equal(m_wb.styles.dxfs.size, s_wb.styles.dxfs.size, "dxf dedup count mismatch between build and stream")
          assert_equal(m_wb.styles.dxfs.size, r_wb.styles.dxfs.size, "dxf count mismatch after to_data")

          stream_cfs = []
          Xlsxrb.read(StringIO.new(mem_bytes)) do |stream_sheet|
            stream_cfs << normalize_cf_list(stream_sheet, m_wb.styles)
          end

          state.sheets.each_with_index do |_sh, idx|
            m_norm = normalize_cf_list(m_wb.sheet(idx), m_wb.styles)
            s_norm = normalize_cf_list(s_wb.sheet(idx), s_wb.styles)
            r_norm = normalize_cf_list(r_wb.sheet(idx), r_wb.styles)
            assert_equal(m_norm, s_norm, "CF parity mismatch on sheet #{idx}")
            assert_equal(m_norm, r_norm, "CF roundtrip mismatch on sheet #{idx}")
            assert_equal(m_norm, stream_cfs[idx], "CF StreamSheet mismatch on sheet #{idx}")
          end
        end
      end
    end
  end

  # --- Item 5a: DslHelpers run compression & chunked sqrefs algebraic invariants ---

  def expand_sqref_coords(sqref_chunks)
    coords = []
    sqref_chunks.each do |chunk|
      tokens = chunk.split
      assert_operator(tokens.size, :>=, 1)
      assert_operator(tokens.size, :<=, Xlsxrb::MAX_REFS_PER_VALIDATION)
      tokens.each do |token|
        start_ref, end_ref = token.split(":", 2)
        sr, sc = Xlsxrb::Elements::Cell._parse_ref(start_ref)
        er, ec = end_ref ? Xlsxrb::Elements::Cell._parse_ref(end_ref) : [sr, sc]
        (sr..er).each do |r|
          (sc..ec).each { |c| coords << [r, c] }
        end
      end
    end
    coords
  end

  test "property: DslHelpers compress_runs, complement_runs, and chunked_sqrefs algebraic invariants" do
    gen = Pbt.tuple(
      Pbt.array(Pbt.integer(min: 0, max: 35), min: 1, max: 25),
      Pbt.array(Pbt.integer(min: 0, max: 35), min: 0, max: 15),
      Pbt.array(Pbt.integer(min: 0, max: 12), min: 1, max: 10),
      Pbt.array(Pbt.integer(min: 0, max: 25), min: 1, max: 12)
    )

    Pbt.assert(num_runs: 80) do
      Pbt.property(gen) do |(indices, excl, col_indices, row_indices)|
        runs = Xlsxrb::DslHelpers.compress_runs(indices)
        assert_equal(indices.uniq.sort, runs.flat_map { |a, b| (a..b).to_a })
        runs.each { |a, b| assert_operator(a, :<=, b) }
        runs.each_cons(2) { |(_a1, b1), (a2, _b2)| assert_operator(b1 + 1, :<, a2) }

        lo = indices.min
        hi = indices.max
        comp = Xlsxrb::DslHelpers.complement_runs(excl, lo, hi)
        assert_equal((lo..hi).to_a - excl, comp.flat_map { |a, b| (a..b).to_a })
        comp.each { |a, b| assert_operator(a, :<=, b) }
        comp.each_cons(2) { |(_a1, b1), (a2, _b2)| assert_operator(b1 + 1, :<, a2) }

        col_runs = Xlsxrb::DslHelpers.compress_runs(col_indices)
        row_runs = Xlsxrb::DslHelpers.compress_runs(row_indices)
        sqrefs = Xlsxrb::DslHelpers.chunked_sqrefs(col_runs, row_runs)
        expanded = expand_sqref_coords(sqrefs)
        expected = row_indices.uniq.sort.product(col_indices.uniq.sort)
        assert_equal(expected.size, expanded.uniq.size)
        assert_equal(expected.sort, expanded.sort)
      end
    end
  end

  # --- Item 5b: Data Validations Stateful PBT ---

  DvState = Data.define(:validations, :dirty)

  # Drives WorkbookBuilder and StreamWriter with data validations across all OOXML types.
  class DvSut
    DV_BASES = {
      list_shortcut: { list: %w[High Medium Low] }.freeze,
      list_formula: { type: :list, formula1: '"Alpha,Beta,Gamma"' }.freeze,
      whole_between: { type: :whole, operator: :between, formula1: "1", formula2: "100" }.freeze,
      decimal_gt: { type: :decimal, operator: :greaterThan, formula1: "0.5" }.freeze,
      date_between: { type: :date, operator: :between, formula1: "45000", formula2: "46000" }.freeze,
      time_lt: { type: :time, operator: :lessThan, formula1: "0.5" }.freeze,
      text_length: { type: :textLength, operator: :lessThanOrEqual, formula1: "20" }.freeze,
      custom: { type: :custom, formula1: "A2>0" }.freeze,
      block_dsl: { type: :whole, operator: :greaterThan, formula1: "10", _block: true }.freeze
    }.freeze

    def self.build_dv_spec(slot, args)
      base = DV_BASES.fetch(args[:kind])
      target = if args[:use_coords]
                 { columns: 0..1, rows: 1..4, except_rows: [2] }
               else
                 col_l = Xlsxrb::Elements::Cell._column_letter(slot % 4)
                 { sqref: "#{col_l}2:#{col_l}5" }
               end
      msgs = if args[:with_multiline]
               {
                 prompt_title: "Input\nTitle",
                 prompt: "Line 1\nLine 2\r\nLine 3",
                 error_title: "Err\nHead",
                 error: "Bad value\nTry again & <fix>"
               }
             else
               {}
             end
      base.merge(target).merge(allow_blank: args[:allow_blank], error_style: args[:error_style]).merge(msgs)
    end

    def initialize
      @validations = []
    end

    def add_validation(spec)
      @validations << spec
      nil
    end

    def snapshot_bytes
      wb = Xlsxrb.build { |b| b.sheet("S") { |s| populate_dv_sheet(s) } }
      mem_bytes = Xlsxrb.write(wb)
      stream_bytes = Xlsxrb.write { |sw| sw.sheet("S") { |s| populate_dv_sheet(s) } }
      [mem_bytes, stream_bytes]
    end

    private

    def populate_dv_sheet(sheet)
      sheet.row(%w[A B C D])
      4.times { |i| sheet.row([i + 1, i + 2, i + 3, i + 4]) }
      @validations.each do |spec|
        if spec[:_block]
          coord_opts = spec.slice(:columns, :rows, :except_rows)
          sheet.data_validation(spec[:sqref], **coord_opts) do |dv|
            dv.type = spec[:type]
            dv.operator = spec[:operator]
            dv.formula1 = spec[:formula1]
            dv.allow_blank = spec[:allow_blank]
            dv.error_style = spec[:error_style]
            if spec[:prompt_title]
              dv.prompt_title = spec[:prompt_title]
              dv.prompt_message = spec[:prompt]
              dv.error_title = spec[:error_title]
              dv.error_message = spec[:error]
            end
          end
        else
          sheet.data_validation(spec[:sqref], **spec.except(:sqref))
        end
      end
    end
  end

  def normalize_dvs(worksheet)
    worksheet.data_validations.map do |dv|
      [
        dv.sqref,
        dv.type,
        dv.operator,
        dv.formula1,
        dv.formula2,
        dv.allow_blank?,
        dv.error_style,
        dv.show_input_message?,
        dv.show_error_message?,
        dv.prompt_title,
        dv.prompt,
        dv.error_title,
        dv.error
      ]
    end
  end

  test "stateful property: data validations across all types, multiline messages, and 4-way roundtrip" do
    init_state = DvState.new(validations: [], dirty: false)

    Pbt.assert(num_runs: 15) do
      Pbt.stateful(max_steps: 7) do |m|
        m.initial_state Pbt.constant(init_state)
        m.sut { |_init| DvSut.new }

        m.command :add_validation, weight: 5 do |c|
          c.guard { |s| s.validations.size < 6 }
          c.arguments(
            kind: Pbt.constant_from(:list_shortcut, :list_formula, :whole_between, :decimal_gt, :date_between, :time_lt, :text_length, :custom, :block_dsl),
            use_coords: Pbt.boolean,
            allow_blank: Pbt.boolean,
            error_style: Pbt.constant_from("stop", "warning", "information"),
            with_multiline: Pbt.boolean
          )
          c.next_state do |s, args|
            spec = DvSut.build_dv_spec(s.validations.size, args)
            s.with(validations: [*s.validations, spec], dirty: true)
          end
          c.run { |sut, args| sut.add_validation(DvSut.build_dv_spec(sut.instance_variable_get(:@validations).size, args)) }
        end

        m.invariant(:dv_four_way_roundtrip) do |state, sut|
          next unless state.dirty

          mem_bytes, stream_bytes = sut.snapshot_bytes
          assert_empty(ExcelRules.check(mem_bytes).map(&:to_s))
          assert_empty(ExcelRules.check(stream_bytes).map(&:to_s))
          assert_empty(XsdValidator.check(mem_bytes))
          assert_empty(XsdValidator.check(stream_bytes))

          m_wb = Xlsxrb.load(StringIO.new(mem_bytes))
          s_wb = Xlsxrb.load(StringIO.new(stream_bytes))
          m_norm = normalize_dvs(m_wb.sheet(0))
          s_norm = normalize_dvs(s_wb.sheet(0))

          assert_equal(state.validations.size, m_norm.size)
          assert_equal(m_norm, s_norm)

          m_stream_dvs = nil
          Xlsxrb.read(StringIO.new(mem_bytes)) { |ss| m_stream_dvs = normalize_dvs(ss) }
          assert_equal(m_norm, m_stream_dvs)

          resaved = m_wb.to_data
          assert_empty(ExcelRules.check(resaved).map(&:to_s))
          assert_empty(XsdValidator.check(resaved))
          r_norm = normalize_dvs(Xlsxrb.load(StringIO.new(resaved)).sheet(0))
          assert_equal(m_norm, r_norm)
        end
      end
    end
  end

  # --- Item 6: ChartBuilder + DrawingXml / DrawingListeners & Rich Shape Stateful PBT ---

  DrawingState = Data.define(:charts, :shapes, :dirty)

  # Drives WorkbookBuilder and StreamWriter with charts and rich shapes.
  class DrawingSut
    NO_AXIS = %i[pie doughnut].freeze
    TRENDLINE_TYPES = %i[bar col line scatter area].freeze

    def self.build_chart_spec(slot, args)
      ctype = args[:chart_type]
      series = if ctype == :bubble
                 [{ name: "Bubbles", cat_ref: "ChartData!$B$2:$B$5", val_ref: "ChartData!$C$2:$C$5", bubble_size_ref: "ChartData!$D$2:$D$5" }]
               elsif ctype == :combo
                 [
                   { type: :bar, name: "Rev", cat_ref: "ChartData!$A$2:$A$5", val_ref: "ChartData!$B$2:$B$5" },
                   { type: :line, name: "Qty", cat_ref: "ChartData!$A$2:$A$5", val_ref: "ChartData!$C$2:$C$5", secondary_axis: true }
                 ]
               else
                 s0 = { name: "S1", cat_ref: "ChartData!$A$2:$A$5", val_ref: "ChartData!$B$2:$B$5" }
                 s0[:trendline] = { type: "linear" } if args[:with_trendline] && TRENDLINE_TYPES.include?(ctype)
                 list = [s0]
                 list << { name: "S2", cat_ref: "ChartData!$A$2:$A$5", val_ref: "ChartData!$C$2:$C$5" } if args[:two_series]
                 list
               end
      spec = {
        type: ctype,
        use_block: args[:use_block] && ctype != :combo,
        title: "Q1 & Q2 <#{ctype}>",
        series: series,
        legend: { position: args[:legend_pos] },
        data_labels: (args[:with_data_labels] ? { show_val: true } : nil),
        from_column: 5,
        from_row: slot * 12,
        to_column: 12,
        to_row: (slot * 12) + 10
      }
      unless NO_AXIS.include?(ctype)
        spec[:category_axis] = { title: "Cat" }
        spec[:value_axis] = { title: "Val" }
      end
      spec[:secondary_value_axis] = { title: "Units" } if ctype == :combo
      spec.compact
    end

    def self.build_shape_spec(slot, args)
      spec = {
        preset: args[:preset],
        text: (args[:multiline_text] ? "Header\nDetail & <Info>" : "Box #{slot}"),
        fill_color: args[:fill_color],
        from_column: 0,
        from_row: 6 + (slot * 5),
        to_column: 3,
        to_row: 10 + (slot * 5)
      }
      if args[:with_line]
        spec[:line_color] = "000000"
        spec[:line_width] = 19_050
      end
      spec[:outer_shadow] = { blur_rad: 20_000, dist: 30_000, color: "000000" } if args[:with_shadow]
      spec
    end

    def initialize
      @charts = []
      @shapes = []
    end

    def add_chart(spec)
      @charts << spec
      nil
    end

    def add_shape(spec)
      @shapes << spec
      nil
    end

    def snapshot_bytes
      wb = Xlsxrb.build { |b| b.sheet("ChartData") { |s| populate_drawing_sheet(s) } }
      mem_bytes = Xlsxrb.write(wb)
      stream_bytes = Xlsxrb.write { |sw| sw.sheet("ChartData") { |s| populate_drawing_sheet(s) } }
      [mem_bytes, stream_bytes]
    end

    private

    def populate_drawing_sheet(sheet)
      sheet.row(%w[Cat X Y Size])
      sheet.row(["Q1", 10, 20, 5])
      sheet.row(["Q2", 20, 25, 8])
      sheet.row(["Q3", 30, 30, 12])
      sheet.row(["Q4", 40, 35, 15])

      @charts.each do |ch|
        if ch[:use_block]
          anchor_opts = ch.slice(:from_column, :from_row, :to_column, :to_row)
          sheet.chart(ch[:type], **anchor_opts) do |c|
            c.title = ch[:title]
            c.legend = ch[:legend] if ch[:legend]
            c.data_labels = ch[:data_labels] if ch[:data_labels]
            c.category_axis = ch[:category_axis] if ch[:category_axis]
            c.value_axis = ch[:value_axis] if ch[:value_axis]
            ch[:series].each do |sr|
              if ch[:type] == :bubble
                c.series(**sr)
              else
                c.series do |sb|
                  sb.name = sr[:name]
                  sb.categories = sr[:cat_ref]
                  sb.values = sr[:val_ref]
                  sb.trendline = sr[:trendline] if sr[:trendline]
                end
              end
            end
          end
        else
          sheet.chart(**ch.except(:use_block))
        end
      end

      @shapes.each { |sh| sheet.shape(**sh) }
    end
  end

  def normalize_charts(worksheet)
    worksheet.charts.map do |ch|
      [ch.type, ch.title, ch.series&.size, ch.legend&.[](:position), ch.from_column, ch.from_row, ch.to_column, ch.to_row]
    end
  end

  def normalize_shapes(worksheet)
    worksheet.shapes.map do |sh|
      [sh.preset, sh.text, sh.from, sh.to]
    end
  end

  test "stateful property: charts across all types, axes, trendlines, and rich shapes parity and roundtrip" do
    init_state = DrawingState.new(charts: [], shapes: [], dirty: false)

    Pbt.assert(num_runs: 12) do
      Pbt.stateful(max_steps: 6) do |m|
        m.initial_state Pbt.constant(init_state)
        m.sut { |_init| DrawingSut.new }

        m.command :add_chart, weight: 4 do |c|
          c.guard { |s| s.charts.size < 3 }
          c.arguments(
            chart_type: Pbt.constant_from(:bar, :col, :line, :pie, :doughnut, :scatter, :area, :radar, :bubble, :combo),
            use_block: Pbt.boolean,
            two_series: Pbt.boolean,
            legend_pos: Pbt.constant_from("b", "r", "t"),
            with_data_labels: Pbt.boolean,
            with_trendline: Pbt.boolean
          )
          c.next_state do |s, args|
            spec = DrawingSut.build_chart_spec(s.charts.size, args)
            s.with(charts: [*s.charts, spec], dirty: true)
          end
          c.run { |sut, args| sut.add_chart(DrawingSut.build_chart_spec(sut.instance_variable_get(:@charts).size, args)) }
        end

        m.command :add_shape, weight: 3 do |c|
          c.guard { |s| s.shapes.size < 3 }
          c.arguments(
            preset: Pbt.constant_from("rect", "roundRect", "ellipse", "triangle"),
            fill_color: Pbt.constant_from("4F81BD", "C0504D", "9BBB59"),
            with_line: Pbt.boolean,
            with_shadow: Pbt.boolean,
            multiline_text: Pbt.boolean
          )
          c.next_state do |s, args|
            spec = DrawingSut.build_shape_spec(s.shapes.size, args)
            s.with(shapes: [*s.shapes, spec], dirty: true)
          end
          c.run { |sut, args| sut.add_shape(DrawingSut.build_shape_spec(sut.instance_variable_get(:@shapes).size, args)) }
        end

        m.invariant(:drawing_parity_and_roundtrip) do |state, sut|
          next unless state.dirty

          mem_bytes, stream_bytes = sut.snapshot_bytes
          assert_empty(ExcelRules.check(mem_bytes).map(&:to_s))
          assert_empty(ExcelRules.check(stream_bytes).map(&:to_s))
          assert_empty(XsdValidator.check(mem_bytes))
          assert_empty(XsdValidator.check(stream_bytes))

          m_wb = Xlsxrb.load(StringIO.new(mem_bytes))
          s_wb = Xlsxrb.load(StringIO.new(stream_bytes))
          m_ws = m_wb.sheet(0)
          s_ws = s_wb.sheet(0)

          assert_equal(state.charts.size, m_ws.charts.size)
          assert_equal(normalize_charts(m_ws), normalize_charts(s_ws))
          assert_equal(state.shapes.size, m_ws.shapes.size)
          assert_equal(normalize_shapes(m_ws), normalize_shapes(s_ws))

          Xlsxrb.read(StringIO.new(mem_bytes)) do |ss|
            assert_equal(normalize_charts(m_ws), normalize_charts(ss))
            assert_equal(normalize_shapes(m_ws), normalize_shapes(ss))
          end

          resaved = m_wb.to_data
          assert_empty(ExcelRules.check(resaved).map(&:to_s))
          assert_empty(XsdValidator.check(resaved))
          r_ws = Xlsxrb.load(StringIO.new(resaved)).sheet(0)
          assert_equal(normalize_charts(m_ws), normalize_charts(r_ws))
          assert_equal(normalize_shapes(m_ws), normalize_shapes(r_ws))
        end
      end
    end
  end

  # --- Item 7: Nested outlines + remaining sheet metadata + Worksheet#with Stateful PBT ---

  SheetMetaState = Data.define(
    :nested_outlines,
    :col_groups,
    :outline_props,
    :split_pane,
    :selection,
    :header_footer,
    :page_margins,
    :print_options,
    :sheet_protection,
    :protected_ranges,
    :ignored_errors,
    :cell_watches,
    :phonetic_props,
    :code_name,
    :fit_to_page,
    :sheet_calc_pr,
    :dirty
  )

  # Drives WorkbookBuilder and StreamWriter with outlines and remaining sheet metadata.
  class SheetMetaSut
    def initialize
      @config = {}
    end

    def apply_outlines(nested_outlines:, col_groups:, outline_props:)
      @config[:nested_outlines] = nested_outlines
      @config[:col_groups] = col_groups
      @config[:outline_props] = outline_props
      nil
    end

    def apply_split_and_selection(split_pane:, selection:)
      @config[:split_pane] = split_pane
      @config[:selection] = selection
      nil
    end

    def apply_print_and_hf(header_footer:, page_margins:, print_options:, fit_to_page:)
      @config[:header_footer] = header_footer
      @config[:page_margins] = page_margins
      @config[:print_options] = print_options
      @config[:fit_to_page] = fit_to_page
      nil
    end

    def apply_protection(sheet_protection:, protected_ranges:)
      @config[:sheet_protection] = sheet_protection
      @config[:protected_ranges] = protected_ranges
      nil
    end

    def apply_misc_meta(ignored_errors:, cell_watches:, phonetic_props:, code_name:, sheet_calc_pr:)
      @config[:ignored_errors] = ignored_errors
      @config[:cell_watches] = cell_watches
      @config[:phonetic_props] = phonetic_props
      @config[:code_name] = code_name
      @config[:sheet_calc_pr] = sheet_calc_pr
      nil
    end

    def snapshot_bytes
      wb = Xlsxrb.build { |b| b.sheet("Meta") { |s| populate_meta_sheet(s) } }
      mem_bytes = Xlsxrb.write(wb)
      stream_bytes = Xlsxrb.write { |sw| sw.sheet("Meta") { |s| populate_meta_sheet(s) } }
      [mem_bytes, stream_bytes]
    end

    private

    def populate_meta_sheet(sheet)
      # Pre-row sheet settings must precede rows to keep StreamWriter ordering valid.
      sheet.outline_properties(**@config[:outline_props]) if @config[:outline_props]
      if @config[:col_groups]
        cg = @config[:col_groups]
        sheet.group_columns(1..2, level: 1, hidden: cg[:hidden], collapsed: cg[:collapsed])
      end
      sheet.split_pane(**@config[:split_pane]) if @config[:split_pane]
      if @config[:selection]
        sel = @config[:selection]
        sheet.selection(sel[:active_cell], sqref: sel[:sqref], pane: sel[:pane])
      end
      sheet.code_name = @config[:code_name] if @config[:code_name]
      sheet.fit_to_page = @config[:fit_to_page] unless @config[:fit_to_page].nil?
      sheet.sheet_calc_properties = @config[:sheet_calc_pr] if @config[:sheet_calc_pr]

      sheet.row(%w[H1 H2 H3])
      if @config[:nested_outlines]
        no = @config[:nested_outlines]
        sheet.group_rows(level: 1) do
          sheet.row(["L1-A", 1, 2])
          sheet.group_rows(level: 2, hidden: no[:hidden], collapsed: no[:collapsed]) do
            sheet.row(["L2-A", 3, 4])
            sheet.row(["L2-B", 5, 6])
          end
        end
      else
        sheet.row(["R1", 10, 20])
      end

      sheet.header_footer(**@config[:header_footer]) if @config[:header_footer]
      sheet.page_margins(**@config[:page_margins]) if @config[:page_margins]
      sheet.print_options(**@config[:print_options]) if @config[:print_options]
      sheet.sheet_protection(**@config[:sheet_protection]) if @config[:sheet_protection]
      @config[:protected_ranges]&.each do |pr|
        sheet.protected_range(pr[:name], pr[:sqref], password: pr[:password])
      end
      @config[:ignored_errors]&.each do |ie|
        sheet.ignored_error(ie[:sqref], **ie.except(:sqref))
      end
      @config[:cell_watches]&.each { |ref| sheet.cell_watch(ref) }
      sheet.phonetic_properties = @config[:phonetic_props] if @config[:phonetic_props]
    end
  end

  def extract_sheet_meta_snapshot(worksheet)
    row_outlines = if worksheet.respond_to?(:rows)
                     worksheet.rows.map { |r| [r.index, r.outline_level, r.hidden?, r.collapsed?] }
                   else
                     []
                   end
    {
      row_outlines: row_outlines,
      col_outlines: worksheet.columns.map { |c| [c.index, c.outline_level, c.hidden?, c.collapsed?] },
      outline_props: worksheet.outline_properties.transform_keys(&:to_sym),
      split_pane: worksheet.split_pane&.to_h&.slice(:x_split, :y_split, :top_left_cell, :active_pane),
      selection: worksheet.selection&.to_h&.slice(:active_cell, :sqref, :pane),
      header_footer: worksheet.header_footer.to_h,
      page_margins: worksheet.page_margins&.to_h,
      print_options: worksheet.print_options.to_h,
      sheet_protection: worksheet.sheet_protection&.to_h&.except(:hash_value, :salt_value, :algorithm_name, :spin_count, :password),
      protected_ranges: worksheet.protected_ranges.map { |pr| [pr.name, pr.sqref, !pr.hash_value.nil?] },
      ignored_errors: worksheet.ignored_errors.map { |ie| [ie.sqref, ie.number_stored_as_text?, ie.eval_error?] },
      cell_watches: worksheet.cell_watches,
      phonetic: worksheet.phonetic_properties&.to_h,
      code_name: worksheet.code_name,
      fit_to_page: worksheet.fit_to_page?,
      full_calc_on_load: worksheet.full_calc_on_load?
    }
  end

  test "stateful property: nested outlines and remaining sheet metadata parity, Worksheet#with, and roundtrip" do
    init_state = SheetMetaState.new(
      nested_outlines: nil,
      col_groups: nil,
      outline_props: nil,
      split_pane: nil,
      selection: nil,
      header_footer: nil,
      page_margins: nil,
      print_options: nil,
      sheet_protection: nil,
      protected_ranges: nil,
      ignored_errors: nil,
      cell_watches: nil,
      phonetic_props: nil,
      code_name: nil,
      fit_to_page: nil,
      sheet_calc_pr: nil,
      dirty: false
    )

    Pbt.assert(num_runs: 12) do
      Pbt.stateful(max_steps: 9) do |m|
        m.initial_state Pbt.constant(init_state)
        m.sut { |_init| SheetMetaSut.new }

        m.command :apply_outlines, weight: 3 do |c|
          c.arguments(
            row_hidden: Pbt.boolean,
            row_collapsed: Pbt.boolean,
            col_hidden: Pbt.boolean,
            col_collapsed: Pbt.boolean,
            summary_below: Pbt.boolean,
            summary_right: Pbt.boolean,
            show_symbols: Pbt.boolean,
            apply_styles: Pbt.boolean
          )
          c.next_state do |s, args|
            no = { hidden: args[:row_hidden], collapsed: args[:row_collapsed] }
            cg = { hidden: args[:col_hidden], collapsed: args[:col_collapsed] }
            op = {
              summary_below: args[:summary_below],
              summary_right: args[:summary_right],
              show_outline_symbols: args[:show_symbols],
              apply_styles: args[:apply_styles]
            }
            s.with(nested_outlines: no, col_groups: cg, outline_props: op, dirty: true)
          end
          c.run do |sut, args|
            sut.apply_outlines(
              nested_outlines: { hidden: args[:row_hidden], collapsed: args[:row_collapsed] },
              col_groups: { hidden: args[:col_hidden], collapsed: args[:col_collapsed] },
              outline_props: {
                summary_below: args[:summary_below],
                summary_right: args[:summary_right],
                show_outline_symbols: args[:show_symbols],
                apply_styles: args[:apply_styles]
              }
            )
          end
        end

        m.command :apply_split_and_selection, weight: 2 do |c|
          c.arguments(
            x_split: Pbt.constant_from(1500, 2400),
            y_split: Pbt.constant_from(1200, 1800),
            pane: Pbt.constant_from("bottomRight", "bottomLeft")
          )
          c.next_state do |s, args|
            sp = { x_split: args[:x_split], y_split: args[:y_split], top_left_cell: "C3", active_pane: args[:pane] }
            sel = { active_cell: "C3", sqref: "C3:D4", pane: args[:pane] }
            s.with(split_pane: sp, selection: sel, dirty: true)
          end
          c.run do |sut, args|
            sut.apply_split_and_selection(
              split_pane: { x_split: args[:x_split], y_split: args[:y_split], top_left_cell: "C3", active_pane: args[:pane] },
              selection: { active_cell: "C3", sqref: "C3:D4", pane: args[:pane] }
            )
          end
        end

        m.command :apply_print_and_hf, weight: 2 do |c|
          c.arguments(
            diff_first: Pbt.boolean,
            grid_lines: Pbt.boolean,
            headings: Pbt.boolean,
            h_center: Pbt.boolean,
            v_center: Pbt.boolean,
            fit_to_page: Pbt.boolean
          )
          c.next_state do |s, args|
            hf = { odd_header: "&CReport &A", odd_footer: "&RPage &P", different_first: args[:diff_first] }
            hf[:first_header] = "&LFirst" if args[:diff_first]
            pm = { left: 0.5, right: 0.5, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 }
            po = {
              grid_lines: args[:grid_lines],
              headings: args[:headings],
              horizontal_centered: args[:h_center],
              vertical_centered: args[:v_center]
            }
            s.with(header_footer: hf, page_margins: pm, print_options: po, fit_to_page: args[:fit_to_page], dirty: true)
          end
          c.run do |sut, args|
            hf = { odd_header: "&CReport &A", odd_footer: "&RPage &P", different_first: args[:diff_first] }
            hf[:first_header] = "&LFirst" if args[:diff_first]
            sut.apply_print_and_hf(
              header_footer: hf,
              page_margins: { left: 0.5, right: 0.5, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
              print_options: {
                grid_lines: args[:grid_lines],
                headings: args[:headings],
                horizontal_centered: args[:h_center],
                vertical_centered: args[:v_center]
              },
              fit_to_page: args[:fit_to_page]
            )
          end
        end

        m.command :apply_protection, weight: 2 do |c|
          c.arguments(
            allow_sort: Pbt.boolean,
            allow_filter: Pbt.boolean,
            sp_pw: Pbt.constant_from(nil, "pw12"),
            rng_pw: Pbt.constant_from(nil, "rng1")
          )
          c.next_state do |s, args|
            sp = { sheet: true, sort: args[:allow_sort], auto_filter: args[:allow_filter], password: args[:sp_pw] }.compact
            prs = [{ name: "Editable", sqref: "B2:C4", password: args[:rng_pw] }.compact]
            s.with(sheet_protection: sp, protected_ranges: prs, dirty: true)
          end
          c.run do |sut, args|
            sut.apply_protection(
              sheet_protection: { sheet: true, sort: args[:allow_sort], auto_filter: args[:allow_filter], password: args[:sp_pw] }.compact,
              protected_ranges: [{ name: "Editable", sqref: "B2:C4", password: args[:rng_pw] }.compact]
            )
          end
        end

        m.command :apply_misc_meta, weight: 2 do |c|
          c.arguments(eval_err: Pbt.boolean, full_calc: Pbt.boolean)
          c.next_state do |s, args|
            ie = [{ sqref: "A1:B3", number_stored_as_text: true, eval_error: args[:eval_err] }]
            cw = ["A1"]
            pp = { font_id: 0, type: "halfwidthKatakana", alignment: "left" }
            cn = "SheetCode1"
            scp = { full_calc_on_load: args[:full_calc] }
            s.with(ignored_errors: ie, cell_watches: cw, phonetic_props: pp, code_name: cn, sheet_calc_pr: scp, dirty: true)
          end
          c.run do |sut, args|
            sut.apply_misc_meta(
              ignored_errors: [{ sqref: "A1:B3", number_stored_as_text: true, eval_error: args[:eval_err] }],
              cell_watches: ["A1"],
              phonetic_props: { font_id: 0, type: "halfwidthKatakana", alignment: "left" },
              code_name: "SheetCode1",
              sheet_calc_pr: { full_calc_on_load: args[:full_calc] }
            )
          end
        end

        m.invariant(:sheet_meta_parity_and_roundtrip) do |state, sut|
          next unless state.dirty

          mem_bytes, stream_bytes = sut.snapshot_bytes
          assert_empty(ExcelRules.check(mem_bytes).map(&:to_s))
          assert_empty(ExcelRules.check(stream_bytes).map(&:to_s))
          assert_empty(XsdValidator.check(mem_bytes))
          assert_empty(XsdValidator.check(stream_bytes))

          m_wb = Xlsxrb.load(StringIO.new(mem_bytes))
          s_wb = Xlsxrb.load(StringIO.new(stream_bytes))
          m_snap = extract_sheet_meta_snapshot(m_wb.sheet(0))
          s_snap = extract_sheet_meta_snapshot(s_wb.sheet(0))
          assert_equal(m_snap, s_snap, "build vs stream sheet metadata mismatch")

          stream_snap = nil
          Xlsxrb.read(StringIO.new(mem_bytes)) { |ss| stream_snap = extract_sheet_meta_snapshot(ss) }
          assert_equal(m_snap.except(:row_outlines), stream_snap.except(:row_outlines), "StreamSheet metadata mismatch")

          resaved = m_wb.to_data
          assert_empty(ExcelRules.check(resaved).map(&:to_s))
          assert_empty(XsdValidator.check(resaved))
          r_snap = extract_sheet_meta_snapshot(Xlsxrb.load(StringIO.new(resaved)).sheet(0))
          assert_equal(m_snap, r_snap, "sheet metadata roundtrip mismatch after to_data")

          mut_wb = m_wb.with(
            sheets: [m_wb.sheet(0).with(cell_watches: ["B2"], print_options: { grid_lines: true, headings: true })]
          )
          mut_bytes = mut_wb.to_data
          assert_empty(ExcelRules.check(mut_bytes).map(&:to_s))
          assert_empty(XsdValidator.check(mut_bytes))
          mut_ws = Xlsxrb.load(StringIO.new(mut_bytes)).sheet(0)
          assert_equal(["B2"], mut_ws.cell_watches)
          assert_true(mut_ws.print_options.grid_lines?)
          assert_true(mut_ws.print_options.headings?)
        end
      end
    end
  end

  # --- Item 8: Elements::Styles mutators + dual styles.xml writers + dual parsers Stateful PBT ---

  StylesState = Data.define(:fonts, :fills, :borders, :num_fmts, :cell_xfs, :dxfs, :dirty)

  # Drives Elements::Styles mutators and Ooxml::Writer stylesheet registration in parallel.
  class StylesSut
    attr_reader :styles

    def initialize
      @styles = Xlsxrb::Elements::Styles.new(
        fonts: [{ name: "Calibri", size: 11.0 }],
        fills: [{ pattern_type: "none" }, { pattern_type: "gray125" }],
        borders: [{}],
        num_fmts: {},
        cell_xfs: [{ num_fmt_id: 0, font_id: 0, fill_id: 0, border_id: 0 }],
        dxfs: []
      )
      @low_writer = Xlsxrb::Ooxml::Writer.new
    end

    def add_font(entry)
      id = @styles.add_font(entry)
      low_entry = entry.dup
      low_entry[:sz] = low_entry.delete(:size) if low_entry.key?(:size)
      @low_writer.add_font(**low_entry)
      id
    end

    def add_fill(entry)
      id = @styles.add_fill(entry)
      @low_writer.add_fill(**entry)
      id
    end

    def add_border(entry)
      id = @styles.add_border(entry)
      @low_writer.add_border(**entry)
      id
    end

    def add_number_format(id, code)
      @styles.add_number_format(id, code)
    end

    def add_style(entry)
      @styles.add_style(entry)
    end

    def add_dxf(entry)
      id = @styles.add_dxf(entry)
      @low_writer.add_dxf(**entry)
      id
    end

    def build_workbook_styles_xml
      row = Xlsxrb::Elements::Row.new(
        index: 0,
        cells: [Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: 42, style_index: @styles.cell_xfs.size - 1)]
      )
      ws = Xlsxrb::Elements::Worksheet.new(name: "S", rows: [row])
      wb = Xlsxrb::Elements::Workbook.new(sheets: [ws], styles: @styles)
      bytes = Xlsxrb.write(wb)
      [bytes, Xlsxrb::Ooxml::ZipReader.new(StringIO.new(bytes)).read_entry("xl/styles.xml")]
    end

    def low_level_styles_xml
      @low_writer.send(:generate_styles_xml)
    end
  end

  def normalize_parsed_color(color, parent = nil, prefix = nil)
    if color.is_a?(String)
      base = { rgb: color }
      if parent
        t_key = prefix ? :"#{prefix}_theme" : :theme
        i_key = prefix ? :"#{prefix}_indexed" : :indexed
        n_key = prefix ? :"#{prefix}_tint" : :tint
        a_key = prefix ? :"#{prefix}_auto" : :auto
        base[:auto] = true if parent[a_key]
        base[:theme] = parent[t_key] if parent[t_key]
        base[:indexed] = parent[i_key] if parent[i_key]
        base[:tint] = parent[n_key]&.round(4) if parent[n_key]
      end
      return base
    end

    if color.nil? && parent
      t_key = prefix ? :"#{prefix}_theme" : :theme
      i_key = prefix ? :"#{prefix}_indexed" : :indexed
      n_key = prefix ? :"#{prefix}_tint" : :tint
      a_key = prefix ? :"#{prefix}_auto" : :auto
      out = {
        auto: parent[a_key] ? true : nil,
        theme: parent[t_key],
        indexed: parent[i_key],
        tint: parent[n_key]&.round(4)
      }.compact
      return out.empty? ? nil : out
    end

    return nil if color.nil? || color.empty?

    {
      auto: color[:auto] || nil,
      rgb: color[:rgb],
      theme: color[:theme],
      indexed: color[:indexed],
      tint: color[:tint]&.round(4)
    }.compact
  end

  def normalize_parsed_styles_xml(xml)
    parsed = Xlsxrb::Ooxml::StylesParser.parse(xml)
    listener = Xlsxrb::Ooxml::Reader::StylesListener.new
    Xlsxrb::Ooxml::XmlParser.parse(xml, listener)

    norm_u = ->(u) { u == true ? "single" : u }
    norm_fonts = lambda do |list|
      list.map do |f|
        h = f.is_a?(Xlsxrb::Elements::MetadataRecord) ? f.to_h : f
        {
          name: h[:name],
          size: (h[:size] || h[:sz])&.to_f,
          bold: h[:bold] ? true : false,
          italic: h[:italic] ? true : false,
          strike: h[:strike] ? true : false,
          underline: norm_u.call(h[:underline]),
          vert_align: h[:vert_align],
          color: normalize_parsed_color(h[:color], h)
        }
      end
    end

    norm_fills = lambda do |list|
      list.map do |fi|
        h = fi.is_a?(Xlsxrb::Elements::MetadataRecord) ? fi.to_h : fi
        if h[:gradient]
          g = h[:gradient]
          {
            gradient: {
              type: g[:type],
              degree: g[:degree]&.to_f,
              stops: (g[:stops] || []).map { |st| { position: st[:position]&.to_f, color: normalize_parsed_color(st[:color], st) } }
            }
          }
        else
          {
            pattern_type: h[:pattern_type] || h[:pattern],
            fg_color: normalize_parsed_color(h[:fg_color], h, "fg_color"),
            bg_color: normalize_parsed_color(h[:bg_color], h, "bg_color")
          }
        end
      end
    end

    norm_borders = lambda do |list|
      list.map do |b|
        h = b.is_a?(Xlsxrb::Elements::MetadataRecord) ? b.to_h : b
        sides = %i[left right top bottom diagonal].each_with_object({}) do |side, acc|
          s = h[side]
          acc[side] = { style: s[:style], color: normalize_parsed_color(s[:color], s) }.compact if s && !s.empty?
        end
        sides.merge(diagonal_up: h[:diagonal_up] ? true : false, diagonal_down: h[:diagonal_down] ? true : false)
      end
    end

    norm_xfs = lambda do |list|
      list.map do |x|
        h = x.is_a?(Xlsxrb::Elements::MetadataRecord) ? x.to_h : x
        al = h[:alignment].is_a?(Xlsxrb::Elements::MetadataRecord) ? h[:alignment].to_h : h[:alignment]
        pr = h[:protection].is_a?(Xlsxrb::Elements::MetadataRecord) ? h[:protection].to_h : h[:protection]
        {
          num_fmt_id: h[:num_fmt_id] || 0,
          font_id: h[:font_id] || 0,
          fill_id: h[:fill_id] || 0,
          border_id: h[:border_id] || 0,
          alignment: al,
          protection: pr
        }
      end
    end

    norm_dxfs = lambda do |list|
      list.map do |d|
        h = d.is_a?(Xlsxrb::Elements::MetadataRecord) ? d.to_h : d
        df = h[:font].is_a?(Xlsxrb::Elements::MetadataRecord) ? h[:font].to_h : h[:font]
        dfi = h[:fill].is_a?(Xlsxrb::Elements::MetadataRecord) ? h[:fill].to_h : h[:fill]
        db = h[:border].is_a?(Xlsxrb::Elements::MetadataRecord) ? h[:border].to_h : h[:border]
        al = h[:alignment].is_a?(Xlsxrb::Elements::MetadataRecord) ? h[:alignment].to_h : h[:alignment]
        pr = h[:protection].is_a?(Xlsxrb::Elements::MetadataRecord) ? h[:protection].to_h : h[:protection]
        {
          font: df ? norm_fonts.call([df]).first : nil,
          fill: dfi ? norm_fills.call([dfi]).first : nil,
          border: db ? norm_borders.call([db]).first : nil,
          alignment: al,
          protection: pr
        }
      end
    end

    sp_view = {
      num_fmts: parsed.number_formats,
      fonts: norm_fonts.call(parsed.fonts),
      fills: norm_fills.call(parsed.fills),
      borders: norm_borders.call(parsed.borders),
      cell_xfs: norm_xfs.call(parsed.cell_xfs),
      dxfs: norm_dxfs.call(parsed.dxfs)
    }
    sl_view = {
      num_fmts: listener.num_fmts,
      fonts: norm_fonts.call(listener.fonts),
      fills: norm_fills.call(listener.fills),
      borders: norm_borders.call(listener.borders),
      cell_xfs: norm_xfs.call(listener.cell_xfs),
      dxfs: norm_dxfs.call(listener.dxfs)
    }
    [sp_view, sl_view, parsed]
  end

  test "stateful property: Elements::Styles mutators, dual styles writers, and dual styles parsers equivalence" do
    init_state = StylesState.new(
      fonts: [{ name: "Calibri", size: 11.0 }],
      fills: [{ pattern_type: "none" }, { pattern_type: "gray125" }],
      borders: [{}],
      num_fmts: {},
      cell_xfs: [{ num_fmt_id: 0, font_id: 0, fill_id: 0, border_id: 0 }],
      dxfs: [],
      dirty: false
    )

    Pbt.assert(num_runs: 15) do
      Pbt.stateful(max_steps: 10) do |m|
        m.initial_state Pbt.constant(init_state)
        m.sut { |_init| StylesSut.new }

        m.command :add_font, weight: 3 do |c|
          c.arguments(
            bold: Pbt.boolean,
            italic: Pbt.boolean,
            strike: Pbt.boolean,
            underline: Pbt.constant_from(nil, "single", "double"),
            vert_align: Pbt.constant_from(nil, "superscript", "subscript"),
            size: Pbt.constant_from(10.0, 14.0),
            name: Pbt.constant_from("Arial", "Calibri"),
            use_theme: Pbt.boolean
          )
          c.next_state do |s, args|
            color = args[:use_theme] ? { theme: 4, tint: -0.25 } : { rgb: "FFFF0000" }
            entry = {
              name: args[:name],
              size: args[:size],
              bold: args[:bold],
              italic: args[:italic],
              strike: args[:strike],
              underline: args[:underline],
              vert_align: args[:vert_align],
              color: color
            }.compact
            s.with(fonts: [*s.fonts, entry], dirty: true)
          end
          c.run do |sut, args|
            color = args[:use_theme] ? { theme: 4, tint: -0.25 } : { rgb: "FFFF0000" }
            entry = {
              name: args[:name],
              size: args[:size],
              bold: args[:bold],
              italic: args[:italic],
              strike: args[:strike],
              underline: args[:underline],
              vert_align: args[:vert_align],
              color: color
            }.compact
            sut.add_font(entry)
          end
        end

        m.command :add_fill, weight: 3 do |c|
          c.arguments(
            mode: Pbt.constant_from(:solid, :gradient, :gradient_nil_stops),
            rgb: Pbt.constant_from("FFFFFF00", "FFCCDDFF")
          )
          c.next_state do |s, args|
            entry = case args[:mode]
                    when :solid
                      { pattern_type: "solid", fg_color: { rgb: args[:rgb] } }
                    when :gradient
                      { gradient: { type: "linear", degree: 90.0, stops: [{ position: 0.0, color: { rgb: "FFFF0000" } }, { position: 1.0, color: { rgb: "FF0000FF" } }] } }
                    when :gradient_nil_stops
                      { gradient: { type: "linear", degree: 45.0 } }
                    end
            s.with(fills: [*s.fills, entry], dirty: true)
          end
          c.run do |sut, args|
            entry = case args[:mode]
                    when :solid
                      { pattern_type: "solid", fg_color: { rgb: args[:rgb] } }
                    when :gradient
                      { gradient: { type: "linear", degree: 90.0, stops: [{ position: 0.0, color: { rgb: "FFFF0000" } }, { position: 1.0, color: { rgb: "FF0000FF" } }] } }
                    when :gradient_nil_stops
                      { gradient: { type: "linear", degree: 45.0 } }
                    end
            sut.add_fill(entry)
          end
        end

        m.command :add_border, weight: 2 do |c|
          c.arguments(
            style: Pbt.constant_from("thin", "medium"),
            diag_up: Pbt.boolean,
            diag_down: Pbt.boolean
          )
          c.next_state do |s, args|
            side = { style: args[:style], color: { rgb: "FF000000" } }
            entry = { left: side, right: side, diagonal_up: args[:diag_up], diagonal_down: args[:diag_down] }
            entry[:diagonal] = side if args[:diag_up] || args[:diag_down]
            s.with(borders: [*s.borders, entry], dirty: true)
          end
          c.run do |sut, args|
            side = { style: args[:style], color: { rgb: "FF000000" } }
            entry = { left: side, right: side, diagonal_up: args[:diag_up], diagonal_down: args[:diag_down] }
            entry[:diagonal] = side if args[:diag_up] || args[:diag_down]
            sut.add_border(entry)
          end
        end

        m.command :add_number_format, weight: 2 do |c|
          c.arguments(code: Pbt.constant_from("#,##0.00", "yyyy-mm-dd", '"$"#,##0'))
          c.next_state do |s, args|
            next_id = 164 + s.num_fmts.size
            s.with(num_fmts: s.num_fmts.merge(next_id => args[:code]), dirty: true)
          end
          c.run { |sut, args| sut.add_number_format(164 + sut.styles.number_formats.size, args[:code]) }
        end

        m.command :add_style, weight: 3 do |c|
          c.arguments(
            horiz: Pbt.constant_from("left", "center", "right"),
            wrap: Pbt.boolean,
            locked: Pbt.boolean,
            hidden: Pbt.boolean,
            with_fmt: Pbt.boolean
          )
          c.next_state do |s, args|
            font = { name: "Arial", size: 12.0, bold: true }
            fill = { pattern_type: "solid", fg_color: { rgb: "FFEEFFEE" } }
            border = { top: { style: "thin", color: { rgb: "FF112233" } } }
            fmt_id = args[:with_fmt] ? 164 + s.num_fmts.size : 0
            num_fmts = args[:with_fmt] ? s.num_fmts.merge(fmt_id => "#,##0.000") : s.num_fmts
            xf = {
              font_id: s.fonts.size,
              fill_id: s.fills.size,
              border_id: s.borders.size,
              num_fmt_id: fmt_id,
              alignment: { horizontal: args[:horiz], vertical: "center", wrap_text: args[:wrap] },
              protection: { locked: args[:locked], hidden: args[:hidden] }
            }
            s.with(
              fonts: [*s.fonts, font],
              fills: [*s.fills, fill],
              borders: [*s.borders, border],
              num_fmts: num_fmts,
              cell_xfs: [*s.cell_xfs, xf],
              dirty: true
            )
          end
          c.run do |sut, args|
            spec = {
              font: { name: "Arial", size: 12.0, bold: true },
              fill: { pattern_type: "solid", fg_color: { rgb: "FFEEFFEE" } },
              border: { top: { style: "thin", color: { rgb: "FF112233" } } },
              alignment: { horizontal: args[:horiz], vertical: "center", wrap_text: args[:wrap] },
              protection: { locked: args[:locked], hidden: args[:hidden] }
            }
            if args[:with_fmt]
              spec[:number_format] =
                sut.styles.number_formats.value?("#,##0.000") ? "#,##0.00#{sut.styles.number_formats.size}" : "#,##0.000"
            end
            sut.add_style(spec)
          end
        end

        m.command :add_dxf, weight: 3 do |c|
          c.arguments(bold: Pbt.boolean, italic: Pbt.boolean, rgb: Pbt.constant_from("FFFFAAAA", "FFAAFFAA"))
          c.next_state do |s, args|
            dxf = {
              font: { bold: args[:bold], italic: args[:italic], color: { rgb: "FF000000" } },
              fill: { pattern_type: "solid", fg_color: { rgb: args[:rgb] } },
              border: { left: { style: "thin", color: { rgb: "FF000000" } } },
              alignment: { horizontal: "center", wrap_text: true },
              protection: { locked: false, hidden: true }
            }
            s.with(dxfs: [*s.dxfs, dxf], dirty: true)
          end
          c.run do |sut, args|
            sut.add_dxf(
              font: { bold: args[:bold], italic: args[:italic], color: { rgb: "FF000000" } },
              fill: { pattern_type: "solid", fg_color: { rgb: args[:rgb] } },
              border: { left: { style: "thin", color: { rgb: "FF000000" } } },
              alignment: { horizontal: "center", wrap_text: true },
              protection: { locked: false, hidden: true }
            )
          end
        end

        m.invariant(:styles_dual_writer_and_parser_parity) do |state, sut|
          next unless state.dirty

          assert_equal(state.fonts.size, sut.styles.fonts.size, "live fonts size mismatch")
          assert_equal(state.fills.size, sut.styles.fills.size, "live fills size mismatch")
          assert_equal(state.borders.size, sut.styles.borders.size, "live borders size mismatch")
          assert_equal(state.cell_xfs.size, sut.styles.cell_xfs.size, "live cell_xfs size mismatch")
          assert_equal(state.dxfs.size, sut.styles.dxfs.size, "live dxfs size mismatch")

          bytes, wb_styles_xml = sut.build_workbook_styles_xml
          assert_empty(ExcelRules.check(bytes).map(&:to_s))
          assert_empty(XsdValidator.check(bytes))

          wb_sp, wb_sl, parsed_wb = normalize_parsed_styles_xml(wb_styles_xml)
          assert_equal(wb_sp, wb_sl, "StylesParser vs StylesListener mismatch on WorkbookWriter styles.xml")

          low_styles_xml = sut.low_level_styles_xml
          low_sp, low_sl, _parsed_low = normalize_parsed_styles_xml(low_styles_xml)
          assert_equal(low_sp, low_sl, "StylesParser vs StylesListener mismatch on Ooxml::Writer styles.xml")

          assert_equal(state.fonts.size, parsed_wb.fonts.size)
          assert_equal(state.fills.size, parsed_wb.fills.size)
          assert_equal(state.borders.size, parsed_wb.borders.size)
          assert_equal(state.cell_xfs.size, parsed_wb.cell_xfs.size)
          assert_equal(state.dxfs.size, parsed_wb.dxfs.size)
        end
      end
    end
  end
end
