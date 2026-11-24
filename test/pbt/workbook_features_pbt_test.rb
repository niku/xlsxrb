# frozen_string_literal: true

require "test_helper"
require "pbt"
require "stringio"
require "rexml/document"
require "time"
require_relative "../support/excel_rules"

class WorkbookFeaturesPbtTest < Test::Unit::TestCase
  TINY_PNG = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
             .unpack1("m0").b.freeze

  def setup
    omit "Skip property-based tests under RBS runtime testing to prevent OOM / slowdown" if defined?(RBS::Test)
  end

  # --- Item 1: Multi-sheet threaded_comments, pivot_tables, sparkline_groups, background_image ---

  AuxiliaryFeaturesState = Data.define(
    :sheets,
    :active_sheet_idx,
    :shared_strings,
    :next_comment_seq,
    :next_pivot_seq,
    :dirty
  )

  # Drives WorkbookBuilder and StreamWriter with multi-sheet auxiliary parts.
  class AuxiliaryFeaturesSut
    def initialize
      @sheets = [
        {
          name: "Data",
          rows: [%w[Region Quarter Sales], ["East", "Q1", 100], ["West", "Q2", 250]],
          threaded_comments: [],
          pivot_tables: [],
          sparkline_groups: [],
          background_image: nil
        }
      ]
      @active_sheet_idx = 0
      @shared_strings = true
      @next_comment_seq = 1
      @next_pivot_seq = 1
    end

    def add_sheet
      idx = @sheets.size + 1
      @sheets << {
        name: "Sheet#{idx}",
        rows: [%w[Label Value], ["Row1", 42]],
        threaded_comments: [],
        pivot_tables: [],
        sparkline_groups: [],
        background_image: nil
      }
      @active_sheet_idx = @sheets.size - 1
      nil
    end

    def add_threaded_comment(cell, text, author:, done:)
      seq = @next_comment_seq
      @next_comment_seq += 1
      id = format("{00000000-0000-0000-0000-%012d}", seq)
      @sheets[@active_sheet_idx][:threaded_comments] << {
        cell: cell,
        text: text,
        author: author,
        id: id,
        parent_id: nil,
        done: done
      }
      nil
    end

    def reply_threaded_comment(text, author:, done:)
      root = @sheets[@active_sheet_idx][:threaded_comments].first
      seq = @next_comment_seq
      @next_comment_seq += 1
      id = format("{00000000-0000-0000-0000-%012d}", seq)
      @sheets[@active_sheet_idx][:threaded_comments] << {
        cell: root[:cell],
        text: text,
        author: author,
        id: id,
        parent_id: root[:id],
        done: done
      }
      nil
    end

    def add_pivot_table(with_col_field:, subtotal:)
      seq = @next_pivot_seq
      @next_pivot_seq += 1
      existing_count = @sheets[@active_sheet_idx][:pivot_tables].size
      @sheets[@active_sheet_idx][:pivot_tables] << {
        source_ref: "Data!A1:C3",
        field_names: %w[Region Quarter Sales],
        row_fields: ["Region"],
        column_fields: (with_col_field ? ["Quarter"] : []),
        data_fields: [{ field: "Sales", subtotal: subtotal }],
        ref: "E#{1 + (existing_count * 10)}",
        name: "Pivot#{seq}",
        items: { 0 => %w[East West], 1 => %w[Q1 Q2] }
      }
      nil
    end

    def add_sparkline_group(type:, markers:, high:, low:)
      existing_count = @sheets[@active_sheet_idx][:sparkline_groups].size
      @sheets[@active_sheet_idx][:sparkline_groups] << {
        sparklines: [{ data_ref: "Data!C2:C3", location_ref: "H#{1 + existing_count}" }],
        type: type,
        markers: markers,
        high: high,
        low: low,
        color_series: "FF336699"
      }
      nil
    end

    def set_background_image(mode:)
      val = case mode
            when :hash then { data: TINY_PNG, ext: "png" }
            when :raw then TINY_PNG
            end
      @sheets[@active_sheet_idx][:background_image] = val
      nil
    end

    def toggle_shared_strings
      @shared_strings = !@shared_strings
      nil
    end

    def snapshot_bytes
      wb = Xlsxrb.build do |b|
        @sheets.each do |spec|
          b.sheet(spec[:name]) { |s| populate_sheet(s, spec) }
        end
      end
      mem_bytes = Xlsxrb.write(wb, shared_strings: @shared_strings)

      stream_bytes = Xlsxrb.write(shared_strings: @shared_strings) do |sw|
        @sheets.each do |spec|
          sw.sheet(spec[:name]) { |s| populate_sheet(s, spec) }
        end
      end
      [mem_bytes, stream_bytes]
    end

    private

    def populate_sheet(sheet, spec)
      spec[:rows].each { |r| sheet.row(r) }
      spec[:threaded_comments].each do |tc|
        sheet.threaded_comment(tc[:cell], tc[:text], author: tc[:author], id: tc[:id], parent_id: tc[:parent_id], done: tc[:done])
      end
      spec[:pivot_tables].each do |pt|
        sheet.pivot_table(
          pt[:source_ref],
          field_names: pt[:field_names],
          row_fields: pt[:row_fields],
          column_fields: pt[:column_fields],
          data_fields: pt[:data_fields],
          ref: pt[:ref],
          name: pt[:name],
          items: pt[:items]
        )
      end
      spec[:sparkline_groups].each do |sg|
        sheet.sparkline_group(**sg)
      end
      sheet.background_image = spec[:background_image] if spec[:background_image]
    end
  end

  def assert_threaded_comment_package_integrity!(bytes, label)
    zip = Xlsxrb::Ooxml::ZipReader.new(StringIO.new(bytes))
    entries = zip.read_all
    ct_doc = REXML::Document.new(entries.fetch("[Content_Types].xml"))
    overrides = []
    REXML::XPath.each(ct_doc, "//Override") { |el| overrides << el.attributes["PartName"] }

    tc_parts = entries.keys.grep(%r{\Axl/threadedComments/threadedComment\d+\.xml\z}).sort
    return if tc_parts.empty?

    assert(entries.key?("xl/persons/person.xml"), "[#{label}] Missing xl/persons/person.xml when threaded comments exist")
    assert_include(overrides, "/xl/persons/person.xml", "[#{label}] Missing Content_Types override for /xl/persons/person.xml")

    person_doc = REXML::Document.new(entries["xl/persons/person.xml"])
    person_ids = []
    REXML::XPath.each(person_doc, "//person") { |el| person_ids << el.attributes["id"] }

    tc_parts.each do |tc_part|
      assert_include(overrides, "/#{tc_part}", "[#{label}] Missing Content_Types override for /#{tc_part}")
      tc_doc = REXML::Document.new(entries[tc_part])
      REXML::XPath.each(tc_doc, "//threadedComment") do |el|
        pid = el.attributes["personId"]
        assert_include(person_ids, pid, "[#{label}] Dangling personId #{pid.inspect} in #{tc_part} (person.xml has #{person_ids.inspect})")
      end
    end
  end

  def assert_auxiliary_workbooks_match!(state, m_wb, s_wb, bytes_list)
    m_stream_sheets = []
    Xlsxrb.read(StringIO.new(bytes_list[0])) do |sh|
      m_stream_sheets << {
        tcs: sh.threaded_comments.map { |tc| [tc.text, tc.author] },
        pts: sh.pivot_tables.map(&:name),
        sgs: sh.sparkline_groups.map { |sg| [sg.type || "line", sg.sparklines.size] },
        bg: sh.background_image&.data&.b
      }
    end
    s_stream_sheets = []
    Xlsxrb.read(StringIO.new(bytes_list[1])) do |sh|
      s_stream_sheets << {
        tcs: sh.threaded_comments.map { |tc| [tc.text, tc.author] },
        pts: sh.pivot_tables.map(&:name),
        sgs: sh.sparkline_groups.map { |sg| [sg.type || "line", sg.sparklines.size] },
        bg: sh.background_image&.data&.b
      }
    end

    state.sheets.each_with_index do |spec, idx|
      m_ws = m_wb.sheet(idx)
      s_ws = s_wb.sheet(idx)

      exp_tcs = spec[:threaded_comments].map { |tc| [tc[:cell], tc[:text], tc[:author], tc[:parent_id], tc[:done] ? true : false] }
      m_tcs = m_ws.threaded_comments.map { |tc| [tc.ref, tc.text, tc.author, tc.parent_id, tc.resolved?] }
      s_tcs = s_ws.threaded_comments.map { |tc| [tc.ref, tc.text, tc.author, tc.parent_id, tc.resolved?] }
      assert_equal(exp_tcs, m_tcs)
      assert_equal(exp_tcs, s_tcs)
      assert_equal(exp_tcs.map { |tc| [tc[1], tc[2]] }, m_stream_sheets[idx][:tcs])
      assert_equal(exp_tcs.map { |tc| [tc[1], tc[2]] }, s_stream_sheets[idx][:tcs])

      exp_pts = spec[:pivot_tables].map { |pt| [pt[:name], pt[:ref], [0], (pt[:column_fields].empty? ? [] : [1]), pt[:data_fields].size] }
      m_pts = m_ws.pivot_tables.map { |pt| [pt.name, pt.ref, pt.row_fields, pt.column_fields, pt.data_fields.size] }
      s_pts = s_ws.pivot_tables.map { |pt| [pt.name, pt.ref, pt.row_fields, pt.column_fields, pt.data_fields.size] }
      assert_equal(exp_pts, m_pts)
      assert_equal(exp_pts, s_pts)
      assert_equal(exp_pts.map(&:first), m_stream_sheets[idx][:pts])
      assert_equal(exp_pts.map(&:first), s_stream_sheets[idx][:pts])

      exp_sgs = spec[:sparkline_groups].map do |sg|
        [sg[:type].to_s, sg[:markers] ? true : false, sg[:high] ? true : false, sg[:low] ? true : false,
         sg[:sparklines].map { |sp| [sp[:location_ref], sp[:data_ref]] }]
      end
      m_sgs = m_ws.sparkline_groups.map do |sg|
        [sg.type || "line", sg.markers?, sg.high?, sg.low?, sg.sparklines.map { |sp| [sp.sqref, sp.formula] }]
      end
      s_sgs = s_ws.sparkline_groups.map do |sg|
        [sg.type || "line", sg.markers?, sg.high?, sg.low?, sg.sparklines.map { |sp| [sp.sqref, sp.formula] }]
      end
      assert_equal(exp_sgs, m_sgs)
      assert_equal(exp_sgs, s_sgs)
      assert_equal(exp_sgs.map { |sg| [sg[0], sg[4].size] }, m_stream_sheets[idx][:sgs])
      assert_equal(exp_sgs.map { |sg| [sg[0], sg[4].size] }, s_stream_sheets[idx][:sgs])

      exp_bg = spec[:background_image] ? TINY_PNG : nil
      if exp_bg.nil?
        assert_nil(m_ws.background_image)
        assert_nil(s_ws.background_image)
        assert_nil(m_stream_sheets[idx][:bg])
        assert_nil(s_stream_sheets[idx][:bg])
      else
        assert_equal(exp_bg, m_ws.background_image&.data&.b)
        assert_equal(exp_bg, s_ws.background_image&.data&.b)
        assert_equal(exp_bg, m_stream_sheets[idx][:bg])
        assert_equal(exp_bg, s_stream_sheets[idx][:bg])
      end
    end
  end

  test "stateful property: multi-sheet threaded comments, pivot tables, sparklines, and background image parity and roundtrip" do
    init_state = AuxiliaryFeaturesState.new(
      sheets: [
        {
          name: "Data",
          rows: [%w[Region Quarter Sales], ["East", "Q1", 100], ["West", "Q2", 250]],
          threaded_comments: [],
          pivot_tables: [],
          sparkline_groups: [],
          background_image: nil
        }
      ],
      active_sheet_idx: 0,
      shared_strings: true,
      next_comment_seq: 1,
      next_pivot_seq: 1,
      dirty: false
    )

    Pbt.assert(num_runs: 15) do
      Pbt.stateful(max_steps: 9) do |m|
        m.initial_state Pbt.constant(init_state)
        m.sut { |_init| AuxiliaryFeaturesSut.new }

        m.command :add_sheet, weight: 2 do |c|
          c.guard { |s| s.sheets.size < 3 }
          c.next_state do |s|
            idx = s.sheets.size + 1
            new_sheet = {
              name: "Sheet#{idx}",
              rows: [%w[Label Value], ["Row1", 42]],
              threaded_comments: [],
              pivot_tables: [],
              sparkline_groups: [],
              background_image: nil
            }
            s.with(sheets: [*s.sheets, new_sheet], active_sheet_idx: s.sheets.size, dirty: true)
          end
          c.run { |sut, _args| sut.add_sheet }
        end

        m.command :add_threaded_comment, weight: 3 do |c|
          c.guard { |s| s.sheets[s.active_sheet_idx][:threaded_comments].size < 3 }
          c.arguments(
            cell: Pbt.constant_from("A1", "B2", "C3"),
            author: Pbt.constant_from("Alice", "Bob", "Carol"),
            text: Pbt.constant_from("Check value", "A < B & \"C\"", "Review needed"),
            done: Pbt.boolean
          )
          c.next_state do |s, args|
            id = format("{00000000-0000-0000-0000-%012d}", s.next_comment_seq)
            new_tc = { cell: args[:cell], text: args[:text], author: args[:author], id: id, parent_id: nil, done: args[:done] }
            sheets = s.sheets.each_with_index.map do |sh, i|
              i == s.active_sheet_idx ? sh.merge(threaded_comments: [*sh[:threaded_comments], new_tc]) : sh
            end
            s.with(sheets: sheets, next_comment_seq: s.next_comment_seq + 1, dirty: true)
          end
          c.run { |sut, args| sut.add_threaded_comment(args[:cell], args[:text], author: args[:author], done: args[:done]) }
        end

        m.command :reply_threaded_comment, weight: 2 do |c|
          c.guard do |s|
            tcs = s.sheets[s.active_sheet_idx][:threaded_comments]
            tcs.any? && tcs.size < 4
          end
          c.arguments(
            author: Pbt.constant_from("Alice", "Bob", "Carol"),
            text: Pbt.constant_from("Reply ok", "Done & verified"),
            done: Pbt.boolean
          )
          c.next_state do |s, args|
            root = s.sheets[s.active_sheet_idx][:threaded_comments].first
            id = format("{00000000-0000-0000-0000-%012d}", s.next_comment_seq)
            new_tc = { cell: root[:cell], text: args[:text], author: args[:author], id: id, parent_id: root[:id], done: args[:done] }
            sheets = s.sheets.each_with_index.map do |sh, i|
              i == s.active_sheet_idx ? sh.merge(threaded_comments: [*sh[:threaded_comments], new_tc]) : sh
            end
            s.with(sheets: sheets, next_comment_seq: s.next_comment_seq + 1, dirty: true)
          end
          c.run { |sut, args| sut.reply_threaded_comment(args[:text], author: args[:author], done: args[:done]) }
        end

        m.command :add_pivot_table, weight: 2 do |c|
          c.guard { |s| s.sheets[s.active_sheet_idx][:pivot_tables].size < 2 }
          c.arguments(
            with_col_field: Pbt.boolean,
            subtotal: Pbt.constant_from("sum", "count", "average", "max", "min")
          )
          c.next_state do |s, args|
            existing_count = s.sheets[s.active_sheet_idx][:pivot_tables].size
            new_pt = {
              source_ref: "Data!A1:C3",
              field_names: %w[Region Quarter Sales],
              row_fields: ["Region"],
              column_fields: (args[:with_col_field] ? ["Quarter"] : []),
              data_fields: [{ field: "Sales", subtotal: args[:subtotal] }],
              ref: "E#{1 + (existing_count * 10)}",
              name: "Pivot#{s.next_pivot_seq}",
              items: { 0 => %w[East West], 1 => %w[Q1 Q2] }
            }
            sheets = s.sheets.each_with_index.map do |sh, i|
              i == s.active_sheet_idx ? sh.merge(pivot_tables: [*sh[:pivot_tables], new_pt]) : sh
            end
            s.with(sheets: sheets, next_pivot_seq: s.next_pivot_seq + 1, dirty: true)
          end
          c.run { |sut, args| sut.add_pivot_table(with_col_field: args[:with_col_field], subtotal: args[:subtotal]) }
        end

        m.command :add_sparkline_group, weight: 2 do |c|
          c.guard { |s| s.sheets[s.active_sheet_idx][:sparkline_groups].size < 2 }
          c.arguments(
            type: Pbt.constant_from(:line, :column, :stacked),
            markers: Pbt.boolean,
            high: Pbt.boolean,
            low: Pbt.boolean
          )
          c.next_state do |s, args|
            existing_count = s.sheets[s.active_sheet_idx][:sparkline_groups].size
            new_sg = {
              sparklines: [{ data_ref: "Data!C2:C3", location_ref: "H#{1 + existing_count}" }],
              type: args[:type],
              markers: args[:markers],
              high: args[:high],
              low: args[:low],
              color_series: "FF336699"
            }
            sheets = s.sheets.each_with_index.map do |sh, i|
              i == s.active_sheet_idx ? sh.merge(sparkline_groups: [*sh[:sparkline_groups], new_sg]) : sh
            end
            s.with(sheets: sheets, dirty: true)
          end
          c.run { |sut, args| sut.add_sparkline_group(type: args[:type], markers: args[:markers], high: args[:high], low: args[:low]) }
        end

        m.command :set_background_image, weight: 1 do |c|
          c.arguments(mode: Pbt.constant_from(:hash, :raw, :clear))
          c.next_state do |s, args|
            val = case args[:mode]
                  when :hash then { data: TINY_PNG, ext: "png" }
                  when :raw then TINY_PNG
                  end
            sheets = s.sheets.each_with_index.map do |sh, i|
              i == s.active_sheet_idx ? sh.merge(background_image: val) : sh
            end
            s.with(sheets: sheets, dirty: true)
          end
          c.run { |sut, args| sut.set_background_image(mode: args[:mode]) }
        end

        m.command :toggle_shared_strings, weight: 1 do |c|
          c.next_state { |s| s.with(shared_strings: !s.shared_strings) }
          c.run { |sut, _args| sut.toggle_shared_strings }
        end

        m.invariant(:auxiliary_parts_integrity) do |state, sut|
          next unless state.dirty

          mem_bytes, stream_bytes = sut.snapshot_bytes
          assert_empty(ExcelRules.check(mem_bytes).map(&:to_s))
          assert_empty(ExcelRules.check(stream_bytes).map(&:to_s))

          assert_threaded_comment_package_integrity!(mem_bytes, "mem")
          assert_threaded_comment_package_integrity!(stream_bytes, "stream")

          m_wb = Xlsxrb.load(StringIO.new(mem_bytes))
          s_wb = Xlsxrb.load(StringIO.new(stream_bytes))
          assert_auxiliary_workbooks_match!(state, m_wb, s_wb, [mem_bytes, stream_bytes])

          resaved = m_wb.to_data(shared_strings: state.shared_strings)
          assert_empty(ExcelRules.check(resaved).map(&:to_s))
          assert_threaded_comment_package_integrity!(resaved, "resaved")
          r_wb = Xlsxrb.load(StringIO.new(resaved))
          assert_auxiliary_workbooks_match!(state, m_wb, r_wb, [resaved, resaved])
        end
      end
    end
  end

  # --- Item 2: Sheet visibility, workbook view, custom/app properties, protection, and Workbook#with ---

  WorkbookMetaState = Data.define(
    :sheets,
    :active_sheet,
    :first_sheet,
    :show_sheet_tabs,
    :tab_ratio,
    :full_calc_on_load,
    :custom_props,
    :app_props,
    :wb_protection,
    :dirty
  )

  # Drives WorkbookBuilder and StreamWriter with workbook-level view, visibility, properties, and protection.
  class WorkbookMetaSut
    def self.resolve_sheet_state(sheets, idx, active_sheet, candidate_state)
      return candidate_state if candidate_state == :visible

      would_leave_no_visible = sheets.each_index.none? { |i| i != idx && sheets[i][:state] == :visible }
      is_explicit_active = (active_sheet.is_a?(Integer) && active_sheet == idx) ||
                           (active_sheet.is_a?(String) && sheets[idx][:name] == active_sheet)
      would_leave_no_visible || is_explicit_active ? :visible : candidate_state
    end

    def self.build_custom_prop(kind, key_idx, int_val, bool_val, str_val)
      name = "Prop#{key_idx}"
      case kind
      when :string then [name, str_val, :string]
      when :number then [name, int_val, :number]
      when :float then [name, (int_val / 4.0).round(2), :float]
      when :boolean then [name, bool_val, :boolean]
      when :date then [name, "2026-06-15T10:30:00Z", :date]
      end
    end

    def initialize
      @sheets = [{ name: "Sheet1", state: :visible }, { name: "Sheet2", state: :visible }]
      @active_sheet = nil
      @first_sheet = nil
      @show_sheet_tabs = nil
      @tab_ratio = nil
      @full_calc_on_load = nil
      @custom_props = {}
      @app_props = {}
      @wb_protection = nil
    end

    def add_sheet(state:)
      idx = @sheets.size + 1
      @sheets << { name: "Sheet#{idx}", state: state }
      nil
    end

    def set_sheet_state(idx:, new_state:)
      target_idx = idx % @sheets.size
      eff_state = self.class.resolve_sheet_state(@sheets, target_idx, @active_sheet, new_state)
      @sheets[target_idx][:state] = eff_state
      nil
    end

    def set_active_sheet(pick:, by_name:, clear:)
      if clear
        @active_sheet = nil
      else
        visible_indices = @sheets.each_index.select { |i| @sheets[i][:state] == :visible }
        vis_idx = visible_indices[pick % visible_indices.size]
        @active_sheet = by_name ? @sheets[vis_idx][:name] : vis_idx
      end
      nil
    end

    def set_first_sheet(pick:, by_name:, clear:)
      if clear
        @first_sheet = nil
      else
        target_idx = pick % @sheets.size
        @first_sheet = by_name ? @sheets[target_idx][:name] : target_idx
      end
      nil
    end

    def set_view_options(show_tabs:, tab_ratio:, full_calc:)
      @show_sheet_tabs = show_tabs
      @tab_ratio = tab_ratio
      @full_calc_on_load = full_calc
      nil
    end

    def set_custom_property(kind:, key_idx:, int_val:, bool_val:, str_val:)
      name, val, type = self.class.build_custom_prop(kind, key_idx, int_val, bool_val, str_val)
      @custom_props[name] = { value: val, type: type }
      nil
    end

    def set_app_properties(company:, manager:, app_name:)
      @app_props = { company: company, manager: manager, app_name: app_name }
      nil
    end

    def set_workbook_protection(lock_structure:, lock_windows:, password:)
      @wb_protection = { lock_structure: lock_structure, lock_windows: lock_windows, password: password }.compact
      nil
    end

    def snapshot_bytes
      wb = Xlsxrb.build do |b|
        apply_workbook_meta(b)
        @sheets.each_with_index do |sh, i|
          b.sheet(sh[:name], state: sh[:state]) { |s| s.row(["Data", i + 1]) }
        end
      end
      mem_bytes = Xlsxrb.write(wb)

      stream_bytes = Xlsxrb.write do |sw|
        apply_workbook_meta(sw)
        @sheets.each_with_index do |sh, i|
          sw.sheet(sh[:name], state: sh[:state]) { |s| s.row(["Data", i + 1]) }
        end
      end
      [mem_bytes, stream_bytes]
    end

    private

    def apply_workbook_meta(target)
      target.active_sheet = @active_sheet unless @active_sheet.nil?
      target.first_sheet = @first_sheet unless @first_sheet.nil?
      target.show_sheet_tabs = @show_sheet_tabs unless @show_sheet_tabs.nil?
      target.tab_ratio = @tab_ratio unless @tab_ratio.nil?
      target.workbook_properties(full_calc_on_load: @full_calc_on_load) unless @full_calc_on_load.nil?
      @custom_props.each do |name, spec|
        target.custom_property(name, spec[:value], type: spec[:type])
      end
      target.app_properties(**@app_props) unless @app_props.empty?
      target.workbook_protection(**@wb_protection) if @wb_protection && !@wb_protection.empty?
    end
  end

  def assert_no_hidden_tab_selected!(bytes, expected_sheets, label)
    zip = Xlsxrb::Ooxml::ZipReader.new(StringIO.new(bytes))
    expected_sheets.each_with_index do |sh, idx|
      xml = zip.read_entry("xl/worksheets/sheet#{idx + 1}.xml")
      has_tab_selected = xml.match?(/<sheetView\b[^>]*\btabSelected="1"/)
      assert_equal(false, has_tab_selected, "[#{label}] Hidden sheet #{sh[:name]} (#{sh[:state]}) has tabSelected=\"1\"") if sh[:state] != :visible
    end
  end

  def assert_workbook_meta_matches!(state, workbook, label)
    assert_equal(state.sheets.map { |sh| [sh[:name], sh[:state]] }, workbook.sheets.map { |ws| [ws.name, ws.state] }, "[#{label}] sheet states")
    workbook.sheets.each_with_index do |ws, idx|
      assert_equal(state.sheets[idx][:state] != :visible, ws.hidden?, "[#{label}] sheet #{idx} hidden?")
    end

    exp_active_idx = if state.active_sheet.is_a?(Integer)
                       state.active_sheet
                     elsif state.active_sheet.is_a?(String)
                       state.sheets.find_index { |sh| sh[:name] == state.active_sheet }
                     else
                       state.sheets.find_index { |sh| sh[:state] == :visible }
                     end
    assert_equal(state.sheets[exp_active_idx][:name], workbook.active_sheet&.name, "[#{label}] active_sheet")
    assert_equal(exp_active_idx, workbook.active_sheet_index, "[#{label}] active_sheet_index") if exp_active_idx&.positive?

    exp_first_idx = if state.first_sheet.is_a?(Integer)
                      state.first_sheet
                    elsif state.first_sheet.is_a?(String)
                      state.sheets.find_index { |sh| sh[:name] == state.first_sheet }
                    end
    assert_equal(exp_first_idx, workbook.first_sheet, "[#{label}] first_sheet")
    assert_equal(state.show_sheet_tabs != false, workbook.show_sheet_tabs?, "[#{label}] show_sheet_tabs?")
    assert_equal(state.tab_ratio, workbook.tab_ratio, "[#{label}] tab_ratio")
    assert_equal(state.full_calc_on_load == true, workbook.full_calc_on_load?, "[#{label}] full_calc_on_load?")

    assert_equal(state.custom_props.keys.sort, workbook.custom_properties.keys.sort, "[#{label}] custom_properties keys")
    state.custom_props.each do |name, spec|
      cp = workbook.custom_properties.fetch(name)
      exp_type = spec[:type] == :boolean ? :bool : spec[:type]
      assert_equal(exp_type, cp.type, "[#{label}] custom property #{name} type")
      if exp_type == :float
        assert_in_delta(spec[:value], cp.value, 1e-6, "[#{label}] custom property #{name} float value")
      else
        assert_equal(spec[:value], cp.value, "[#{label}] custom property #{name} value")
      end
    end

    if state.app_props.empty?
      assert_nil(workbook.app_properties.company)
      assert_nil(workbook.app_properties.manager)
    else
      assert_equal(state.app_props[:company], workbook.app_properties.company, "[#{label}] app company")
      assert_equal(state.app_props[:manager], workbook.app_properties.manager, "[#{label}] app manager")
      assert_equal(state.app_props[:app_name], workbook.app_properties.app_name, "[#{label}] app_name")
    end

    return unless state.wb_protection
    return unless state.wb_protection[:lock_structure] || state.wb_protection[:lock_windows] || state.wb_protection[:password]

    wp = workbook.workbook_protection
    assert_not_nil(wp, "[#{label}] workbook_protection")
    assert_equal(state.wb_protection[:lock_structure] ? true : false, wp.lock_structure?, "[#{label}] lock_structure?")
    assert_equal(state.wb_protection[:lock_windows] ? true : false, wp.lock_windows?, "[#{label}] lock_windows?")
    assert_not_nil(wp.hash_value, "[#{label}] workbook protection password hash") if state.wb_protection[:password]
  end

  test "stateful property: sheet visibility, workbook view, properties, protection, and Workbook#with lifecycle" do
    init_state = WorkbookMetaState.new(
      sheets: [{ name: "Sheet1", state: :visible }, { name: "Sheet2", state: :visible }],
      active_sheet: nil,
      first_sheet: nil,
      show_sheet_tabs: nil,
      tab_ratio: nil,
      full_calc_on_load: nil,
      custom_props: {},
      app_props: {},
      wb_protection: nil,
      dirty: false
    )

    Pbt.assert(num_runs: 18) do
      Pbt.stateful(max_steps: 10) do |m|
        m.initial_state Pbt.constant(init_state)
        m.sut { |_init| WorkbookMetaSut.new }

        m.command :add_sheet, weight: 2 do |c|
          c.guard { |s| s.sheets.size < 4 }
          c.arguments(state: Pbt.constant_from(:visible, :hidden, :very_hidden))
          c.next_state do |s, args|
            idx = s.sheets.size + 1
            s.with(sheets: [*s.sheets, { name: "Sheet#{idx}", state: args[:state] }], dirty: true)
          end
          c.run { |sut, args| sut.add_sheet(state: args[:state]) }
        end

        m.command :set_sheet_state, weight: 3 do |c|
          c.guard { |s| s.sheets.size >= 2 }
          c.arguments(idx: Pbt.integer(min: 0, max: 3), new_state: Pbt.constant_from(:visible, :hidden, :very_hidden))
          c.next_state do |s, args|
            target_idx = args[:idx] % s.sheets.size
            eff = WorkbookMetaSut.resolve_sheet_state(s.sheets, target_idx, s.active_sheet, args[:new_state])
            sheets = s.sheets.each_with_index.map do |sh, i|
              i == target_idx ? sh.merge(state: eff) : sh
            end
            s.with(sheets: sheets, dirty: true)
          end
          c.run { |sut, args| sut.set_sheet_state(idx: args[:idx], new_state: args[:new_state]) }
        end

        m.command :set_active_sheet, weight: 2 do |c|
          c.arguments(pick: Pbt.integer(min: 0, max: 3), by_name: Pbt.boolean, clear: Pbt.boolean)
          c.next_state do |s, args|
            val = if args[:clear]
                    nil
                  else
                    visible_indices = s.sheets.each_index.select { |i| s.sheets[i][:state] == :visible }
                    vis_idx = visible_indices[args[:pick] % visible_indices.size]
                    args[:by_name] ? s.sheets[vis_idx][:name] : vis_idx
                  end
            s.with(active_sheet: val, dirty: true)
          end
          c.run { |sut, args| sut.set_active_sheet(pick: args[:pick], by_name: args[:by_name], clear: args[:clear]) }
        end

        m.command :set_first_sheet, weight: 1 do |c|
          c.arguments(pick: Pbt.integer(min: 0, max: 3), by_name: Pbt.boolean, clear: Pbt.boolean)
          c.next_state do |s, args|
            val = if args[:clear]
                    nil
                  else
                    target_idx = args[:pick] % s.sheets.size
                    args[:by_name] ? s.sheets[target_idx][:name] : target_idx
                  end
            s.with(first_sheet: val, dirty: true)
          end
          c.run { |sut, args| sut.set_first_sheet(pick: args[:pick], by_name: args[:by_name], clear: args[:clear]) }
        end

        m.command :set_view_options, weight: 2 do |c|
          c.arguments(
            show_tabs: Pbt.constant_from(nil, true, false),
            tab_ratio: Pbt.constant_from(nil, 250, 600, 850),
            full_calc: Pbt.constant_from(nil, true, false)
          )
          c.next_state do |s, args|
            s.with(show_sheet_tabs: args[:show_tabs], tab_ratio: args[:tab_ratio], full_calc_on_load: args[:full_calc], dirty: true)
          end
          c.run { |sut, args| sut.set_view_options(show_tabs: args[:show_tabs], tab_ratio: args[:tab_ratio], full_calc: args[:full_calc]) }
        end

        m.command :set_custom_property, weight: 3 do |c|
          c.arguments(
            kind: Pbt.constant_from(:string, :number, :float, :boolean, :date),
            key_idx: Pbt.integer(min: 0, max: 4),
            int_val: Pbt.integer(min: -50, max: 500),
            bool_val: Pbt.boolean,
            str_val: Pbt.constant_from("Alpha", "A & B <C>", "v1.0")
          )
          c.next_state do |s, args|
            name, val, type = WorkbookMetaSut.build_custom_prop(args[:kind], args[:key_idx], args[:int_val], args[:bool_val], args[:str_val])
            s.with(custom_props: s.custom_props.merge(name => { value: val, type: type }), dirty: true)
          end
          c.run do |sut, args|
            sut.set_custom_property(kind: args[:kind], key_idx: args[:key_idx], int_val: args[:int_val], bool_val: args[:bool_val], str_val: args[:str_val])
          end
        end

        m.command :set_app_properties, weight: 2 do |c|
          c.arguments(
            company: Pbt.constant_from("Acme & Co", "Beta Corp"),
            manager: Pbt.constant_from("Jane Doe", "Taro"),
            app_name: Pbt.constant_from("XlsxrbTest", "ReportGen")
          )
          c.next_state do |s, args|
            s.with(app_props: { company: args[:company], manager: args[:manager], app_name: args[:app_name] }, dirty: true)
          end
          c.run { |sut, args| sut.set_app_properties(company: args[:company], manager: args[:manager], app_name: args[:app_name]) }
        end

        m.command :set_workbook_protection, weight: 2 do |c|
          c.arguments(
            lock_structure: Pbt.boolean,
            lock_windows: Pbt.boolean,
            password: Pbt.constant_from(nil, "secret1", "pw99")
          )
          c.next_state do |s, args|
            s.with(wb_protection: { lock_structure: args[:lock_structure], lock_windows: args[:lock_windows], password: args[:password] }.compact, dirty: true)
          end
          c.run { |sut, args| sut.set_workbook_protection(lock_structure: args[:lock_structure], lock_windows: args[:lock_windows], password: args[:password]) }
        end

        m.invariant(:workbook_meta_parity_and_roundtrip) do |state, sut|
          next unless state.dirty

          mem_bytes, stream_bytes = sut.snapshot_bytes
          assert_empty(ExcelRules.check(mem_bytes).map(&:to_s))
          assert_empty(ExcelRules.check(stream_bytes).map(&:to_s))

          assert_no_hidden_tab_selected!(mem_bytes, state.sheets, "mem")
          assert_no_hidden_tab_selected!(stream_bytes, state.sheets, "stream")

          m_wb = Xlsxrb.load(StringIO.new(mem_bytes))
          s_wb = Xlsxrb.load(StringIO.new(stream_bytes))
          assert_workbook_meta_matches!(state, m_wb, "mem")
          assert_workbook_meta_matches!(state, s_wb, "stream")

          resaved = m_wb.to_data
          assert_empty(ExcelRules.check(resaved).map(&:to_s))
          r_wb = Xlsxrb.load(StringIO.new(resaved))
          assert_workbook_meta_matches!(state, r_wb, "resaved")

          mutated = m_wb.with(
            workbook_properties: m_wb.workbook_properties.with(show_sheet_tabs: !m_wb.show_sheet_tabs?),
            app_properties: { company: "MutatedCo" }
          )
          mut_bytes = mutated.to_data
          assert_empty(ExcelRules.check(mut_bytes).map(&:to_s))
          mw = Xlsxrb.load(StringIO.new(mut_bytes))
          assert_equal(!m_wb.show_sheet_tabs?, mw.show_sheet_tabs?)
          assert_equal("MutatedCo", mw.app_properties.company)
        end
      end
    end
  end

  # --- Item 3: Advanced Table, auto_filter / filter_column (all 6 types + flags), and sort_state ---

  TableFilterSortState = Data.define(:sheets, :active_sheet_idx, :next_table_seq, :dirty)

  # Drives WorkbookBuilder and StreamWriter with tables, auto filters, filter columns, and sort state.
  class TableFilterSortSut
    def self.build_table_spec(seq, slot_idx, args)
      totals_row = args[:totals_row]
      end_row = totals_row ? 6 : 5
      ref = slot_idx.zero? ? "F1:I#{end_row}" : "K1:N#{end_row}"
      name = "Tbl#{seq}"
      col0 = { name: "Item", totals_row_label: (totals_row ? "Total" : nil) }.compact
      col1 = { name: "Qty", totals_row_function: (totals_row ? args[:tot_fn] : nil) }.compact
      col2 = { name: "Price", calculated_column_formula: (args[:with_calc_formula] ? "#{name}[[#This Row],[Qty]]*10" : nil) }.compact
      col3 = if totals_row && args[:with_tot_formula]
               { name: "Amount", totals_row_formula: "SUM(#{name}[Qty])" }
             elsif totals_row
               { name: "Amount", totals_row_function: "sum" }
             else
               { name: "Amount" }
             end
      {
        ref: ref,
        name: name,
        display_name: name,
        columns: [col0, col1, col2, col3],
        totals_row_count: (totals_row ? 1 : 0),
        auto_filter: args[:auto_filter],
        style: args[:style],
        show_first_column: args[:show_first],
        show_last_column: args[:show_last],
        show_row_stripes: args[:show_row_stripes],
        show_column_stripes: args[:show_col_stripes]
      }
    end

    def self.build_filter_spec(args)
      spec = case args[:kind]
             when :filters
               { type: :filters, values: %w[Alpha Beta], blank: (args[:flag_a] ? true : nil) }.compact
             when :custom_single
               { type: :custom, operator: "greaterThan", val: "10" }
             when :custom_multi
               { type: :custom, filters: [{ operator: "greaterThanOrEqual", val: "5" }, { operator: "lessThan", val: "50" }], and: args[:flag_a] }
             when :dynamic
               { type: :dynamic, dynamic_type: (args[:flag_a] ? "aboveAverage" : "today") }
             when :top10
               { type: :top10, top: args[:flag_a], percent: args[:flag_b], val: 5 }
             when :color_filter
               { type: :color_filter, dxf_id: 0, cell_color: (args[:flag_a] ? nil : false) }.compact
             when :icon_filter
               { type: :icon_filter, icon_set: "3Arrows", icon_id: (args[:flag_a] ? 1 : 2) }
             end
      spec = spec.dup
      spec[:hidden] = true if args[:hidden]
      spec[:show_button] = false unless args[:show_button]
      spec
    end

    def self.build_sort_spec(args)
      conds = [{ ref: "B2:B5", descending: args[:desc0] }]
      conds << { ref: "C2:C5", descending: args[:desc1] } if args[:two_conds]
      {
        ref: "A2:D5",
        sort_conditions: conds,
        case_sensitive: args[:case_sensitive],
        sort_method: args[:sort_method]
      }
    end

    def initialize
      @sheets = [{ name: "Sheet1", tables: [], auto_filter: nil, filter_columns: {}, sort_state: nil, needs_dxf: false }]
      @active_sheet_idx = 0
      @next_table_seq = 1
    end

    def add_sheet
      idx = @sheets.size + 1
      @sheets << { name: "Sheet#{idx}", tables: [], auto_filter: nil, filter_columns: {}, sort_state: nil, needs_dxf: false }
      @active_sheet_idx = @sheets.size - 1
      nil
    end

    def add_table(args)
      slot_idx = @sheets[@active_sheet_idx][:tables].size
      spec = self.class.build_table_spec(@next_table_seq, slot_idx, args)
      @next_table_seq += 1
      @sheets[@active_sheet_idx][:tables] << spec
      nil
    end

    def set_auto_filter
      @sheets[@active_sheet_idx][:auto_filter] = "A1:D5"
      nil
    end

    def apply_filter_column(args)
      spec = self.class.build_filter_spec(args)
      @sheets[@active_sheet_idx][:filter_columns][args[:col_id]] = spec
      @sheets[@active_sheet_idx][:needs_dxf] = true if args[:kind] == :color_filter
      nil
    end

    def apply_sort_state(args)
      @sheets[@active_sheet_idx][:sort_state] = self.class.build_sort_spec(args)
      nil
    end

    def snapshot_bytes
      wb = Xlsxrb.build do |b|
        @sheets.each { |spec| b.sheet(spec[:name]) { |s| populate_sheet(s, spec) } }
      end
      mem_bytes = Xlsxrb.write(wb)

      stream_bytes = Xlsxrb.write do |sw|
        @sheets.each { |spec| sw.sheet(spec[:name]) { |s| populate_sheet(s, spec) } }
      end
      [mem_bytes, stream_bytes]
    end

    private

    def populate_sheet(sheet, spec)
      sheet.row(["Item", "Qty", "Price", "Total", "", "Item", "Qty", "Price", "Amount", "", "Item", "Qty", "Price", "Amount"])
      4.times do |i|
        sheet.row(["Alpha", 10 + i, 2.5, 25 + i, "", "A#{i}", 1 + i, 10, 10 + i, "", "B#{i}", 2 + i, 10, 20 + i])
      end
      sheet.row(["End", 0, 0, 0, "", "Total", 10, 0, 10, "", "Total", 14, 0, 14])

      sheet.conditional_format("A2:A5", type: :cell_is, operator: :greater_than, formula: "0", fill_color: "FFEEEE") if spec[:needs_dxf]
      spec[:tables].each do |tbl|
        sheet.table(
          tbl[:ref],
          columns: tbl[:columns],
          name: tbl[:name],
          display_name: tbl[:display_name],
          style: tbl[:style],
          totals_row_count: tbl[:totals_row_count],
          auto_filter: tbl[:auto_filter],
          show_first_column: tbl[:show_first_column],
          show_last_column: tbl[:show_last_column],
          show_row_stripes: tbl[:show_row_stripes],
          show_column_stripes: tbl[:show_column_stripes]
        )
      end
      return unless spec[:auto_filter]

      sheet.auto_filter = spec[:auto_filter]
      spec[:filter_columns].each { |col_id, fspec| sheet.filter_column(col_id, fspec) }
      return unless spec[:sort_state]

      ss = spec[:sort_state]
      sheet.sort_state(ss[:ref], ss[:sort_conditions], case_sensitive: ss[:case_sensitive], sort_method: ss[:sort_method])
    end
  end

  def assert_table_filter_sort_sheet_matches!(exp_sheet, actual_ws, expected_table_ids, label)
    assert_equal(expected_table_ids, actual_ws.tables.map(&:id), "[#{label}] table IDs")
    assert_equal(exp_sheet[:tables].size, actual_ws.tables.size, "[#{label}] table count")

    exp_sheet[:tables].zip(actual_ws.tables).each_with_index do |(exp_t, act_t), ti|
      assert_equal(exp_t[:name], act_t.name, "[#{label}] table[#{ti}].name")
      assert_equal(exp_t[:display_name], act_t.display_name, "[#{label}] table[#{ti}].display_name")
      assert_equal(exp_t[:ref], act_t.ref, "[#{label}] table[#{ti}].ref")
      assert_equal(exp_t[:totals_row_count], act_t.totals_row_count || 0, "[#{label}] table[#{ti}].totals_row_count")
      assert_equal(exp_t[:style], act_t.style_name, "[#{label}] table[#{ti}].style_name")
      assert_equal(exp_t[:show_first_column] ? true : false, act_t.show_first_column?, "[#{label}] table[#{ti}].show_first_column?")
      assert_equal(exp_t[:show_last_column] ? true : false, act_t.show_last_column?, "[#{label}] table[#{ti}].show_last_column?")
      assert_equal(exp_t[:show_row_stripes] ? true : false, act_t.show_row_stripes?, "[#{label}] table[#{ti}].show_row_stripes?")
      assert_equal(exp_t[:show_column_stripes] ? true : false, act_t.show_column_stripes?, "[#{label}] table[#{ti}].show_column_stripes?")
      if exp_t[:auto_filter]
        assert_not_nil(act_t.auto_filter, "[#{label}] table[#{ti}].auto_filter expected present")
      else
        assert_nil(act_t.auto_filter, "[#{label}] table[#{ti}].auto_filter expected nil")
      end

      assert_equal(exp_t[:columns].size, act_t.columns.size, "[#{label}] table[#{ti}].columns size")
      exp_t[:columns].zip(act_t.columns).each_with_index do |(ec, ac), ci|
        assert_equal(ec[:name], ac.name, "[#{label}] table[#{ti}].col[#{ci}].name")
        assert_equal(ec[:totals_row_label], ac.totals_row_label, "[#{label}] table[#{ti}].col[#{ci}].totals_row_label")
        exp_trf = ec[:totals_row_formula] ? "custom" : ec[:totals_row_function]
        assert_equal(exp_trf, ac.totals_row_function, "[#{label}] table[#{ti}].col[#{ci}].totals_row_function")
        assert_equal(ec[:calculated_column_formula], ac.calculated_column_formula, "[#{label}] table[#{ti}].col[#{ci}].calculated_column_formula")
        assert_equal(ec[:totals_row_formula], ac.totals_row_formula, "[#{label}] table[#{ti}].col[#{ci}].totals_row_formula")
      end
    end

    assert_equal(exp_sheet[:auto_filter], actual_ws.auto_filter, "[#{label}] auto_filter")
    assert_equal(exp_sheet[:filter_columns].keys.sort, actual_ws.filter_columns.keys.sort, "[#{label}] filter_column keys")
    exp_sheet[:filter_columns].each do |col_id, exp_fc|
      act_fc = actual_ws.filter_columns.fetch(col_id)
      assert_equal(exp_fc[:hidden] == true, act_fc.hidden?, "[#{label}] filter_column[#{col_id}].hidden?")
      assert_equal(exp_fc[:show_button] != false, act_fc.show_button?, "[#{label}] filter_column[#{col_id}].show_button?")
      assert_equal(exp_fc, act_fc.to_spec, "[#{label}] filter_column[#{col_id}].to_spec")
    end

    if exp_sheet[:sort_state].nil?
      assert_nil(actual_ws.sort_state, "[#{label}] sort_state")
    else
      exp_ss = exp_sheet[:sort_state]
      act_ss = actual_ws.sort_state
      assert_not_nil(act_ss, "[#{label}] sort_state")
      assert_equal(exp_ss[:ref], act_ss.ref, "[#{label}] sort_state.ref")
      assert_equal(exp_ss[:case_sensitive] == true, act_ss.case_sensitive?, "[#{label}] sort_state.case_sensitive?")
      assert_equal(exp_ss[:sort_method], act_ss.sort_method, "[#{label}] sort_state.sort_method")
      exp_conds = exp_ss[:sort_conditions].map { |sc| [sc[:ref], sc[:descending] ? true : false] }
      act_conds = (act_ss.conditions || []).map { |sc| [sc[:ref], sc[:descending] ? true : false] }
      assert_equal(exp_conds, act_conds, "[#{label}] sort_state.conditions")
    end
  end

  def assert_table_filter_sort_workbook_matches!(state, workbook, label)
    next_id = 1
    state.sheets.each_with_index do |exp_sheet, idx|
      expected_ids = exp_sheet[:tables].map do
        id = next_id
        next_id += 1
        id
      end
      assert_table_filter_sort_sheet_matches!(exp_sheet, workbook.sheet(idx), expected_ids, "#{label} sheet #{idx}")
    end
  end

  def assert_table_filter_sort_stream_matches!(state, bytes, label)
    next_id = 1
    sheet_idx = 0
    Xlsxrb.read(StringIO.new(bytes)) do |stream_sheet|
      exp_sheet = state.sheets[sheet_idx]
      expected_ids = exp_sheet[:tables].map do
        id = next_id
        next_id += 1
        id
      end
      assert_table_filter_sort_sheet_matches!(exp_sheet, stream_sheet, expected_ids, "#{label} stream sheet #{sheet_idx}")
      sheet_idx += 1
    end
    assert_equal(state.sheets.size, sheet_idx)
  end

  test "stateful property: advanced tables, filter columns, and sort state parity and roundtrip" do
    init_state = TableFilterSortState.new(
      sheets: [{ name: "Sheet1", tables: [], auto_filter: nil, filter_columns: {}, sort_state: nil, needs_dxf: false }],
      active_sheet_idx: 0,
      next_table_seq: 1,
      dirty: false
    )

    Pbt.assert(num_runs: 15) do
      Pbt.stateful(max_steps: 10) do |m|
        m.initial_state Pbt.constant(init_state)
        m.sut { |_init| TableFilterSortSut.new }

        m.command :add_sheet, weight: 2 do |c|
          c.guard { |s| s.sheets.size < 2 }
          c.next_state do |s|
            new_sheet = { name: "Sheet#{s.sheets.size + 1}", tables: [], auto_filter: nil, filter_columns: {}, sort_state: nil, needs_dxf: false }
            s.with(sheets: [*s.sheets, new_sheet], active_sheet_idx: s.sheets.size, dirty: true)
          end
          c.run { |sut, _args| sut.add_sheet }
        end

        m.command :add_table, weight: 3 do |c|
          c.guard { |s| s.sheets[s.active_sheet_idx][:tables].size < 2 }
          c.arguments(
            totals_row: Pbt.boolean,
            auto_filter: Pbt.boolean,
            style: Pbt.constant_from("TableStyleMedium2", "TableStyleLight9", "TableStyleDark4"),
            show_first: Pbt.boolean,
            show_last: Pbt.boolean,
            show_row_stripes: Pbt.boolean,
            show_col_stripes: Pbt.boolean,
            tot_fn: Pbt.constant_from("sum", "average", "count", "min", "max"),
            with_calc_formula: Pbt.boolean,
            with_tot_formula: Pbt.boolean
          )
          c.next_state do |s, args|
            slot_idx = s.sheets[s.active_sheet_idx][:tables].size
            new_tbl = TableFilterSortSut.build_table_spec(s.next_table_seq, slot_idx, args)
            sheets = s.sheets.each_with_index.map do |sh, i|
              i == s.active_sheet_idx ? sh.merge(tables: [*sh[:tables], new_tbl]) : sh
            end
            s.with(sheets: sheets, next_table_seq: s.next_table_seq + 1, dirty: true)
          end
          c.run { |sut, args| sut.add_table(args) }
        end

        m.command :set_auto_filter, weight: 3 do |c|
          c.guard { |s| s.sheets[s.active_sheet_idx][:auto_filter].nil? }
          c.next_state do |s|
            sheets = s.sheets.each_with_index.map do |sh, i|
              i == s.active_sheet_idx ? sh.merge(auto_filter: "A1:D5") : sh
            end
            s.with(sheets: sheets, dirty: true)
          end
          c.run { |sut, _args| sut.set_auto_filter }
        end

        m.command :set_filter_column, weight: 4 do |c|
          c.guard { |s| !s.sheets[s.active_sheet_idx][:auto_filter].nil? }
          c.arguments(
            col_id: Pbt.integer(min: 0, max: 3),
            kind: Pbt.constant_from(:filters, :custom_single, :custom_multi, :dynamic, :top10, :color_filter, :icon_filter),
            flag_a: Pbt.boolean,
            flag_b: Pbt.boolean,
            hidden: Pbt.boolean,
            show_button: Pbt.boolean
          )
          c.next_state do |s, args|
            spec = TableFilterSortSut.build_filter_spec(args)
            sheets = s.sheets.each_with_index.map do |sh, i|
              if i == s.active_sheet_idx
                sh.merge(filter_columns: sh[:filter_columns].merge(args[:col_id] => spec), needs_dxf: sh[:needs_dxf] || args[:kind] == :color_filter)
              else
                sh
              end
            end
            s.with(sheets: sheets, dirty: true)
          end
          c.run { |sut, args| sut.apply_filter_column(args) }
        end

        m.command :set_sort_state, weight: 2 do |c|
          c.guard { |s| !s.sheets[s.active_sheet_idx][:auto_filter].nil? }
          c.arguments(
            case_sensitive: Pbt.boolean,
            sort_method: Pbt.constant_from(nil, "pinYin", "stroke"),
            desc0: Pbt.boolean,
            two_conds: Pbt.boolean,
            desc1: Pbt.boolean
          )
          c.next_state do |s, args|
            ss = TableFilterSortSut.build_sort_spec(args)
            sheets = s.sheets.each_with_index.map do |sh, i|
              i == s.active_sheet_idx ? sh.merge(sort_state: ss) : sh
            end
            s.with(sheets: sheets, dirty: true)
          end
          c.run { |sut, args| sut.apply_sort_state(args) }
        end

        m.invariant(:table_filter_sort_parity_and_roundtrip) do |state, sut|
          next unless state.dirty

          mem_bytes, stream_bytes = sut.snapshot_bytes
          assert_empty(ExcelRules.check(mem_bytes).map(&:to_s))
          assert_empty(ExcelRules.check(stream_bytes).map(&:to_s))

          m_wb = Xlsxrb.load(StringIO.new(mem_bytes))
          s_wb = Xlsxrb.load(StringIO.new(stream_bytes))
          assert_table_filter_sort_workbook_matches!(state, m_wb, "mem")
          assert_table_filter_sort_workbook_matches!(state, s_wb, "stream")

          assert_table_filter_sort_stream_matches!(state, mem_bytes, "mem")
          assert_table_filter_sort_stream_matches!(state, stream_bytes, "stream")

          resaved = m_wb.to_data
          assert_empty(ExcelRules.check(resaved).map(&:to_s))
          r_wb = Xlsxrb.load(StringIO.new(resaved))
          assert_table_filter_sort_workbook_matches!(state, r_wb, "resaved")
        end
      end
    end
  end
end
