# frozen_string_literal: true

require "test_helper"
require "tempfile"

class RoundTripFidelityTest < Test::Unit::TestCase
  MINIMAL_PNG = [
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

  def test_round_trip_preserves_workbook_and_worksheet_attributes
    wb = Xlsxrb.build do |b|
      b.core_property(:creator, "Author One")
      b.core_property(:title, "Comprehensive Round-Trip")
      b.app_property(:company, "Engineering Org")
      b.custom_property("Classification", "Internal")
      b.workbook_protection(lock_structure: true)
      b.defined_name("DataRange", "Sheet1!$A$1:$B$2")
      b.sheet("Sheet1") do |s|
        s.freeze_pane(row: 1, column: 1)
        s.sheet_view(:show_grid_lines, false)
        s.sheet_view(:tab_selected, true)
        s.sheet_view(:zoom_scale, 120)
        s.sheet_properties(:tab_color, "FFFF0000")
        s.sheet_format(default_row_height: 20)
        s.page_setup(orientation: "landscape", paper_size: 9)
        s.page_margins(left: 0.7, right: 0.7, top: 0.75, bottom: 0.75)
        s.header_footer(odd_header: "&CTest Header", odd_footer: "&RPage &P")
        s.print_options(:grid_lines, true)
        s.sheet_protection(sheet: true)
        s.merge("C1:D2")
        s.auto_filter("A1:D1")
        s.conditional_format("A1:A5", type: :cellIs, operator: :greaterThan, formula1: "10")
        s.data_validation("B1:B5", type: :whole, operator: :between, formula1: "1", formula2: "100")
        s.row_break(5)
        s.column_break(3)
        s.hyperlink("A2", "https://example.com", tooltip: "Example")
        s.comment("B2", "Audited cell", author: "QA")
        s.image(MINIMAL_PNG, ext: "png", from_column: 4, from_row: 4, to_column: 6, to_row: 6)
        s.table("A1:B5", columns: %w[ColA ColB], name: "RoundTripTable")

        s.row %w[ColA ColB Merged1 Merged2]
        s.row [15, 50, "MergedVal", ""]
        s.row [5, 20, "", ""]
        s.row [25, 80, "", ""]
        s.row [35, 10, "", ""]
        s.row ["FooterRow", "End", "", ""]
      end
      b.sheet("Sheet2", state: :hidden) do |s|
        s.row ["HiddenContent"]
      end
    end

    tmp1 = Tempfile.new(["fidelity_initial", ".xlsx"])
    tmp2 = Tempfile.new(["fidelity_roundtrip_stream", ".xlsx"])
    tmp3 = Tempfile.new(["fidelity_roundtrip_loaded", ".xlsx"])

    begin
      Xlsxrb.write(tmp1.path, wb)

      # 1. Read back initially
      read1 = Xlsxrb.read(tmp1.path)

      # 2. Round-trip via StreamSheet
      Xlsxrb.write(tmp2.path, read1)
      read_from_stream = Xlsxrb.read(tmp2.path)

      # 3. Round-trip via loaded in-memory Worksheet
      loaded_wb = Xlsxrb.read(tmp1.path)
      loaded_wb.sheets.each(&:load)
      Xlsxrb.write(tmp3.path, loaded_wb)
      read_from_loaded = Xlsxrb.read(tmp3.path)

      [read_from_stream, read_from_loaded].each do |roundtrip_wb|
        assert_equal "Author One", roundtrip_wb.core_properties.creator
        assert_equal "Comprehensive Round-Trip", roundtrip_wb.core_properties.title
        assert_equal "Engineering Org", roundtrip_wb.app_properties.company
        assert_equal({ "Classification" => { name: "Classification", value: "Internal", type: :string } }, roundtrip_wb.custom_properties.transform_values(&:to_h))
        assert_equal({ lock_structure: true }, roundtrip_wb.workbook_protection.to_h)
        assert_equal :hidden, roundtrip_wb.sheet(1).state
        assert_equal([{ name: "DataRange", hidden: false, value: "Sheet1!$A$1:$B$2" }], roundtrip_wb.defined_names.map(&:to_h))

        s = roundtrip_wb.sheet(0)
        assert_equal({ row: 1, column: 1, col: 1, state: :frozen, active_pane: "bottomRight" }, s.freeze_pane.to_h)
        assert_equal({ show_grid_lines: false, tab_selected: true, zoom_scale: 120 }, s.sheet_view.to_h)
        assert_equal({ orientation: "landscape", paper_size: 9 }, s.page_setup.to_h)
        assert_equal(0.7, s.page_margins.left)
        assert_equal(0.7, s.page_margins.right)
        assert_equal(0.75, s.page_margins.top)
        assert_equal(0.75, s.page_margins.bottom)
        assert_equal({ odd_header: "&CTest Header", odd_footer: "&RPage &P" }, s.header_footer.to_h)
        assert_equal({ grid_lines: true }, s.print_options.to_h)
        assert_equal({ sheet: true }, s.sheet_protection.to_h)
        assert_equal ["C1:D2"], s.merged_cells
        assert_equal "A1:D1", s.auto_filter
        assert_equal [{ id: 5, max: 16_383, man: true }], s.row_breaks.map(&:to_h)
        assert_equal [{ id: 3, max: 1_048_575, man: true }], s.column_breaks.map(&:to_h)
        assert_equal({ "A2" => { url: "https://example.com", tooltip: "Example" } }, s.hyperlinks.transform_values(&:to_h))
        assert_equal [{ ref: "B2", text: "Audited cell", author: "QA", visible: false }], s.comments.map(&:to_h)
        assert_equal 1, s.images.size
        assert_equal 1, s.tables.size
        assert_equal "RoundTripTable", s.tables.first.name
        assert_equal "A1:B5", s.tables.first[:ref]
        assert_equal [{ name: "ColA" }, { name: "ColB" }], s.tables.first[:columns]
        assert_equal [{ sqref: "A1:A5", type: "cellIs", priority: 1, operator: "greaterThan" }], s.conditional_formatting
        assert_equal [{ sqref: "B1:B5", type: "whole", operator: "between", formula1: "1", formula2: "100" }], s.data_validations
        assert_equal({ tab_color: "FFFF0000" }, s.sheet_properties)
        assert_equal({ default_row_height: 20.0 }, s.sheet_format)
      end
    ensure
      tmp1.close!
      tmp2.close!
      tmp3.close!
    end
  end

  def test_phase3_features_round_trip_via_load_and_read
    wb = Xlsxrb.build do |b|
      b.sheet("Analytics") do |s|
        s.add_row(%w[Month Sales Expenses])
        s.add_row(["Jan", 100, 80])
        s.add_row(["Feb", 120, 90])
        s.add_row(["Mar", 150, 110])
        s.print_area("A1:C4")
        s.print_titles(rows: "1:1", columns: "A:A")
        s.sparkline_group(
          sparklines: [{ location: "D2", data: "B2:C2" }],
          type: :line,
          markers: true,
          high: true,
          color_series: "FF0000FF",
          manual_max: 200,
          line_weight: 1.5
        )
        s.chart(:bar) do |c|
          c.title = "Quarterly Sales"
          c.series(name: "Sales", values: "Analytics!$B$2:$B$4", categories: "Analytics!$A$2:$A$4")
        end
      end
    end

    tmp = Tempfile.new(["phase3_features", ".xlsx"])
    begin
      Xlsxrb.write(tmp.path, wb)

      # 1. Test via Xlsxrb.load
      loaded = Xlsxrb.load(tmp.path)
      ws = loaded.sheet(0)
      assert_equal("A1:C4", ws.print_area)
      assert_not_nil(ws.print_titles)
      assert_equal("1:1", ws.print_titles.rows)
      assert_equal("A:A", ws.print_titles.columns)

      assert_equal(1, ws.sparkline_groups.size)
      sg = ws.sparkline_groups.first
      assert_instance_of(Xlsxrb::Elements::SparklineGroup, sg)
      assert_equal("line", sg.type)
      assert_equal(true, sg.markers)
      assert_equal(true, sg.high)
      assert_equal("FF0000FF", sg.color_series)
      assert_equal(200, sg.max)
      assert_equal(1.5, sg.line_weight)
      assert_equal(1, sg.sparklines.size)
      assert_equal("D2", sg.sparklines.first.sqref)

      assert_equal(1, ws.charts.size)
      ch = ws.charts.first
      assert_instance_of(Xlsxrb::Elements::Chart, ch)
      assert_equal("barChart", ch.type)
      assert_equal("Quarterly Sales", ch.title)

      # 2. Test via Xlsxrb.read (StreamSheet streaming mode)
      read_sheets = []
      Xlsxrb.read(tmp.path) do |stream_sheet|
        read_sheets << stream_sheet
        assert_equal("A1:C4", stream_sheet.print_area)
        assert_not_nil(stream_sheet.print_titles)
        assert_equal("1:1", stream_sheet.print_titles.rows)
        assert_equal("A:A", stream_sheet.print_titles.columns)

        assert_equal(1, stream_sheet.sparkline_groups.size)
        st_sg = stream_sheet.sparkline_groups.first
        assert_instance_of(Xlsxrb::Elements::SparklineGroup, st_sg)
        assert_equal("line", st_sg.type)
        assert_equal(true, st_sg.markers)
        assert_equal(true, st_sg.high)
        assert_equal("FF0000FF", st_sg.color_series)
        assert_equal(200, st_sg.max)
        assert_equal(1.5, st_sg.line_weight)
        assert_equal(1, stream_sheet.sparkline_groups.first.sparklines.size)
        assert_equal("D2", stream_sheet.sparkline_groups.first.sparklines.first.sqref)

        assert_equal(1, stream_sheet.charts.size)
        st_ch = stream_sheet.charts.first
        assert_instance_of(Xlsxrb::Elements::Chart, st_ch)
        assert_equal("barChart", st_ch.type)
        assert_equal("Quarterly Sales", st_ch.title)

        # 3. Test StreamSheet#load materialization preserving all Phase 3 properties
        mat_ws = stream_sheet.load
        assert_equal("A1:C4", mat_ws.print_area)
        assert_not_nil(mat_ws.print_titles)
        assert_equal("1:1", mat_ws.print_titles.rows)
        assert_equal("A:A", mat_ws.print_titles.columns)
        assert_equal(1, mat_ws.sparkline_groups.size)
        assert_equal(1, mat_ws.charts.size)
      end
      assert_equal(1, read_sheets.size)
    ensure
      tmp.close!
    end
  end
end
