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
      b.protect_workbook(lock_structure: true)
      b.defined_name("DataRange", "Sheet1!$A$1:$B$2")
      b.sheet("Sheet1") do |s|
        s.freeze_pane(row: 1, col: 1)
        s.sheet_view(:show_grid_lines, false)
        s.sheet_view(:tab_selected, true)
        s.sheet_view(:zoom_scale, 120)
        s.sheet_properties(:tab_color, "FFFF0000")
        s.sheet_format(default_row_height: 20)
        s.page_setup(orientation: "landscape", paper_size: 9)
        s.page_margins(left: 0.7, right: 0.7, top: 0.75, bottom: 0.75)
        s.header_footer(odd_header: "&CTest Header", odd_footer: "&RPage &P")
        s.print_options(:grid_lines, true)
        s.protect_sheet(sheet: true)
        s.merge("C1:D2")
        s.auto_filter("A1:D1")
        s.conditional_format("A1:A5", type: :cellIs, operator: :greaterThan, formula1: "10")
        s.validate_data("B1:B5", type: :whole, operator: :between, formula1: "1", formula2: "100")
        s.page_break_row(5)
        s.page_break_col(3)
        s.hyperlink("A2", "https://example.com", tooltip: "Example")
        s.comment("B2", "Audited cell", author: "QA")
        s.image(MINIMAL_PNG, ext: "png", from_col: 4, from_row: 4, to_col: 6, to_row: 6)
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
        assert_equal "Author One", roundtrip_wb.core_properties[:creator]
        assert_equal "Comprehensive Round-Trip", roundtrip_wb.core_properties[:title]
        assert_equal "Engineering Org", roundtrip_wb.app_properties[:company]
        assert_equal [{ name: "Classification", value: "Internal", type: :string }], roundtrip_wb.custom_properties
        assert_equal({ lock_structure: true }, roundtrip_wb.workbook_protection)
        assert_equal :hidden, roundtrip_wb.sheet(1).state
        assert_equal([{ name: "DataRange", hidden: false, value: "Sheet1!$A$1:$B$2" }], roundtrip_wb.defined_names)

        s = roundtrip_wb.sheet(0)
        assert_equal({ row: 1, col: 1, state: :frozen, active_pane: "bottomRight" }, s.freeze_pane)
        assert_equal({ show_grid_lines: false, tab_selected: true, zoom_scale: 120 }, s.sheet_view)
        assert_equal({ orientation: "landscape", paper_size: 9 }, s.page_setup)
        assert_equal(0.7, s.page_margins[:left])
        assert_equal(0.7, s.page_margins[:right])
        assert_equal(0.75, s.page_margins[:top])
        assert_equal(0.75, s.page_margins[:bottom])
        assert_equal({ odd_header: "&CTest Header", odd_footer: "&RPage &P" }, s.header_footer)
        assert_equal({ grid_lines: true }, s.print_options)
        assert_equal({ sheet: true }, s.sheet_protection)
        assert_equal ["C1:D2"], s.merged_cells
        assert_equal "A1:D1", s.auto_filter
        assert_equal [{ id: 5, max: 16_383, man: true }], s.row_breaks
        assert_equal [{ id: 3, max: 1_048_575, man: true }], s.col_breaks
        assert_equal({ "A2" => { url: "https://example.com", tooltip: "Example" } }, s.hyperlinks)
        assert_equal [{ ref: "B2", text: "Audited cell", author: "QA" }], s.comments
        assert_equal 1, s.images.size
        assert_equal 1, s.tables.size
        assert_equal "RoundTripTable", s.tables.first[:name]
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
end
