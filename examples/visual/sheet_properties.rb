# frozen_string_literal: true

require "xlsxrb"

output_path = ARGV[0] || "sheet_properties.xlsx"

# 1x1 subtle neutral pixel as background image
subtle_bg = [
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg=="
].pack("m")

Xlsxrb.write(output_path) do |wb|
  wb.tab_ratio(600)
  wb.active_sheet("Active Summary")

  wb.style(:header, font: { bold: true, color: "FFFFFF" }, fill_color: "1E3A8A", align_horizontal: :center)
  wb.style(:meta_hdr, font: { bold: true, color: "FFFFFF" }, fill_color: "047857", align_horizontal: :center)
  wb.style(:border_cell, border_all: { style: :thin, color: "CBD5E1" })

  # Sheet 1: Active Summary with tab color, fit to page, code name, and ignored errors
  wb.sheet("Active Summary", tab_color: { rgb: "FF2563EB" }, fit_to_page: true, code_name: "SheetSummary") do |s|
    s.column(0, width: 16)
    s.column(1, width: 28)
    s.column(2, width: 16)

    s.phonetic_properties(font_id: 0, type: :no_conversion, alignment: :left)
    s.ignore_error("A2:A4", number_stored_as_text: true)

    s.row(["Code", "Product Description", "Inventory"], styles: :header)
    s.row(["00101", "Compact Mechanical Keyboard", 42], styles: :border_cell)
    s.row(["00102", "Ergonomic Vertical Mouse", 78], styles: :border_cell)
    s.row(["00103", "Ultra-wide Monitor Stand", 15], styles: :border_cell)
  end

  # Sheet 2: Secondary sheet with green tab, sheet calculation pr, and background image
  wb.sheet("Calculations & Notes", tab_color: { rgb: "FF10B981" }) do |s|
    s.column(0, width: 22)
    s.column(1, width: 20)

    s.sheet_calc_pr(full_calc_on_load: true)
    s.background_image(subtle_bg)

    s.row(%w[Metric Calculation], styles: :meta_hdr)
    s.row(["Total Units", "=SUM('Active Summary'!C2:C4)"], styles: :border_cell)
    s.row(["Average Stock", "=AVERAGE('Active Summary'!C2:C4)"], styles: :border_cell)
  end

  # Sheet 3: Hidden configuration sheet
  wb.sheet("Internal Config", state: :hidden) do |s|
    s.column(0, width: 20)
    s.column(1, width: 20)
    s.row(["Config Key", "Value"])
    s.row(%w[environment production])
    s.row(%w[telemetry_enabled true])
  end
end
