# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "chart_bar_stacked.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.sheet("Data") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(["Year", "Sales A", "Sales B"])
    s.row([2024, 100, 150])
    s.row([2025, 120, 180])
    s.row([2026, 140, 210])
    s.set_sheet_property(:fit_to_page, true)
    s.page_setup(fit_to_width: 1, fit_to_height: 1)
    s.chart(
      type: :bar,
      grouping: :stacked,
      title: "Stacked Bar Chart",
      from_col: 0, from_row: 5, to_col: 6, to_row: 17,
      series: [
        { cat_ref: "'Data'!$A$2:$A$4", val_ref: "'Data'!$B$2:$B$4", name: "'Data'!$B$1", fill_color: "4F81BD" },
        { cat_ref: "'Data'!$A$2:$A$4", val_ref: "'Data'!$C$2:$C$4", name: "'Data'!$C$1", fill_color: "C0504D" }
      ]
    )
  end
end

# 2. Read the generated sheet and print the chart count
puts "=== Read Validation ==="
workbook = Xlsxrb.read(output_path)
sheet = workbook.sheets.first
puts "Sheet '#{sheet.name}' has #{sheet.charts.size} chart(s)"
