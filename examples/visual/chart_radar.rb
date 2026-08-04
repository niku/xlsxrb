# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "chart_radar.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.sheet("Data") do |s|
    s.set_sheet_property(:fit_to_page, true)
    s.set_page_setup(fit_to_width: 1, fit_to_height: 1)
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(%w[Stat Value])
    s.row(["Atk", 80])
    s.row(["Def", 60])
    s.chart(
      type: :radar,
      title: "Stats",
      from_col: 3, from_row: 0, to_col: 8, to_row: 12,
      series: [{ cat_ref: "Data!$A$2:$A$3", val_ref: "Data!$B$2:$B$3", line_color: "4F81BD", line_width: 2.0 }]
    )
  end
end

# 2. Read the generated sheet and print the chart count
puts "=== Read Validation ==="
workbook = Xlsxrb.read(output_path)
sheet = workbook.sheets.first
puts "Sheet '#{sheet.name}' has #{sheet.charts.size} chart(s)"
