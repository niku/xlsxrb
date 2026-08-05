# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "chart_scatter.xlsx"
Xlsxrb.generate(output_path) do |wb|
  wb.sheet("Data") do |s|
    s.sheet_properties(:fit_to_page, true)
    s.page_setup(fit_to_width: 1, fit_to_height: 1)
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(%w[X Y])
    s.row([1, 10])
    s.row([2, 15])
    s.chart(
      type: :scatter,
      title: "Scatter Plot",
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
