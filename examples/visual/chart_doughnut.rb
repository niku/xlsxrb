# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "chart_doughnut.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.sheet("Data") do |s|
    s.sheet_properties(:fit_to_page, true)
    s.page_setup(fit_to_width: 1, fit_to_height: 1)
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(%w[Label Percent])
    s.row(["A", 40])
    s.row(["B", 60])
    s.chart(
      type: :doughnut,
      title: "Ratio",
      from_col: 3, from_row: 0, to_col: 8, to_row: 12,
      series: [{
        cat_ref: "Data!$A$2:$A$3",
        val_ref: "Data!$B$2:$B$3",
        data_points: [
          { idx: 0, fill_color: "4F81BD" },
          { idx: 1, fill_color: "C0504D" }
        ]
      }]
    )
  end
end

# 2. Read the generated sheet and print the chart count
puts "=== Read Validation ==="
workbook = Xlsxrb.read(output_path).load
sheet = workbook.sheets.first
puts "Sheet '#{sheet.name}' has #{sheet.charts.size} chart(s)"
