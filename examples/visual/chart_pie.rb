# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "chart_pie.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.sheet("Data") do |s|
    s.set_sheet_property(:fit_to_page, true)
    s.set_page_setup(fit_to_width: 1, fit_to_height: 1)
    s.set_column(0, width: 25)
    s.set_column(1, width: 25)
    s.add_row(%w[Label Percent])
    s.add_row(["Yes", 70])
    s.add_row(["No", 30])
    s.add_chart(
      type: :pie,
      title: "Responses",
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
workbook = Xlsxrb.read(output_path)
sheet = workbook.sheets.first
puts "Sheet '#{sheet.name}' has #{sheet.charts.size} chart(s)"
