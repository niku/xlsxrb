# frozen_string_literal: true

require "xlsxrb"

output_path = ARGV[0] || "chart_bar.xlsx"

Xlsxrb.generate(output_path) do |w|
  w.sheet("Sales Data") do |s|
    s.set_sheet_property(:fit_to_page, true)
    s.set_page_setup(fit_to_width: 1, fit_to_height: 1)
    s.set_column(0, width: 25)
    s.set_column(1, width: 25)
    s.add_row(%w[Month Value])
    s.add_row(["Jan", 100])
    s.add_row(["Feb", 200])

    s.add_chart(
      type: :bar,
      title: "Monthly Sales",
      from_col: 3,
      from_row: 0,
      to_col: 8,
      to_row: 12,
      series: [
        { cat_ref: "'Sales Data'!$A$2:$A$3", val_ref: "'Sales Data'!$B$2:$B$3", fill_color: "4F81BD" }
      ]
    )
  end
end

# 2. Read the generated sheet and print the chart count
puts "=== Read Validation ==="
workbook = Xlsxrb.read(output_path)
sheet = workbook.sheets.first
puts "Sheet '#{sheet.name}' has #{sheet.charts.size} chart(s)"
