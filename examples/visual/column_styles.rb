# frozen_string_literal: true

require "xlsxrb"
require "date"

output_path = ARGV[0] || "column_styles.xlsx"

Xlsxrb.write(output_path) do |wb|
  wb.style(:header, bold: true, fill_color: "D9E1F2", align_horizontal: :center)
  wb.style(:currency, number_format: "$#,##0.00", align_horizontal: :right)
  wb.style(:date, number_format: "yyyy-mm-dd", align_horizontal: :center)
  wb.style(:percent, number_format: "0.0%", align_horizontal: :right)

  wb.sheet("Column Styles") do |s|
    # Configure default styles per column so data rows do not need per-cell style mapping
    s.column(0, width: 20)
    s.column(1, width: 16, style: :date)
    s.column(2, width: 16, style: :currency)
    s.column(3, width: 14, style: :percent)

    s.row(%w[Item Date Price Margin], styles: :header)
    s.row(["Widget A", Date.new(2026, 4, 1), 1250.0, 0.25])
    s.row(["Gadget B", Date.new(2026, 5, 15), 340.5, 0.18])
    s.row(["Tool C", Date.new(2026, 6, 20), 89.99, 0.42])
  end
end
