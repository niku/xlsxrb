# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cf_data_bar.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("center") { |style| style.align_horizontal("center") }
  wb.sheet("Data Bars") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row([20], styles: "center")
    s.row([60], styles: "center")
    s.row([100], styles: "center")
    s.conditional_format("A1:A3", type: :dataBar, priority: 1, color: "FF0070C0")
  end
end
