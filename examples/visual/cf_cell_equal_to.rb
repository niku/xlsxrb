# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cf_cell_equal_to.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("center") { |style| style.align_horizontal("center") }
  wb.sheet("CF Equal") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(["Values"])
    s.row([10], styles: "center")
    s.row([20], styles: "center")
    s.row([10], styles: "center")
    s.conditional_format("A2:A4", type: "cellIs", operator: "equal", formula: "10", fill_color: "FFFFC7CE")
  end
end
