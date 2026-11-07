# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cf_expression_formula.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("center") { |style| style.align_horizontal("center") }
  wb.sheet("CF Expression") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(["Values"])
    s.row([10], cell_styles: "center")
    s.row([20], cell_styles: "center")
    s.row([30], cell_styles: "center")
    s.row([100], cell_styles: "center") # Average is 40. 100 is above average.
    s.conditional_format("A2:A5", type: "expression", formula: "A2>AVERAGE($A$2:$A$5)", fill_color: "FFFFC7CE")
  end
end
