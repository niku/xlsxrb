# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cf_cell_greater_equal.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("center") { |style| style.align_horizontal("center") }
  wb.sheet("CF Greater Equal") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(["Values"])
    s.row([10], styles: "center")
    s.row([50], styles: "center")
    s.row([100], styles: "center")
    s.conditional_format("A2:A4", type: "cellIs", operator: "greaterThanOrEqual", formula: "50", fill_color: "FFFFC7CE")
  end
end
