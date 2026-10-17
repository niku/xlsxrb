# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cf_color_scale.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("center") { |style| style.align_horizontal("center") }
  wb.sheet("Colors") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row([10], styles: "center")
    s.row([50], styles: "center")
    s.row([90], styles: "center")
    s.conditional_format("A1:A3", type: :colorScale, priority: 1)
  end
end
