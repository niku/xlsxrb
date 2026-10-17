# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "align_vertical.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("top") { |s| s.align_vertical("top") }
  wb.style("center") { |s| s.align_vertical("center") }
  wb.style("bottom") { |s| s.align_vertical("bottom") }
  wb.sheet("Vertical Alignment") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.print_options(:grid_lines, true)
    s.row(%w[Top Center Bottom], styles: { 0 => "top", 1 => "center", 2 => "bottom" }, height: 40)
  end
end
