# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "align_horizontal.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("left") { |s| s.align_horizontal("left") }
  wb.style("center") { |s| s.align_horizontal("center") }
  wb.style("right") { |s| s.align_horizontal("right") }
  wb.sheet("Alignment") do |s|
    s.print_options(:grid_lines, true)
    s.column(0, width: 20)
    s.column(1, width: 20)
    s.column(2, width: 20)
    s.row(%w[Left Center Right], styles: { 0 => "left", 1 => "center", 2 => "right" })
  end
end
