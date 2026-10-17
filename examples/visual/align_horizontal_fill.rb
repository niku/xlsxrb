# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "align_horizontal_fill.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("fill_align") { |s| s.align_horizontal("fill") }
  wb.sheet("Alignment") do |s|
    s.column(0, width: 30)
    s.row(["X "], styles: { 0 => "fill_align" })
  end
end
