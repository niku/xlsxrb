# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "fill_solid_colors.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("red_fill") { |s| s.fill_color("FFFFC7CE") }
  wb.style("green_fill") { |s| s.fill_color("FFC6EFCE") }
  wb.sheet("Fills") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(%w[Color Preview])
    s.row(["Red", "Red Fill"], styles: { 1 => "red_fill" })
    s.row(["Green", "Green Fill"], styles: { 1 => "green_fill" })
  end
end
