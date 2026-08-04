# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "fill_solid_colors.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.add_style("red_fill") { |s| s.fill_color("FFFFC7CE") }
  w.add_style("green_fill") { |s| s.fill_color("FFC6EFCE") }
  w.sheet("Fills") do |s|
    s.set_column(0, width: 25)
    s.set_column(1, width: 25)
    s.row(%w[Color Preview])
    s.row(["Red", "Red Fill"], styles: { 1 => "red_fill" })
    s.row(["Green", "Green Fill"], styles: { 1 => "green_fill" })
  end
end

# 2. Read the generated sheet and print cell fill properties
puts "=== Read Validation ==="
workbook = Xlsxrb.read(output_path)
sheet = workbook.sheets.first
row = sheet.rows.first
row.cells.each do |c|
  xf = workbook.styles[:cell_xfs][c.style_index] if c.style_index
  fill = xf ? workbook.styles[:fills][xf[:fill_id]] : nil
  puts "Cell #{c.ref} ('#{c.value}'): fill pattern = #{fill&.[](:pattern).inspect}, fg_color = #{fill&.[](:fg_color).inspect}"
end
