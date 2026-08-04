# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "fill_patterns.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.add_style("dark_gray") { |s| s.fill(pattern: "darkGray", fg_color: "FFC0C0C0", bg_color: "FFFFFFFF") }
  w.add_style("grid_fill") { |s| s.fill(pattern: "darkGrid", fg_color: "FFC0C0C0", bg_color: "FFFFFFFF") }
  w.sheet("Patterns") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(%w[Pattern Preview])
    s.row(["Dark Gray", "Pattern Fill"], styles: { 1 => "dark_gray" })
    s.row(["Dark Grid", "Grid Fill"], styles: { 1 => "grid_fill" })
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
