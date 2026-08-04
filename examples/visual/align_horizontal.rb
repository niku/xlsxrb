# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "align_horizontal.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.add_style("left") { |s| s.align_horizontal("left") }
  w.add_style("center") { |s| s.align_horizontal("center") }
  w.add_style("right") { |s| s.align_horizontal("right") }
  w.sheet("Alignment") do |s|
    s.set_print_option(:grid_lines, true)
    s.column(0, width: 20)
    s.column(1, width: 20)
    s.column(2, width: 20)
    s.row(%w[Left Center Right], styles: { 0 => "left", 1 => "center", 2 => "right" })
  end
end

# 2. Read the generated sheet and print the parsed alignments
puts "=== Read Alignment (Xlsxrb.read) ==="
workbook = Xlsxrb.read(output_path)
sheet = workbook.sheets.first
row = sheet.rows.first

row.cells.each do |c|
  xf = workbook.styles[:cell_xfs][c.style_index] if c.style_index
  align_h = xf&.dig(:alignment, :horizontal)
  puts "Cell #{c.ref} ('#{c.value}'): align_horizontal = #{align_h.inspect}"
end
