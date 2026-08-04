# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "align_text_rotation.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.add_style("rot_45") { |s| s.text_rotation(45) }
  w.add_style("rot_90") { |s| s.text_rotation(90) }
  w.sheet("Rotation") do |s|
    s.set_column(0, width: 25)
    s.set_column(1, width: 25)
    s.set_print_option(:grid_lines, true)
    s.add_row(["Rotated 45", "Rotated 90"], styles: { 0 => "rot_45", 1 => "rot_90" }, height: 50)
  end
end

# 2. Read the generated sheet and print the parsed alignments
puts "=== Read Validation ==="
workbook = Xlsxrb.read(output_path)
sheet = workbook.sheets.first
row = sheet.rows.first
row.cells.each do |c|
  xf = workbook.styles[:cell_xfs][c.style_index] if c.style_index
  align_h = xf&.dig(:alignment, :horizontal)
  align_v = xf&.dig(:alignment, :vertical)
  wrap = xf&.dig(:alignment, :wrap_text)
  indent = xf&.dig(:alignment, :indent)
  rot = xf&.dig(:alignment, :text_rotation)
  shrink = xf&.dig(:alignment, :shrink_to_fit)
  puts "Cell #{c.ref} ('#{c.value}'): align_h=#{align_h.inspect}, align_v=#{align_v.inspect}, wrap=#{wrap.inspect}, indent=#{indent.inspect}, rotation=#{rot.inspect}, shrink=#{shrink.inspect}"
end
