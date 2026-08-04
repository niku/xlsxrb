# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "align_horizontal_justify.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.style("justify_align") do |s|
    s.align_horizontal("justify")
    s.wrap_text(true)
  end
  w.sheet("Alignment") do |s|
    s.column(0, width: 25)
    s.row(["Justified alignment wraps and distributes text evenly."], styles: { 0 => "justify_align" })
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
