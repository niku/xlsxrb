# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "align_text_wrap.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.add_style("wrap", &:wrap_text)
  w.add_sheet("Text Wrap") do |s|
    s.set_print_option(:grid_lines, true)
    s.set_column(0, width: 15)
    s.add_row(["This is a long sentence that wraps inside the cell."], styles: { 0 => "wrap" })
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
