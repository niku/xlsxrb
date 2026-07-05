# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "align_vertical.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.add_style("top") { |s| s.align_vertical("top") }
  w.add_style("center") { |s| s.align_vertical("center") }
  w.add_style("bottom") { |s| s.align_vertical("bottom") }
  w.add_sheet("Vertical Alignment") do |s|
    s.set_column(0, width: 25)
    s.set_column(1, width: 25)
    s.set_print_option(:grid_lines, true)
    s.add_row(%w[Top Center Bottom], styles: { 0 => "top", 1 => "center", 2 => "bottom" }, height: 40)
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
