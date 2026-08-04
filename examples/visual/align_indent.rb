# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "align_indent.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.style("indent_1") { |s| s.align_horizontal("left").indent(1) }
  w.style("indent_3") { |s| s.align_horizontal("left").indent(3) }
  w.sheet("Indent") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.set_print_option(:grid_lines, true)
    s.row(["No Indent"])
    s.row(["Indent 1"], styles: { 0 => "indent_1" })
    s.row(["Indent 3"], styles: { 0 => "indent_3" })
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
