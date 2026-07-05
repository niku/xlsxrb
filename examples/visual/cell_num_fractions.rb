# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cell_num_fractions.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.add_style("frac") { |s| s.num_fmt("# ?/?") }
  w.add_sheet("Fractions") do |s|
    s.set_column(0, width: 25)
    s.set_column(1, width: 25)
    s.add_row(%w[Format Value])
    s.add_row(["Half", 0.5], styles: { 1 => "frac" })
    s.add_row(["Third", 0.3333], styles: { 1 => "frac" })
    s.add_row(["Quarter", 0.75], styles: { 1 => "frac" })
  end
end

# 2. Read the generated sheet and print parsed cell numbers and format codes
puts "=== Read Validation ==="
workbook = Xlsxrb.read(output_path)
sheet = workbook.sheets.first
sheet.rows.each do |row|
  row_cells = row.cells.map do |c|
    xf = workbook.styles[:cell_xfs][c.style_index] if c.style_index
    num_fmt = xf ? workbook.styles[:num_fmts][xf[:num_fmt_id]] : nil
    "#{c.ref}: #{c.value.inspect} (Format ID: #{xf&.[](:num_fmt_id)}, Code: #{num_fmt.inspect})"
  end
  puts "Row #{row.index}: #{row_cells.join(", ")}"
end
