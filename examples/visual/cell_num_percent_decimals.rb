# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cell_num_percent_decimals.xlsx"
Xlsxrb.generate(output_path) do |wb|
  wb.style("pct2") { |s| s.num_fmt("0.00%") }
  wb.sheet("Percents") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(%w[Format Value])
    s.row(["Percent with 2 Decimals", 0.12345], styles: { 1 => "pct2" })
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
