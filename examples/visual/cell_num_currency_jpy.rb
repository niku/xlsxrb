# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cell_num_currency_jpy.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.add_style("jpy") { |s| s.num_fmt("¥#,##0;[Red]¥-#,##0") }
  w.sheet("JPY Currency") do |s|
    s.set_column(0, width: 25)
    s.set_column(1, width: 25)
    s.add_row(%w[Format Value])
    s.add_row(["Positive Yen", 12_500], styles: { 1 => "jpy" })
    s.add_row(["Negative Yen", -8000], styles: { 1 => "jpy" })
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
