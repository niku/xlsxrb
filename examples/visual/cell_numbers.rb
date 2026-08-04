# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cell_numbers.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.style("currency") { |s| s.num_fmt("$#,##0.00") }
  w.style("percent") { |s| s.num_fmt("0.0%") }
  w.sheet("Numbers") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(%w[Format Value])
    s.row(["Integer", 12_345])
    s.row(["Float", 123.456])
    s.row(["Currency", 1234.5], styles: { 1 => "currency" })
    s.row(["Percentage", 0.85], styles: { 1 => "percent" })
  end
end

# 2. Read the generated sheet and print parsed cell values and Ruby classes
puts "=== Read Validation ==="
workbook = Xlsxrb.read(output_path)
sheet = workbook.sheets.first
sheet.rows.each do |row|
  row_cells = row.cells.map do |c|
    "#{c.ref}: #{c.value.inspect} (#{c.value.class})"
  end
  puts "Row #{row.index}: #{row_cells.join(", ")}"
end
