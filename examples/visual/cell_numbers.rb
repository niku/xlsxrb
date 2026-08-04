# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cell_numbers.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.add_style("currency") { |s| s.num_fmt("$#,##0.00") }
  w.add_style("percent") { |s| s.num_fmt("0.0%") }
  w.sheet("Numbers") do |s|
    s.set_column(0, width: 25)
    s.set_column(1, width: 25)
    s.add_row(%w[Format Value])
    s.add_row(["Integer", 12_345])
    s.add_row(["Float", 123.456])
    s.add_row(["Currency", 1234.5], styles: { 1 => "currency" })
    s.add_row(["Percentage", 0.85], styles: { 1 => "percent" })
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
