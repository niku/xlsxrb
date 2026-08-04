# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cell_formulas.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.sheet("Formulas") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(%w[Item Value])
    s.row(["A", 10])
    s.row(["B", 20])
    s.row(["SUM", Xlsxrb::Elements::Formula.new(expression: "SUM(B2:B3)", cached_value: 30)])
    s.row(["AVERAGE", Xlsxrb::Elements::Formula.new(expression: "AVERAGE(B2:B3)", cached_value: 15)])
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
