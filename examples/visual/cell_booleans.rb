# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cell_booleans.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.sheet("Booleans") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(["Label", "Boolean Value"])
    s.row(["Is Active", true])
    s.row(["Is Pending", false])
  end
end

# 2. Read the generated sheet and print parsed cell values and Ruby classes
puts "=== Read Validation ==="
workbook = Xlsxrb.read(output_path).load
sheet = workbook.sheets.first
sheet.rows.each do |row|
  row_cells = row.cells.map do |c|
    "#{c.ref}: #{c.value.inspect} (#{c.value.class})"
  end
  puts "Row #{row.index}: #{row_cells.join(", ")}"
end
