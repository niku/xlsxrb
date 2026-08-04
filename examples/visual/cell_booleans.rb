# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cell_booleans.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.sheet("Booleans") do |s|
    s.set_column(0, width: 25)
    s.set_column(1, width: 25)
    s.add_row(["Label", "Boolean Value"])
    s.add_row(["Is Active", true])
    s.add_row(["Is Pending", false])
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
