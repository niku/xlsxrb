# frozen_string_literal: true

require "xlsxrb"
require "time"
output_path = ARGV[0] || "cell_times.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("time_fmt") { |s| s.num_fmt("hh:mm:ss") }
  wb.sheet("Times") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(["Format", "Time Value"])
    s.row(["DateTime", Time.new(2026, 7, 1, 12, 34, 56)])
    s.row(["Time Only", Time.new(2026, 7, 1, 9, 15, 0)], styles: { 1 => "time_fmt" })
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
