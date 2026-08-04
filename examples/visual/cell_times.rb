# frozen_string_literal: true

require "xlsxrb"
require "time"
output_path = ARGV[0] || "cell_times.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.add_style("time_fmt") { |s| s.num_fmt("hh:mm:ss") }
  w.sheet("Times") do |s|
    s.set_column(0, width: 25)
    s.set_column(1, width: 25)
    s.add_row(["Format", "Time Value"])
    s.add_row(["DateTime", Time.new(2026, 7, 1, 12, 34, 56)])
    s.add_row(["Time Only", Time.new(2026, 7, 1, 9, 15, 0)], styles: { 1 => "time_fmt" })
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
