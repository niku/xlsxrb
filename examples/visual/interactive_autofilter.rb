# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "interactive_autofilter.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.sheet("Filter") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(%w[Name Department])
    s.row(%w[Alice HR])
    s.row(%w[Bob Eng])
    s.auto_filter("A1:B3")
  end
  w.defined_name("_xlnm._FilterDatabase", "Filter!$A$1:$B$3", sheet: "Filter", hidden: true)
end

# 2. Read the generated sheet and print cell values
puts "=== Read Validation ==="
workbook = Xlsxrb.read(output_path)
sheet = workbook.sheets.first
sheet.rows.first(5).each do |row|
  row_cells = row.cells.map { |c| "#{c.ref}: #{c.value.inspect}" }
  puts "Row #{row.index}: #{row_cells.join(", ")}"
end
