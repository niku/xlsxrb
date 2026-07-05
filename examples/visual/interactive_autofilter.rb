# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "interactive_autofilter.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.add_sheet("Filter") do |s|
    s.set_column(0, width: 25)
    s.set_column(1, width: 25)
    s.add_row(%w[Name Department])
    s.add_row(%w[Alice HR])
    s.add_row(%w[Bob Eng])
    s.set_auto_filter("A1:B3")
  end
  w.add_defined_name("_xlnm._FilterDatabase", "Filter!$A$1:$B$3", sheet: "Filter", hidden: true)
end

# 2. Read the generated sheet and print cell values
puts "=== Read Validation ==="
workbook = Xlsxrb.read(output_path)
sheet = workbook.sheets.first
sheet.rows.first(5).each do |row|
  row_cells = row.cells.map { |c| "#{c.ref}: #{c.value.inspect}" }
  puts "Row #{row.index}: #{row_cells.join(", ")}"
end
