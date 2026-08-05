# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cf_begins_with.xlsx"
Xlsxrb.generate(output_path) do |wb|
  wb.sheet("CF Begins") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(["Code"])
    s.row(["A-100"])
    s.row(["B-200"])
    s.row(["A-300"])
    s.conditional_format("A2:A4", type: "beginsWith", operator: "beginsWith", text: "A", formula: 'LEFT(A2,1)="A"', fill_color: "FFFFC7CE")
  end
end

# 2. Read the generated sheet and print cell values
puts "=== Read Validation ==="
workbook = Xlsxrb.read(output_path)
sheet = workbook.sheets.first
sheet.rows.first(5).each do |row|
  row_cells = row.cells.map { |c| "#{c.ref}: #{c.value.inspect}" }
  puts "Row #{row.index}: #{row_cells.join(", ")}"
end
