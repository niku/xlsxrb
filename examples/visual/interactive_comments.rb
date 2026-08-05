# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "interactive_comments.xlsx"
Xlsxrb.generate(output_path) do |wb|
  wb.sheet("Comments") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(["Item A", "Item B"])
    s.comment("A1", "This is an important comment.", author: "System")
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
