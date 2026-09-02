# frozen_string_literal: true

require "xlsxrb"

output_path = ARGV[0] || "drawing_shapes.xlsx"

Xlsxrb.write(output_path) do |wb|
  wb.sheet("Shapes") do |sheet|
    sheet.column(0..4, width: 12)
    sheet.row(["Diagram with shapes and annotations"])
    sheet.row([])
    sheet.shape(preset: "rect", text: "Start", from_col: 0, from_row: 2, to_col: 1, to_row: 4)
    sheet.shape(preset: "rightArrow", text: "Next", from_col: 1, from_row: 3, to_col: 2, to_row: 4)
    sheet.shape(preset: "roundRect", text: "Processing", from_col: 2, from_row: 2, to_col: 4, to_row: 4)
  end
end

puts "=== Read Validation ==="
workbook = Xlsxrb.read(output_path).load
sheet = workbook.sheets.first
sheet.rows.each do |row|
  row_cells = row.cells.map { |c| "#{c.ref}: #{c.value.inspect}" }
  puts "Row #{row.index}: #{row_cells.join(", ")}"
end
