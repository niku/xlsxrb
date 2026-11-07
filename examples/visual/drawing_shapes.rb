# frozen_string_literal: true

require "xlsxrb"

output_path = ARGV[0] || "drawing_shapes.xlsx"

Xlsxrb.write(output_path) do |wb|
  wb.sheet("Shapes") do |sheet|
    sheet.column(0..4, width: 12)
    sheet.row(["Diagram with shapes and annotations"])
    sheet.row([])
    sheet.shape(preset: "rect", text: "Start", from_column: 0, from_row: 2, to_column: 1, to_row: 4)
    sheet.shape(preset: "rightArrow", text: "Next", from_column: 1, from_row: 3, to_column: 2, to_row: 4)
    sheet.shape(preset: "roundRect", text: "Processing", from_column: 2, from_row: 2, to_column: 4, to_row: 4)
  end
end
