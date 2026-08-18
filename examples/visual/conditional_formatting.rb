# frozen_string_literal: true

require "xlsxrb"

output_path = ARGV[0] || "conditional_formatting.xlsx"

Xlsxrb.write(output_path) do |wb|
  wb.style("center") { |style| style.align_horizontal("center") }
  wb.sheet("Scores") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row([90, 45, 72, 88], styles: %w[center center center center])

    s.conditional_format("A1:D1",
                         type: :cell_is, operator: :greaterThan,
                         formula: "80", priority: 1,
                         fill_color: "FFFFC7CE")
  end
end

# 2. Read the generated sheet and print cell values
puts "=== Read Validation ==="
workbook = Xlsxrb.read(output_path).load
sheet = workbook.sheets.first
sheet.rows.first(5).each do |row|
  row_cells = row.cells.map { |c| "#{c.ref}: #{c.value.inspect}" }
  puts "Row #{row.index}: #{row_cells.join(", ")}"
end
