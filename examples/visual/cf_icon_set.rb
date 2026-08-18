# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cf_icon_set.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("center") { |style| style.align_horizontal("center") }
  wb.sheet("Icons") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row([25], styles: "center")
    s.row([50], styles: "center")
    s.row([75], styles: "center")
    s.conditional_format("A1:A3", type: :iconSet, icon_style: "3Arrows", priority: 1)
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
