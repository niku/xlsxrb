# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cf_color_scale.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.style("center") { |style| style.align_horizontal("center") }
  w.sheet("Colors") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row([10], styles: ["center"])
    s.row([50], styles: ["center"])
    s.row([90], styles: ["center"])
    s.conditional_format("A1:A3", type: :colorScale, priority: 1)
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
