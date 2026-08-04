# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cf_cell_between.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.add_style("center") { |style| style.align_horizontal("center") }
  w.sheet("CF Between") do |s|
    s.set_column(0, width: 25)
    s.set_column(1, width: 25)
    s.add_row(["Values"])
    s.add_row([15], styles: ["center"])
    s.add_row([25], styles: ["center"])
    s.add_row([5], styles: ["center"])
    s.add_conditional_format("A2:A4", type: "cellIs", operator: "between", formulas: %w[10 20], fill_color: "FF00FF00")
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
