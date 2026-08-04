# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cf_cell_less_than.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.style("center") { |style| style.align_horizontal("center") }
  w.sheet("CF Less") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(["Values"])
    s.row([25], styles: ["center"])
    s.row([75], styles: ["center"])
    s.row([10], styles: ["center"])
    s.add_conditional_format("A2:A4", type: "cellIs", operator: "lessThan", formula: "20", fill_color: "FFFFC7CE")
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
