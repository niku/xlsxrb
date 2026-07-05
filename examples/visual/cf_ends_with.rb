# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cf_ends_with.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.add_sheet("CF Ends") do |s|
    s.set_column(0, width: 25)
    s.set_column(1, width: 25)
    s.add_row(["Code"])
    s.add_row(["100-Z"])
    s.add_row(["200-Y"])
    s.add_row(["300-Z"])
    s.add_conditional_format("A2:A4", type: "endsWith", operator: "endsWith", text: "Z", formula: 'RIGHT(A2,1)="Z"', fill_color: "FFFFC7CE")
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
