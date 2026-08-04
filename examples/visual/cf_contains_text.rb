# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cf_contains_text.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.sheet("CF Contains") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(["Status"])
    s.row(["Error"])
    s.row(["Success"])
    s.row(["Pending"])
    s.add_conditional_format("A2:A4", type: "containsText", operator: "containsText", text: "Error", formula: 'NOT(ISERROR(SEARCH("Error",A2)))', fill_color: "FFFF0000")
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
