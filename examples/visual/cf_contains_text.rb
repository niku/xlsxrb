# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cf_contains_text.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.add_sheet("CF Contains") do |s|
    s.set_column(0, width: 25)
    s.set_column(1, width: 25)
    s.add_row(["Status"])
    s.add_row(["Error"])
    s.add_row(["Success"])
    s.add_row(["Pending"])
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
