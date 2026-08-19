# frozen_string_literal: true

require "xlsxrb"

output_path = ARGV[0] || "table_styles.xlsx"

Xlsxrb.write(output_path) do |wb|
  wb.sheet("Table Example") do |sheet|
    sheet.column(0..3, width: 20)
    sheet.row(%w[ID Name Department Salary])
    sheet.row([101, "Alice Smith", "Engineering", 120_000])
    sheet.row([102, "Bob Jones", "Marketing", 95_000])
    sheet.row([103, "Carol White", "Sales", 110_000])
    sheet.row([104, "David Brown", "Engineering", 130_000])

    sheet.table("A1:D5", columns: %w[ID Name Department Salary], name: "EmployeeTable", style: "TableStyleMedium9", total_row: false)
  end
end

puts "=== Read Validation ==="
workbook = Xlsxrb.read(output_path).load
sheet = workbook.sheets.first
sheet.rows.each do |row|
  row_cells = row.cells.map { |c| "#{c.ref}: #{c.value.inspect}" }
  puts "Row #{row.index}: #{row_cells.join(", ")}"
end
