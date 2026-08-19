# frozen_string_literal: true

require "xlsxrb"

output_path = ARGV[0] || "pivot_table.xlsx"

Xlsxrb.write(output_path) do |wb|
  wb.sheet("SalesData") do |sheet|
    sheet.column(0..3, width: 18)
    sheet.row(%w[Region Quarter Sales Rep])
    sheet.row(["East", "Q1", 1000, "Alice"])
    sheet.row(["West", "Q1", 1500, "Bob"])
    sheet.row(["East", "Q2", 1200, "Alice"])
    sheet.row(["West", "Q2", 1800, "Bob"])
    sheet.row(["North", "Q1", 800, "Charlie"])
    sheet.row(["North", "Q2", 950, "Charlie"])

    sheet.pivot_table(
      "SalesData!A1:D7",
      row_fields: ["Region"],
      data_fields: [{ name: "Sales", subtotal: "sum" }],
      col_fields: ["Quarter"],
      dest_ref: "F1",
      name: "RegionalSalesSummary"
    )
  end
end

puts "=== Read Validation ==="
workbook = Xlsxrb.read(output_path).load
sheet = workbook.sheets.first
sheet.rows.each do |row|
  row_cells = row.cells.map { |c| "#{c.ref}: #{c.value.inspect}" }
  puts "Row #{row.index}: #{row_cells.join(", ")}"
end
