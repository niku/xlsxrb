# frozen_string_literal: true

require "xlsxrb"

output_path = ARGV[0] || "pivot_table.xlsx"

Xlsxrb.write(output_path) do |wb|
  wb.sheet("SalesData") do |sheet|
    sheet.column(0..3, width: 14)
    sheet.row(%w[Region Quarter Sales Rep])
    sheet.row(["East", "Q1", 1000, "Alice"])
    sheet.row(["West", "Q1", 1500, "Bob"])
    sheet.row(["East", "Q2", 1200, "Alice"])
    sheet.row(["West", "Q2", 1800, "Bob"])
    sheet.row(["North", "Q1", 800, "Charlie"])
    sheet.row(["North", "Q2", 950, "Charlie"])

    sheet.row([])
    sheet.row([])
    # Pre-populate summary cells in the pivot table target region (A10:D15) for headless renderers
    sheet.row(["Sum of Sales", "Quarter", nil, nil])
    sheet.row(["Region", "Q1", "Q2", "Grand Total"])
    sheet.row(["East", 1000, 1200, 2200])
    sheet.row(["North", 800, 950, 1750])
    sheet.row(["West", 1500, 1800, 3300])
    sheet.row(["Grand Total", 3300, 3950, 7250])

    sheet.pivot_table(
      "SalesData!A1:D7",
      row_fields: ["Region"],
      data_fields: [{ name: "Sales", subtotal: "sum" }],
      col_fields: ["Quarter"],
      dest_ref: "A10",
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
