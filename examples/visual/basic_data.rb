# frozen_string_literal: true

require "xlsxrb"
require "date"

output_path = ARGV[0] || "basic_data.xlsx"

Xlsxrb.generate(output_path) do |wb|
  wb.style("currency") { |style| style.number_format("$#,##0.00") }
  wb.style("date") { |style| style.number_format("yyyy-mm-dd") }
  wb.sheet("Basic Data") do |sheet|
    sheet.sheet_properties(:fit_to_page, true)
    sheet.page_setup(fit_to_width: 1, fit_to_height: 1)
    sheet.column(0, width: 25)
    sheet.column(1, width: 25)
    sheet.row(%w[Product Qty Price Date Active])
    sheet.row(["Gadget A", 10, 99.99, Date.new(2026, 1, 15), true], styles: { 2 => "currency", 3 => "date" })
    sheet.row(["Widget B", 5, 49.50, Date.new(2026, 2, 20), false], styles: { 2 => "currency", 3 => "date" })
  end
end

# 2. Read the generated sheet and print parsed cell values and Ruby classes
puts "=== Read Validation ==="
workbook = Xlsxrb.read(output_path)
sheet = workbook.sheets.first
sheet.rows.each do |row|
  row_cells = row.cells.map do |c|
    "#{c.ref}: #{c.value.inspect} (#{c.value.class})"
  end
  puts "Row #{row.index}: #{row_cells.join(", ")}"
end
