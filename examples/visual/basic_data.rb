# frozen_string_literal: true

require "xlsxrb"
require "date"

output_path = ARGV[0] || "basic_data.xlsx"

Xlsxrb.generate(output_path) do |w|
  w.add_style("currency") { |style| style.number_format("$#,##0.00") }
  w.add_style("date") { |style| style.number_format("yyyy-mm-dd") }
  w.sheet("Basic Data") do
    w.set_sheet_property(:fit_to_page, true)
    w.set_page_setup(fit_to_width: 1, fit_to_height: 1)
    w.set_column(0, width: 25)
    w.set_column(1, width: 25)
    w.add_row(%w[Product Qty Price Date Active])
    w.add_row(["Gadget A", 10, 99.99, Date.new(2026, 1, 15), true], styles: { 2 => "currency", 3 => "date" })
    w.add_row(["Widget B", 5, 49.50, Date.new(2026, 2, 20), false], styles: { 2 => "currency", 3 => "date" })
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
