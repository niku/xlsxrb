# frozen_string_literal: true

require "xlsxrb"

output_path = ARGV[0] || "conditional_formatting.xlsx"

Xlsxrb.generate(output_path) do |w|
  w.add_style("center") { |style| style.align_horizontal("center") }
  w.add_sheet("Scores") do |s|
    s.set_column(0, width: 25)
    s.set_column(1, width: 25)
    s.add_row([90, 45, 72, 88], styles: %w[center center center center])

    s.add_conditional_format("A1:D1",
                             type: :cell_is, operator: :greaterThan,
                             formula: "80", priority: 1,
                             fill_color: "FFFFC7CE")
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
