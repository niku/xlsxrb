# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "row_grouping.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("parent") { |style| style.border_all(style: "thin", color: "FF000000").bold }
  wb.style("child") { |style| style.border_all(style: "thin", color: "FF000000").align_horizontal("left").indent(2) }
  wb.sheet("Row Grouping") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(["Parent Row 1", ""], styles: %w[parent parent])
    s.row(["Child Row 1.1", ""], outline_level: 1, styles: %w[child child])
    s.row(["Child Row 1.2", ""], outline_level: 1, styles: %w[child child])
  end
end

# 2. Read the generated sheet and print row attributes
puts "=== Read Validation ==="
workbook = Xlsxrb.read(output_path).load
sheet = workbook.sheets.first
sheet.rows.each do |row|
  puts "Row #{row.index}: height=#{row.height}, hidden=#{row.hidden}, outline_level=#{row.outline_level}"
end
