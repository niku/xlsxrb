# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "row_height_tall.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("border") { |style| style.border_all(style: "thin", color: "FF000000") }
  wb.sheet("Row Height") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(["Very Tall Row (Height 60)", ""], height: 60, styles: %w[border border])
    s.row(["Normal Row", ""], styles: %w[border border])
  end
end

# 2. Read the generated sheet and print row attributes
puts "=== Read Validation ==="
workbook = Xlsxrb.read(output_path).load
sheet = workbook.sheets.first
sheet.rows.each do |row|
  puts "Row #{row.index}: height=#{row.height}, hidden=#{row.hidden}, outline_level=#{row.outline_level}"
end
