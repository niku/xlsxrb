# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "row_heights.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.add_style("border") { |style| style.border_all(style: "thin", color: "FF000000") }
  w.sheet("Heights") do |s|
    s.set_column(0, width: 25)
    s.set_column(1, width: 25)
    s.add_row(["Normal Row", ""], styles: %w[border border])
    s.add_row(["Tall Row", ""], height: 40, styles: %w[border border])
  end
end

# 2. Read the generated sheet and print row attributes
puts "=== Read Validation ==="
workbook = Xlsxrb.read(output_path)
sheet = workbook.sheets.first
sheet.rows.each do |row|
  puts "Row #{row.index}: height=#{row.height}, hidden=#{row.hidden}, outline_level=#{row.outline_level}"
end
