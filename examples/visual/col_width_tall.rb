# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "col_width_tall.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.add_style("border") { |style| style.border_all(style: "thin", color: "FF000000") }
  w.sheet("Col Width") do |s|
    s.set_column(0, width: 50)
    s.set_column(1, width: 10)
    s.row(["Very Wide Column (Width 50)", "Normal (10)"], styles: %w[border border])
  end
end

# 2. Read the generated sheet and print column dimensions
puts "=== Read Validation ==="
workbook = Xlsxrb.read(output_path)
sheet = workbook.sheets.first
sheet.columns.each do |col|
  puts "Column #{col.index}: width=#{col.width}, hidden=#{col.hidden}, outline_level=#{col.outline_level}"
end
