# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "col_grouping.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.style("border") { |style| style.border_all(style: "thin", color: "FF000000") }
  w.sheet("Col Grouping") do |s|
    s.set_sheet_property(:fit_to_page, true)
    s.page_setup(fit_to_width: 1, fit_to_height: 1)
    s.column(0, width: 25, outline_level: 0)
    s.column(1, width: 25, outline_level: 1)
    s.column(2, width: 25, outline_level: 1)
    s.row(["Col A", "Col B (Grouped)", "Col C (Grouped)"], styles: %w[border border border])
  end
end

# 2. Read the generated sheet and print column dimensions
puts "=== Read Validation ==="
workbook = Xlsxrb.read(output_path)
sheet = workbook.sheets.first
sheet.columns.each do |col|
  puts "Column #{col.index}: width=#{col.width}, hidden=#{col.hidden}, outline_level=#{col.outline_level}"
end
