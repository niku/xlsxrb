# frozen_string_literal: true

require "xlsxrb"

output_path = ARGV[0] || "merge_freeze.xlsx"

Xlsxrb.generate(output_path) do |wb|
  wb.style("title") { |style| style.border_all(style: "thin", color: "FF000000").align_horizontal("center") }
  wb.style("border") { |style| style.border_all(style: "thin", color: "FF000000") }
  wb.sheet("Merge & Freeze") do |s|
    s.sheet_properties(:fit_to_page, true)
    s.page_setup(fit_to_width: 1, fit_to_height: 1)
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.column(2, width: 25)
    s.row(["Merged Title Row", nil, nil], styles: %w[title title title])
    s.row(["Header A", "Header B", "Header C"], styles: %w[border border border])
    s.row(["Row 1 Col A", "Row 1 Col B", "Row 1 Col C"], styles: %w[border border border])
    s.row(["Row 2 Col A", "Row 2 Col B", "Row 2 Col C"], styles: %w[border border border])

    s.merge(row: 0, col_start: 0, col_end: 2)
    s.freeze_pane(row: 2, col: 0)
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
