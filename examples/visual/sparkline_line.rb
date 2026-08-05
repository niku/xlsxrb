# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "sparkline_line.xlsx"
Xlsxrb.generate(output_path) do |wb|
  wb.sheet("Sparkline") do |s|
    s.sheet_properties(:fit_to_page, true)
    s.page_setup(fit_to_width: 1, fit_to_height: 1)
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row([10, 20, 15, 30, nil])
    s.sparkline_group(
      type: :line,
      markers: true,
      color_series: "FF000000",
      color_markers: :red,
      sparklines: [{ location_ref: "E1", data_ref: "Sparkline!A1:D1" }]
    )
  end
end

# 2. Read the generated sheet and print cell values
puts "=== Read Validation ==="
workbook = Xlsxrb.read(output_path)
sheet = workbook.sheets.first
sheet.rows.first(3).each do |row|
  row_cells = row.cells.map { |c| "#{c.ref}: #{c.value.inspect}" }
  puts "Row #{row.index}: #{row_cells.join(", ")}"
end
