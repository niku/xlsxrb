# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "sparkline_line.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.sheet("Sparkline") do |s|
    s.set_sheet_property(:fit_to_page, true)
    s.set_page_setup(fit_to_width: 1, fit_to_height: 1)
    s.set_column(0, width: 25)
    s.set_column(1, width: 25)
    s.add_row([10, 20, 15, 30, nil])
    s.add_sparkline_group(
      type: :line,
      markers: true,
      color_series: "FF000000",
      color_markers: "FFFF0000",
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
