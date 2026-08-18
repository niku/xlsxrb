# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "page_grid_lines_print.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.sheet("Print Grid Lines") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.print_options(:grid_lines, true)
    s.row(["Grid Lines Printed sheet"])
  end
end

# 2. Read the generated sheet and print page setup properties
puts "=== Read Validation ==="
reader = Xlsxrb::Ooxml::Reader.new(output_path)
workbook = Xlsxrb.read(output_path).load
sheet = workbook.sheets.first
margins = reader.page_margins(sheet: sheet.name)
setup = reader.page_setup(sheet: sheet.name)
opts = reader.print_options(sheet: sheet.name)
puts "Page Margins: #{margins.inspect}"
puts "Page Setup: #{setup.inspect}"
puts "Print Options: #{opts.inspect}"
