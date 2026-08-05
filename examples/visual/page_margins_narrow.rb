# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "page_margins_narrow.xlsx"
Xlsxrb.generate(output_path) do |wb|
  wb.sheet("Narrow Margins") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.page_margins(top: 0.25, bottom: 0.25, left: 0.25, right: 0.25, header: 0.1, footer: 0.1)
    s.row(["Narrow Margins sheet"])
  end
end

# 2. Read the generated sheet and print page setup properties
puts "=== Read Validation ==="
reader = Xlsxrb::Ooxml::Reader.new(output_path)
workbook = Xlsxrb.read(output_path)
sheet = workbook.sheets.first
margins = reader.page_margins(sheet: sheet.name)
setup = reader.page_setup(sheet: sheet.name)
opts = reader.print_options(sheet: sheet.name)
puts "Page Margins: #{margins.inspect}"
puts "Page Setup: #{setup.inspect}"
puts "Print Options: #{opts.inspect}"
