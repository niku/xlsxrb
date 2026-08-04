# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "page_margins_wide.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.sheet("Wide Margins") do |s|
    s.set_column(0, width: 25)
    s.set_column(1, width: 25)
    s.set_page_margins(top: 1.0, bottom: 1.0, left: 1.0, right: 1.0, header: 0.5, footer: 0.5)
    s.row(["Wide Margins sheet"])
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
