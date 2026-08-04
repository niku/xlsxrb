# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "page_header_footer.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.sheet("Header Footer") do |s|
    s.set_column(0, width: 25)
    s.set_column(1, width: 25)
    s.set_header_footer(odd_header: "&LMy Company&RPage &P", odd_footer: "&CConfidential")
    s.row(["Header Footer sheet"])
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
