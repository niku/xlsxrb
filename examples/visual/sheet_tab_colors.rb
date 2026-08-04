# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "sheet_tab_colors.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.sheet("Red Tab") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.sheet_properties(:tab_color, "FFFF0000")
    s.row(["Red tab sheet"])
  end
  w.sheet("Green Tab") do |s|
    s.sheet_properties(:tab_color, "FF00FF00")
    s.row(["Green tab sheet"])
  end
end

# 2. Read the generated sheet and print sheet properties
puts "=== Read Validation ==="
reader = Xlsxrb::Ooxml::Reader.new(output_path)
workbook = Xlsxrb.read(output_path)
workbook.sheet_names.each do |s_name|
  props = reader.sheet_properties(sheet: s_name)
  puts "Sheet: #{s_name}, tab color: #{props[:tab_color].inspect}"
end
