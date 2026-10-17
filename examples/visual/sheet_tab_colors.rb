# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "sheet_tab_colors.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.sheet("Red Tab") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.sheet_properties(:tab_color, :red)
    s.row(["Red tab sheet"])
  end
  wb.sheet("Green Tab") do |s|
    s.sheet_properties(:tab_color, :green)
    s.row(["Green tab sheet"])
  end
end
