# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "view_zoom_scale.xlsx"
Xlsxrb.generate(output_path) do |wb|
  wb.sheet("Zoom 150") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.sheet_view(:zoom_scale, 150)
    s.row(["Zoom scale is set to 150%"])
  end
end

# 2. Read the generated sheet and print view configurations
puts "=== Read Validation ==="
reader = Xlsxrb::Ooxml::Reader.new(output_path)
workbook = Xlsxrb.read(output_path)
workbook.sheet_names.each do |s_name|
  view = reader.sheet_view(sheet: s_name)
  puts "Sheet '#{s_name}' views zoom scale: #{view[:zoom_scale]}%, show grid lines: #{view[:show_grid_lines]}"
end
