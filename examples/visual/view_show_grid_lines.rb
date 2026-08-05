# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "view_show_grid_lines.xlsx"
Xlsxrb.generate(output_path) do |wb|
  wb.sheet("Hide Grid Lines") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.sheet_view(:show_grid_lines, false)
    s.row(["No Grid Lines displayed"])
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
