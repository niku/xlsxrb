# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "view_show_grid_lines.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.sheet("Hide Grid Lines") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.sheet_view(:show_grid_lines, false)
    s.row(["No Grid Lines displayed"])
  end
end
