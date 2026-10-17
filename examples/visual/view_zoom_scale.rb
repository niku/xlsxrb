# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "view_zoom_scale.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.sheet("Zoom 150") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.sheet_view(:zoom_scale, 150)
    s.row(["Zoom scale is set to 150%"])
  end
end
