# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "page_orientation_landscape.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.sheet("Landscape") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.page_setup(orientation: :landscape)
    s.row(["Landscape Orientation sheet"])
  end
end
