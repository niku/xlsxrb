# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "col_width_tall.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("border") { |style| style.border_all(style: "thin", color: "FF000000") }
  wb.sheet("Col Width") do |s|
    s.column(0, width: 50)
    s.column(1, width: 10)
    s.row(["Very Wide Column (Width 50)", "Normal (10)"], styles: %w[border border])
  end
end
