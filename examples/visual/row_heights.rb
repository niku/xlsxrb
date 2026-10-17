# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "row_heights.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("border") { |style| style.border_all(style: "thin", color: "FF000000") }
  wb.sheet("Heights") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(["Normal Row", ""], styles: %w[border border])
    s.row(["Tall Row", ""], height: 40, styles: %w[border border])
  end
end
