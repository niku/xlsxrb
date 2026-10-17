# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "col_widths.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("border") { |style| style.border_all(style: "thin", color: "FF000000") }
  wb.sheet("Widths") do |s|
    s.column(0, width: 30)
    s.column(1, width: 10)
    s.row(["Wide Column A", "Narrow B"], styles: %w[border border])
  end
end
