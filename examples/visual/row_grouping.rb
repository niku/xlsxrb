# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "row_grouping.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("parent") do |style|
    style.border_all(style: "thin", color: "FF000000")
    style.bold(true)
  end
  wb.style("child") do |style|
    style.border_all(style: "thin", color: "FF000000")
    style.align_horizontal("left")
    style.indent(2)
  end
  wb.sheet("Row Grouping") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(["Parent Row 1", ""], cell_styles: %w[parent parent])
    s.row(["Child Row 1.1", ""], outline_level: 1, cell_styles: %w[child child])
    s.row(["Child Row 1.2", ""], outline_level: 1, cell_styles: %w[child child])
  end
end
