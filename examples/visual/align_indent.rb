# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "align_indent.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("indent_1") { |s| s.align_horizontal("left").indent(1) }
  wb.style("indent_3") { |s| s.align_horizontal("left").indent(3) }
  wb.sheet("Indent") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.print_options(:grid_lines, true)
    s.row(["No Indent"])
    s.row(["Indent 1"], styles: { 0 => "indent_1" })
    s.row(["Indent 3"], styles: { 0 => "indent_3" })
  end
end
