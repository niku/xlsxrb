# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "page_grid_lines_print.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.sheet("Print Grid Lines") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.print_options(:grid_lines, true)
    s.row(["Grid Lines Printed sheet"])
  end
end
