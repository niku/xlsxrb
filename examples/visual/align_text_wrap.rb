# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "align_text_wrap.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("wrap", &:wrap_text)
  wb.sheet("Text Wrap") do |s|
    s.print_options(:grid_lines, true)
    s.column(0, width: 15)
    s.row(["This is a long sentence that wraps inside the cell."], styles: { 0 => "wrap" })
  end
end
