# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "align_horizontal_justify.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("justify_align") do |s|
    s.align_horizontal("justify")
    s.wrap_text(true)
  end
  wb.sheet("Alignment") do |s|
    s.column(0, width: 25)
    s.row(["Justified alignment wraps and distributes text evenly."], styles: { 0 => "justify_align" })
  end
end
