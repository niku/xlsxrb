# frozen_string_literal: true

require "xlsxrb"

output_path = ARGV[0] || "styles_fonts_fills.xlsx"

Xlsxrb.write(output_path) do |wb|
  wb.style("header") do |style|
    style.bold.size(14).font_color("FFFFFFFF").fill_color("FF4F81BD")
  end

  wb.style("highlight") do |style|
    style.italic.font_color("FFC00000").fill_color("FFFFFF00")
  end

  wb.sheet("Styles") do |sheet|
    sheet.column(0, width: 25)
    sheet.column(1, width: 25)
    sheet.row(["Header 1", "Header 2"], styles: "header")
    sheet.row(["Normal Text", "Highlighted Text"], styles: [nil, "highlight"])
  end
end
