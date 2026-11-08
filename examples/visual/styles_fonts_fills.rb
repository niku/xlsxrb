# frozen_string_literal: true

require "xlsxrb"

output_path = ARGV[0] || "styles_fonts_fills.xlsx"

Xlsxrb.write(output_path) do |wb|
  wb.style("header") do |style|
    style.bold = true
    style.font_size = 14
    style.font_color = "FFFFFFFF"
    style.fill_color = "FF4F81BD"
  end

  wb.style("highlight") do |style|
    style.italic = true
    style.font_color = "FFC00000"
    style.fill_color = "FFFFFF00"
  end

  wb.sheet("Styles") do |sheet|
    sheet.column(0, width: 25)
    sheet.column(1, width: 25)
    sheet.row(["Header 1", "Header 2"], cell_styles: "header")
    sheet.row(["Normal Text", "Highlighted Text"], cell_styles: [nil, "highlight"])
  end
end
