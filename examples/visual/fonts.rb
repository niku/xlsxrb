# frozen_string_literal: true

require "xlsxrb"

output_path = ARGV[0] || "fonts.xlsx"

Xlsxrb.generate(output_path) do |wb|
  # Font Families (using nested hash options)
  wb.style("f_arial", font: { name: "Arial" })
  wb.style("f_times", font: { name: "Times New Roman" })
  wb.style("f_courier", font: { name: "Courier New" })
  wb.style("f_georgia", font: { name: "Georgia" })
  wb.style("f_tahoma", font: { name: "Tahoma" })

  # Font Sizes
  wb.style("sz_10", font: { size: 10 })
  wb.style("sz_16", font: { size: 16 })
  wb.style("sz_24", font: { size: 24 })

  # Font Colors
  wb.style("c_red", font: { color: "FFC00000" })
  wb.style("c_green", font: { color: "FF008000" })
  wb.style("c_blue", font: { color: :blue })

  # Font Styles
  wb.style("bold", font: { bold: true })
  wb.style("italic", font: { italic: true })
  wb.style("underline", font: { underline: true })
  wb.style("double_underline", font: { underline: "double" })
  wb.style("strike", font: { strike: true })

  # Vertical Alignment (Subscript / Superscript)
  wb.style("superscript", font: { vert_align: "superscript" })
  wb.style("subscript", font: { vert_align: "subscript" })

  wb.style("header", font: { bold: true, color: :white }, fill: { color: "FF4F81BD" })
  wb.style("bg_light", fill: { color: "FFDCE6F1" })

  wb.sheet("Fonts") do |s|
    s.column(0, width: 25)
    s.column(1, width: 35)

    s.row(["Font Feature", "Text Preview"], styles: "header")

    s.row(["Family: Arial", "Arial Text"], styles: %w[bg_light f_arial])
    s.row(["Family: Times New Roman", "Times New Roman"], styles: [nil, "f_times"])
    s.row(["Family: Courier New", "Courier New Text"], styles: %w[bg_light f_courier])
    s.row(["Family: Georgia", "Georgia Text"], styles: [nil, "f_georgia"])
    s.row(["Family: Tahoma", "Tahoma Text"], styles: %w[bg_light f_tahoma])

    s.row(["Size: 10pt", "10pt Font Size"], styles: [nil, "sz_10"])
    s.row(["Size: 16pt", "16pt Font Size"], styles: %w[bg_light sz_16])
    s.row(["Size: 24pt", "24pt Font Size"], styles: [nil, "sz_24"])

    s.row(["Color: Red", "Red Text"], styles: %w[bg_light c_red])
    s.row(["Color: Green", "Green Text"], styles: [nil, "c_green"])
    s.row(["Color: Blue", "Blue Text"], styles: %w[bg_light c_blue])

    s.row(["Style: Bold", "Bold Text"], styles: [nil, "bold"])
    s.row(["Style: Italic", "Italic Text"], styles: %w[bg_light italic])
    s.row(["Style: Underline", "Underline Text"], styles: [nil, "underline"])
    s.row(["Style: Double Underline", "Double Underline"], styles: %w[bg_light double_underline])
    s.row(["Style: Strike-through", "Strike-through Text"], styles: [nil, "strike"])

    s.row(["Align: Superscript", "x2 (2 is super)"], styles: %w[bg_light superscript])
    s.row(["Align: Subscript", "H2O (2 is sub)"], styles: [nil, "subscript"])
  end
end
puts "Created #{output_path}"
