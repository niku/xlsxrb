# frozen_string_literal: true

require "xlsxrb"

output_path = ARGV[0] || "fonts.xlsx"

Xlsxrb.generate(output_path) do |w|
  # Font Families (using nested hash options)
  w.style("f_arial", font: { name: "Arial" })
  w.style("f_times", font: { name: "Times New Roman" })
  w.style("f_courier", font: { name: "Courier New" })
  w.style("f_georgia", font: { name: "Georgia" })
  w.style("f_tahoma", font: { name: "Tahoma" })

  # Font Sizes
  w.style("sz_10", font: { size: 10 })
  w.style("sz_16", font: { size: 16 })
  w.style("sz_24", font: { size: 24 })

  # Font Colors
  w.style("c_red", font: { color: "FFC00000" })
  w.style("c_green", font: { color: "FF008000" })
  w.style("c_blue", font: { color: :blue })

  # Font Styles
  w.style("bold", font: { bold: true })
  w.style("italic", font: { italic: true })
  w.style("underline", font: { underline: true })
  w.style("double_underline", font: { underline: "double" })
  w.style("strike", font: { strike: true })

  # Vertical Alignment (Subscript / Superscript)
  w.style("superscript", font: { vert_align: "superscript" })
  w.style("subscript", font: { vert_align: "subscript" })

  w.style("header", font: { bold: true, color: :white }, fill: { color: "FF4F81BD" })
  w.style("bg_light", fill: { color: "FFDCE6F1" })

  w.sheet("Fonts") do |s|
    s.column(0, width: 25)
    s.column(1, width: 35)

    s.row(["Font Feature", "Text Preview"], styles: "header")
    
    s.row(["Family: Arial", "Arial Text"], styles: ["bg_light", "f_arial"])
    s.row(["Family: Times New Roman", "Times New Roman"], styles: [nil, "f_times"])
    s.row(["Family: Courier New", "Courier New Text"], styles: ["bg_light", "f_courier"])
    s.row(["Family: Georgia", "Georgia Text"], styles: [nil, "f_georgia"])
    s.row(["Family: Tahoma", "Tahoma Text"], styles: ["bg_light", "f_tahoma"])
    
    s.row(["Size: 10pt", "10pt Font Size"], styles: [nil, "sz_10"])
    s.row(["Size: 16pt", "16pt Font Size"], styles: ["bg_light", "sz_16"])
    s.row(["Size: 24pt", "24pt Font Size"], styles: [nil, "sz_24"])
    
    s.row(["Color: Red", "Red Text"], styles: ["bg_light", "c_red"])
    s.row(["Color: Green", "Green Text"], styles: [nil, "c_green"])
    s.row(["Color: Blue", "Blue Text"], styles: ["bg_light", "c_blue"])
    
    s.row(["Style: Bold", "Bold Text"], styles: [nil, "bold"])
    s.row(["Style: Italic", "Italic Text"], styles: ["bg_light", "italic"])
    s.row(["Style: Underline", "Underline Text"], styles: [nil, "underline"])
    s.row(["Style: Double Underline", "Double Underline"], styles: ["bg_light", "double_underline"])
    s.row(["Style: Strike-through", "Strike-through Text"], styles: [nil, "strike"])
    
    s.row(["Align: Superscript", "x2 (2 is super)"], styles: ["bg_light", "superscript"])
    s.row(["Align: Subscript", "H2O (2 is sub)"], styles: [nil, "subscript"])
  end
end
puts "Created #{output_path}"
