# frozen_string_literal: true

require "xlsxrb"

output_path = ARGV[0] || "fonts.xlsx"

Xlsxrb.generate(output_path) do |w|
  # Font Families
  w.add_style("f_arial") { |s| s.font_name("Arial") }
  w.add_style("f_times") { |s| s.font_name("Times New Roman") }
  w.add_style("f_courier") { |s| s.font_name("Courier New") }
  w.add_style("f_georgia") { |s| s.font_name("Georgia") }
  w.add_style("f_tahoma") { |s| s.font_name("Tahoma") }

  # Font Sizes
  w.add_style("sz_10") { |s| s.size(10) }
  w.add_style("sz_16") { |s| s.size(16) }
  w.add_style("sz_24") { |s| s.size(24) }

  # Font Colors
  w.add_style("c_red") { |s| s.font_color("FFC00000") }
  w.add_style("c_green") { |s| s.font_color("FF008000") }
  w.add_style("c_blue") { |s| s.font_color("FF0000FF") }

  # Font Styles
  w.add_style("st_bold", &:bold)
  w.add_style("st_italic", &:italic)
  w.add_style("st_underline") { |s| s.underline("single") }
  w.add_style("st_double_u") { |s| s.underline("double") }
  w.add_style("st_strike", &:strike)

  # Vertical Alignments (Superscript/Subscript)
  w.add_style("v_super") { |s| s.vert_align("superscript") }
  w.add_style("v_sub") { |s| s.vert_align("subscript") }

  w.sheet("Fonts") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.set_print_option(:grid_lines, true)

    s.row(["Font Feature", "Text Preview"])
    s.row(["Family: Arial", "Arial Text"], styles: { 1 => "f_arial" })
    s.row(["Family: Times New Roman", "Times New Roman"], styles: { 1 => "f_times" })
    s.row(["Family: Courier New", "Courier New Text"], styles: { 1 => "f_courier" })
    s.row(["Family: Georgia", "Georgia Text"], styles: { 1 => "f_georgia" })
    s.row(["Family: Tahoma", "Tahoma Text"], styles: { 1 => "f_tahoma" })

    s.row(["Size: 10pt", "10pt Font Size"], styles: { 1 => "sz_10" })
    s.row(["Size: 16pt", "16pt Font Size"], styles: { 1 => "sz_16" })
    s.row(["Size: 24pt", "24pt Font Size"], styles: { 1 => "sz_24" })

    s.row(["Color: Red", "Red Text"], styles: { 1 => "c_red" })
    s.row(["Color: Green", "Green Text"], styles: { 1 => "c_green" })
    s.row(["Color: Blue", "Blue Text"], styles: { 1 => "c_blue" })

    s.row(["Style: Bold", "Bold Text"], styles: { 1 => "st_bold" })
    s.row(["Style: Italic", "Italic Text"], styles: { 1 => "st_italic" })
    s.row(["Style: Underline", "Underline Text"], styles: { 1 => "st_underline" })
    s.row(["Style: Double Underline", "Double Underline"], styles: { 1 => "st_double_u" })
    s.row(["Style: Strike-through", "Strike-through Text"], styles: { 1 => "st_strike" })

    s.row(["Align: Superscript", "x2 (2 is super)"], styles: { 1 => "v_super" })
    s.row(["Align: Subscript", "H2O (2 is sub)"], styles: { 1 => "v_sub" })
  end
end

# Read check
puts "=== Read Validation ==="
workbook = Xlsxrb.read(output_path)
sheet = workbook.sheets.first
sheet.rows.each do |row|
  row_cells = row.cells.map do |c|
    xf = workbook.styles[:cell_xfs][c.style_index] if c.style_index
    font_id = xf ? xf[:font_id] : 0
    "#{c.ref}: #{c.value.inspect} (font_id: #{font_id})"
  end
  puts "Row #{row.index}: #{row_cells.join(", ")}"
end
