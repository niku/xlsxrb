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

  w.add_sheet("Fonts") do |s|
    s.set_column(0, width: 25)
    s.set_column(1, width: 25)
    s.set_print_option(:grid_lines, true)

    s.add_row(["Font Feature", "Text Preview"])
    s.add_row(["Family: Arial", "Arial Text"], styles: { 1 => "f_arial" })
    s.add_row(["Family: Times New Roman", "Times New Roman"], styles: { 1 => "f_times" })
    s.add_row(["Family: Courier New", "Courier New Text"], styles: { 1 => "f_courier" })
    s.add_row(["Family: Georgia", "Georgia Text"], styles: { 1 => "f_georgia" })
    s.add_row(["Family: Tahoma", "Tahoma Text"], styles: { 1 => "f_tahoma" })

    s.add_row(["Size: 10pt", "10pt Font Size"], styles: { 1 => "sz_10" })
    s.add_row(["Size: 16pt", "16pt Font Size"], styles: { 1 => "sz_16" })
    s.add_row(["Size: 24pt", "24pt Font Size"], styles: { 1 => "sz_24" })

    s.add_row(["Color: Red", "Red Text"], styles: { 1 => "c_red" })
    s.add_row(["Color: Green", "Green Text"], styles: { 1 => "c_green" })
    s.add_row(["Color: Blue", "Blue Text"], styles: { 1 => "c_blue" })

    s.add_row(["Style: Bold", "Bold Text"], styles: { 1 => "st_bold" })
    s.add_row(["Style: Italic", "Italic Text"], styles: { 1 => "st_italic" })
    s.add_row(["Style: Underline", "Underline Text"], styles: { 1 => "st_underline" })
    s.add_row(["Style: Double Underline", "Double Underline"], styles: { 1 => "st_double_u" })
    s.add_row(["Style: Strike-through", "Strike-through Text"], styles: { 1 => "st_strike" })

    s.add_row(["Align: Superscript", "x2 (2 is super)"], styles: { 1 => "v_super" })
    s.add_row(["Align: Subscript", "H2O (2 is sub)"], styles: { 1 => "v_sub" })
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
