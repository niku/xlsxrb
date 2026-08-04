# frozen_string_literal: true

require "xlsxrb"

output_path = ARGV[0] || "styles_fonts_fills.xlsx"

Xlsxrb.generate(output_path) do |w|
  w.add_style("header") do |style|
    style.bold.size(14).font_color("FFFFFFFF").fill_color("FF4F81BD")
  end

  w.add_style("highlight") do |style|
    style.italic.font_color("FFC00000").fill_color("FFFFFF00")
  end

  w.sheet("Styles") do
    w.column(0, width: 25)
    w.column(1, width: 25)
    w.row(["Header 1", "Header 2"], styles: { 0 => "header", 1 => "header" })
    w.row(["Normal Text", "Highlighted Text"], styles: { 1 => "highlight" })
  end
end

# 2. Read the generated sheet and print styling details
puts "=== Read Validation ==="
workbook = Xlsxrb.read(output_path)
sheet = workbook.sheets.first
sheet.rows.first(4).each do |row|
  row_cells = row.cells.map do |c|
    xf = workbook.styles[:cell_xfs][c.style_index] if c.style_index
    font = xf ? workbook.styles[:fonts][xf[:font_id]] : nil
    fill = xf ? workbook.styles[:fills][xf[:fill_id]] : nil
    "#{c.ref}: #{c.value.inspect} (font=#{font&.[](:name)}, fill=#{fill&.[](:fg_color)&.[](:rgb)})"
  end
  puts "Row #{row.index}: #{row_cells.join(", ")}"
end
