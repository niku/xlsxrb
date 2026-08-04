# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "fill_gradients.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.add_style("gradient") do |style|
    style.fill_gradient(type: "linear", degree: 45, stops: [
                          { position: 0, color: "FFFFFFFF" },
                          { position: 1, color: "FF4F81BD" }
                        ])
  end
  w.sheet("Gradients") do |s|
    s.set_column(0, width: 25)
    s.set_column(1, width: 25)
    s.row(["Normal Cell", "Gradient Cell"], styles: { 1 => "gradient" })
  end
end

# 2. Read the generated sheet and print cell fill properties
puts "=== Read Validation ==="
workbook = Xlsxrb.read(output_path)
sheet = workbook.sheets.first
row = sheet.rows.first
row.cells.each do |c|
  xf = workbook.styles[:cell_xfs][c.style_index] if c.style_index
  fill = xf ? workbook.styles[:fills][xf[:fill_id]] : nil
  puts "Cell #{c.ref} ('#{c.value}'): fill pattern = #{fill&.[](:pattern).inspect}, fg_color = #{fill&.[](:fg_color).inspect}"
end
