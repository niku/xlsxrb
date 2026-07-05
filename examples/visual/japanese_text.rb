# frozen_string_literal: true

require "xlsxrb"

output_path = ARGV[0] || "japanese_text.xlsx"

Xlsxrb.generate(output_path) do |w|
  w.add_style("ja_font") do |style|
    style.font_name("Noto Sans CJK JP").size(12)
  end

  w.add_sheet("Japanese") do |s|
    s.set_column(0, width: 25)
    s.set_column(1, width: 25)
    s.add_row(%w[日本語ラベル 値], styles: { 0 => "ja_font", 1 => "ja_font" })
    s.add_row(["売上", 12_500], styles: { 0 => "ja_font" })
  end
end

# 2. Read the generated sheet and print parsed cell values and Ruby classes
puts "=== Read Validation ==="
workbook = Xlsxrb.read(output_path)
sheet = workbook.sheets.first
sheet.rows.each do |row|
  row_cells = row.cells.map do |c|
    "#{c.ref}: #{c.value.inspect} (#{c.value.class})"
  end
  puts "Row #{row.index}: #{row_cells.join(", ")}"
end
