# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cell_rich_text.xlsx"
Xlsxrb.generate(output_path) do |wb|
  rt = Xlsxrb.rich_text(
    { text: "Normal " },
    { text: "BOLD RED ", font: { bold: true, color: "FFC00000", sz: 16 } },
    { text: "ITALIC BLUE", font: { italic: true, color: :blue, sz: 20 } }
  )
  wb.sheet("Rich Text") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(%w[Format Value])
    s.row(["Rich Text Cell", rt])
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
