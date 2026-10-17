# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cell_rich_text.xlsx"
Xlsxrb.write(output_path) do |wb|
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
