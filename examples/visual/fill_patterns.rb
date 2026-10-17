# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "fill_patterns.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("dark_gray") { |s| s.fill(pattern: "darkGray", fg_color: "FFC0C0C0", bg_color: "FFFFFFFF") }
  wb.style("grid_fill") { |s| s.fill(pattern: "darkGrid", fg_color: "FFC0C0C0", bg_color: "FFFFFFFF") }
  wb.sheet("Patterns") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(%w[Pattern Preview])
    s.row(["Dark Gray", "Pattern Fill"], styles: { 1 => "dark_gray" })
    s.row(["Dark Grid", "Grid Fill"], styles: { 1 => "grid_fill" })
  end
end
