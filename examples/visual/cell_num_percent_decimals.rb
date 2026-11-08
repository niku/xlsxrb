# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cell_num_percent_decimals.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("pct2") { |s| s.number_format = "0.00%" }
  wb.sheet("Percents") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(%w[Format Value])
    s.row(["Percent with 2 Decimals", 0.12345], cell_styles: { 1 => "pct2" })
  end
end
