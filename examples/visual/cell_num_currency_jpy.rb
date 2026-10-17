# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cell_num_currency_jpy.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("jpy") { |s| s.num_fmt("¥#,##0;[Red]¥-#,##0") }
  wb.sheet("JPY Currency") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(%w[Format Value])
    s.row(["Positive Yen", 12_500], styles: { 1 => "jpy" })
    s.row(["Negative Yen", -8000], styles: { 1 => "jpy" })
  end
end
