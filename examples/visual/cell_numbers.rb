# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cell_numbers.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("currency") { |s| s.num_fmt("$#,##0.00") }
  wb.style("percent") { |s| s.num_fmt("0.0%") }
  wb.sheet("Numbers") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(%w[Format Value])
    s.row(["Integer", 12_345])
    s.row(["Float", 123.456])
    s.row(["Currency", 1234.5], styles: { 1 => "currency" })
    s.row(["Percentage", 0.85], styles: { 1 => "percent" })
  end
end
