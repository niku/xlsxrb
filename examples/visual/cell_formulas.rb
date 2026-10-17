# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cell_formulas.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.sheet("Formulas") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(%w[Item Value])
    s.row(["A", 10])
    s.row(["B", 20])
    s.row(["SUM", Xlsxrb.formula("SUM(B2:B3)", cached_value: 30)])
    s.row(["AVERAGE", Xlsxrb.formula("AVERAGE(B2:B3)", cached_value: 15)])
  end
end
