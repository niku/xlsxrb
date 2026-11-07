# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cell_num_scientific.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("sci") { |s| s.number_format("0.00E+00") }
  wb.sheet("Scientific") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(%w[Format Value])
    s.row(["Scientific", 123_456_789.0], cell_styles: { 1 => "sci" })
    s.row(["Small Scientific", 0.00001234], cell_styles: { 1 => "sci" })
  end
end
