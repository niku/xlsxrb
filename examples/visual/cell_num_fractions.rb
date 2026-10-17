# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cell_num_fractions.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("frac") { |s| s.num_fmt("# ?/?") }
  wb.sheet("Fractions") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(%w[Format Value])
    s.row(["Half", 0.5], styles: { 1 => "frac" })
    s.row(["Third", 0.3333], styles: { 1 => "frac" })
    s.row(["Quarter", 0.75], styles: { 1 => "frac" })
  end
end
