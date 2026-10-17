# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cell_num_custom_colors.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("custom_color") { |s| s.num_fmt("[Green]#,##0;[Red]-#,##0") }
  wb.sheet("Custom Colors") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(%w[Format Value])
    s.row(["Positive (Green)", 5000], styles: { 1 => "custom_color" })
    s.row(["Negative (Red)", -2500], styles: { 1 => "custom_color" })
  end
end
