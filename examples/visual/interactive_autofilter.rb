# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "interactive_autofilter.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.sheet("Filter") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(%w[Name Department])
    s.row(%w[Alice HR])
    s.row(%w[Bob Eng])
    s.auto_filter("A1:B3")
  end
  wb.defined_name("_xlnm._FilterDatabase", "Filter!$A$1:$B$3", sheet: "Filter", hidden: true)
end
