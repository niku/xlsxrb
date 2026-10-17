# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "page_headings_print.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.sheet("Print Headings") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.print_options(:headings, true)
    s.row(["Row/Col Headings Printed sheet"])
  end
end
