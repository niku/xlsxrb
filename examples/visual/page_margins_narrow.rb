# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "page_margins_narrow.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.sheet("Narrow Margins") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.page_margins(top: 0.25, bottom: 0.25, left: 0.25, right: 0.25, header: 0.1, footer: 0.1)
    s.row(["Narrow Margins sheet"])
  end
end
