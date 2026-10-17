# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "page_margins_wide.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.sheet("Wide Margins") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.page_margins(top: 1.0, bottom: 1.0, left: 1.0, right: 1.0, header: 0.5, footer: 0.5)
    s.row(["Wide Margins sheet"])
  end
end
