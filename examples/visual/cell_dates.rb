# frozen_string_literal: true

require "xlsxrb"
require "date"
output_path = ARGV[0] || "cell_dates.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("custom_date") { |s| s.num_fmt("yyyy-mm-dd") }
  wb.sheet("Dates") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(["Format", "Date Value"])
    s.row(["Default Date", Date.new(2026, 7, 1)])
    s.row(["Formatted Date", Date.new(2026, 12, 25)], styles: { 1 => "custom_date" })
  end
end
