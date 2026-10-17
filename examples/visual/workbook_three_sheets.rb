# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "workbook_three_sheets.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.sheet("First Sheet") { |s| s.row(["First Sheet Data"]) }
  wb.sheet("Second Sheet") { |s| s.row(["Second Sheet Data"]) }
  wb.sheet("Third Sheet") { |s| s.row(["Third Sheet Data"]) }
end
