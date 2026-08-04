# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "workbook_three_sheets.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.sheet("First Sheet") { |s| s.row(["First Sheet Data"]) }
  w.sheet("Second Sheet") { |s| s.row(["Second Sheet Data"]) }
  w.sheet("Third Sheet") { |s| s.row(["Third Sheet Data"]) }
end

# 2. Read the generated sheet and print the sheets structure
puts "=== Read Validation ==="
workbook = Xlsxrb.read(output_path)
puts "Workbook sheets: #{workbook.sheet_names.join(", ")}"
