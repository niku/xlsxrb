# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "interactive_validation_text_length.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.sheet("Text Length") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(["Username", "Enter (< 10 chars):"])
    s.validate_data("B2", type: "textLength", operator: "lessThan", formula1: "10", show_error_message: true, error_title: "Too Long", error: "Username must be under 10 characters")
  end
end

# 2. Read the generated sheet and print data validations
puts "=== Read Validation ==="
reader = Xlsxrb::Ooxml::Reader.new(output_path)
workbook = Xlsxrb.read(output_path).load
sheet = workbook.sheets.first
validations = reader.data_validations(sheet: sheet.name)
validations.each do |v|
  puts "Validation range #{v[:sqref]}: type=#{v[:type]}, formula1=#{v[:formula1]}, formula2=#{v[:formula2]}"
end
