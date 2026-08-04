# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "interactive_validation_range.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.sheet("Range Validation") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(["Age", "Enter (18-99):"])
    s.validate_data("B2", type: "whole", operator: "between", formula1: "18", formula2: "99", show_error_message: true, error_title: "Invalid Age", error: "Age must be between 18 and 99!")
  end
end

# 2. Read the generated sheet and print data validations
puts "=== Read Validation ==="
reader = Xlsxrb::Ooxml::Reader.new(output_path)
workbook = Xlsxrb.read(output_path)
sheet = workbook.sheets.first
validations = reader.data_validations(sheet: sheet.name)
validations.each do |v|
  puts "Validation range #{v[:sqref]}: type=#{v[:type]}, formula1=#{v[:formula1]}, formula2=#{v[:formula2]}"
end
