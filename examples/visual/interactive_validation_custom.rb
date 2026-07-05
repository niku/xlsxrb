# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "interactive_validation_custom.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.add_sheet("Custom Rule") do |s|
    s.set_column(0, width: 25)
    s.set_column(1, width: 25)
    s.add_row(["Number A", "Number B (Must be larger A)"])
    s.add_row([10, ""])
    s.add_data_validation("B2", type: "custom", formula1: "B2>A2", show_error_message: true, error_title: "Validation Error", error: "Number B must be greater than Number A")
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
