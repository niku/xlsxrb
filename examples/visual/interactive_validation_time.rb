# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "interactive_validation_time.xlsx"
Xlsxrb.generate(output_path) do |wb|
  wb.sheet("Time Validation") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(["Schedule", "Enter Time (after 08:00):"])
    s.validate_data("B2", type: "time", operator: "greaterThan", formula1: "0.33333", show_error_message: true, error_title: "Too Early", error: "Time must be after 08:00 AM")
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
