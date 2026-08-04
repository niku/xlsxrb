# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "interactive_validation_date.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.sheet("Date Validation") do |s|
    s.set_column(0, width: 25)
    s.set_column(1, width: 25)
    s.add_row(["Date Range", "Enter Date (2026):"])
    s.add_data_validation("B2", type: "date", operator: "between", formula1: "Date(2026,1,1)", formula2: "Date(2026,12,31)", show_error_message: true, error_title: "Invalid Date", error: "Must be a date in 2026")
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
