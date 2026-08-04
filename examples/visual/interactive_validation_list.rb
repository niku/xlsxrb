# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "interactive_validation_list.xlsx"
Xlsxrb.generate(output_path) do |w|
  w.sheet("List Validation") do |s|
    s.set_column(0, width: 25)
    s.set_column(1, width: 25)
    s.row(["Department", "Select:"])
    s.add_data_validation("B2", type: "list", formula1: '"HR,Sales,Engineering"')
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
