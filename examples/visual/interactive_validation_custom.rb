# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "interactive_validation_custom.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.sheet("Custom Rule") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(["Number A", "Number B (Must be larger A)"])
    s.row([10, ""])
    s.validate_data("B2", type: "custom", formula1: "B2>A2", show_error_message: true, error_title: "Validation Error", error: "Number B must be greater than Number A")
  end
end
