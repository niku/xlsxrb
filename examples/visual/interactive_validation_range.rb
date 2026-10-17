# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "interactive_validation_range.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.sheet("Range Validation") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(["Age", "Enter (18-99):"])
    s.validate_data("B2", type: "whole", operator: "between", formula1: "18", formula2: "99", show_error_message: true, error_title: "Invalid Age", error: "Age must be between 18 and 99!")
  end
end
