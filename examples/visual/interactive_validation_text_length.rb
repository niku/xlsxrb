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
