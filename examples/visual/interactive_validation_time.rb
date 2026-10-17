# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "interactive_validation_time.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.sheet("Time Validation") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(["Schedule", "Enter Time (after 08:00):"])
    s.validate_data("B2", type: "time", operator: "greaterThan", formula1: "0.33333", show_error_message: true, error_title: "Too Early", error: "Time must be after 08:00 AM")
  end
end
