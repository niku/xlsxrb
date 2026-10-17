# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "interactive_validation_list.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.sheet("List Validation") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(["Department", "Select:"])
    s.validate_data("B2", in: %w[HR Sales Engineering])
  end
end
