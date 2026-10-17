# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cf_contains_text.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.sheet("CF Contains") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(["Status"])
    s.row(["Error"])
    s.row(["Success"])
    s.row(["Pending"])
    s.conditional_format("A2:A4", type: "containsText", operator: "containsText", text: "Error", formula: 'NOT(ISERROR(SEARCH("Error",A2)))', fill_color: :red)
  end
end
