# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cf_ends_with.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.sheet("CF Ends") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(["Code"])
    s.row(["100-Z"])
    s.row(["200-Y"])
    s.row(["300-Z"])
    s.conditional_format("A2:A4", type: "endsWith", operator: "endsWith", text: "Z", formula: 'RIGHT(A2,1)="Z"', fill_color: "FFFFC7CE")
  end
end
