# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "interactive_comments.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.sheet("Comments") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(["Item A", "Item B"])
    s.comment("A1", "This is an important comment.", author: "System")
  end
end
