# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "cell_booleans.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.sheet("Booleans") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(["Label", "Boolean Value"])
    s.row(["Is Active", true])
    s.row(["Is Pending", false])
  end
end
