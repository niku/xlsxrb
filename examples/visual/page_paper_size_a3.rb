# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "page_paper_size_a3.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.sheet("A3 Sheet") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.page_setup(paper_size: 8) # ISO A3 is paperSize="8"
    s.row(["ISO A3 Paper Size sheet"])
  end
end
