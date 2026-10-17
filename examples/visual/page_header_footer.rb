# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "page_header_footer.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.sheet("Header Footer") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.header_footer(odd_header: "&LMy Company&RPage &P", odd_footer: "&CConfidential")
    s.row(["Header Footer sheet"])
  end
end
