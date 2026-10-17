# frozen_string_literal: true

require "xlsxrb"

output_path = ARGV[0] || "japanese_text.xlsx"

Xlsxrb.write(output_path) do |wb|
  wb.style("ja_font") do |style|
    style.font_name("Noto Sans CJK JP").size(12)
  end

  wb.sheet("Japanese") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(%w[日本語ラベル 値], styles: { 0 => "ja_font", 1 => "ja_font" })
    s.row(["売上", 12_500], styles: { 0 => "ja_font" })
  end
end
