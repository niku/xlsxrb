# frozen_string_literal: true

require "xlsxrb"

output_path = ARGV[0] || "hyperlinks.xlsx"

Xlsxrb.write(output_path) do |wb|
  wb.style(:header, bold: true, fill_color: "4F81BD", font_color: "FFFFFF", align_horizontal: :center)
  wb.style(:link, font_color: "0000FF", underline: true)
  wb.style(:border, border_all: { style: :thin, color: "D9D9D9" })

  wb.sheet("Links") do |s|
    s.column(0, width: 22)
    s.column(1, width: 35)

    s.row(%w[Link_Type Target_Description], styles: :header)
    s.row(["External Web URL", "Visit Ruby-Lang"], styles: %i[border link])
    s.row(["Internal Sheet", "Jump to Details Sheet"], styles: %i[border link])

    s.hyperlink("B2", "https://www.ruby-lang.org", display: "Visit Ruby-Lang", tooltip: "Official Ruby Website")
    s.hyperlink("B3", location: "Details!A1", tooltip: "Go to Details sheet")
  end

  wb.sheet("Details") do |s|
    s.column(0, width: 30)
    s.row(["Target destination sheet (A1)"])
  end
end
