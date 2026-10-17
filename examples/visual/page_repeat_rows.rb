# frozen_string_literal: true

require "xlsxrb"

output_path = ARGV[0] || "page_repeat_rows.xlsx"

Xlsxrb.write(output_path) do |wb|
  wb.style(:header, bold: true, fill_color: "4F81BD", font_color: "FFFFFF", align_horizontal: :center)
  wb.style(:border, border_all: { style: :thin, color: "D9D9D9" })

  wb.sheet("Repeat Rows") do |s|
    s.column(0, width: 12)
    s.column(1, width: 30)
    s.column(2, width: 20)

    s.row(%w[ID Description Status], styles: :header)
    40.times do |i|
      s.row([i + 1, "Report item #{i + 1}", i.even? ? "Completed" : "Pending"], styles: :border)
    end

    # Repeat top header row on every printed page
    s.repeat_rows("1:1")
    s.page_setup(orientation: :portrait, paper_size: 9)
  end
end
