# frozen_string_literal: true

require "xlsxrb"

output_path = ARGV[0] || "sheet_protection.xlsx"

Xlsxrb.write(output_path) do |wb|
  wb.style(:header, bold: true, fill_color: "4F81BD", font_color: "FFFFFF")
  wb.style(:cell, border_all: { style: :thin, color: "D9D9D9" })

  wb.sheet("Protected Sheet") do |s|
    s.column(0, width: 20)
    s.column(1, width: 30)

    s.row(%w[Setting Status], styles: :header)
    s.row(["Protection", "Enabled with password"], styles: :cell)
    s.row(["Select Locked", "Allowed"], styles: :cell)
    s.row(["Select Unlocked", "Allowed"], styles: :cell)

    # Protect worksheet against unauthorized edits
    s.protect_sheet(password: "secret", select_locked_cells: true, select_unlocked_cells: true)
  end
end
