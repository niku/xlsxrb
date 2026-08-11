# frozen_string_literal: true

require "test_helper"

class DateTimeStyleTest < Test::Unit::TestCase
  test "auto injects default date and time styles" do
    temp = Tempfile.new(["test_dates", ".xlsx"])
    Xlsxrb.generate(temp.path) do |w|
      w.sheet("S") do |s|
        s.row([Date.new(2026, 1, 1), Time.new(2026, 1, 1, 12, 0, 0)])
      end
    end
    wb = Xlsxrb.read(temp.path)
    sheet = wb.sheets.first
    date_cell = sheet.rows.first.cells[0]
    time_cell = sheet.rows.first.cells[1]

    # style_index should be > 0 (0 is normal)
    assert date_cell.style_index.positive?
    assert time_cell.style_index.positive?

    # Ensure numFmt is associated with these styles
    date_xf = wb.styles[:cell_xfs][date_cell.style_index]
    time_xf = wb.styles[:cell_xfs][time_cell.style_index]

    assert date_xf[:num_fmt_id].positive?
    assert time_xf[:num_fmt_id].positive?
  end
end
