require "test_helper"

class EnumerableTest < Test::Unit::TestCase
  test "Workbook, Worksheet, and Row are enumerable" do
    wb = Xlsxrb.build do |w|
      w.sheet("S") do |s|
        s.row([1, "2", "2026-01-01"])
      end
    end
    
    assert wb.is_a?(Enumerable)
    assert wb.sheets.first.is_a?(Enumerable)
    assert wb.sheets.first.rows.first.is_a?(Enumerable)
    
    cells = wb.sheets.first.cells.to_a
    assert_equal 3, cells.size
    
    assert_equal 1, cells[0].to_i
    assert_equal 2.0, cells[1].to_f
    assert_equal "2", cells[1].to_s
    assert_equal Date.new(2026, 1, 1), cells[2].to_date
  end
end
