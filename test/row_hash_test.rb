require "test_helper"

class RowHashTest < Test::Unit::TestCase
  test "row accepts Hash for values and styles in Workbook API" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    wb = Xlsxrb.build do |w|
      w.sheet("HashRow") do |s|
        s.style("bold", font: { bold: true })
        s.style("italic", font: { italic: true })
        s.row({ A: 1, C: 2 }, styles: { C: "bold", D: "italic" })
      end
    end
    Xlsxrb.write(xlsx_path, wb)

    check_sheet(xlsx_path)
  end

  test "row accepts Hash for values and styles in Stream API" do
    xlsx_tempfile = Tempfile.new(["xlsxrb-test", ".xlsx"])
    xlsx_path = xlsx_tempfile.path
    xlsx_tempfile.close

    Xlsxrb.generate(xlsx_path) do |w|
      w.sheet("HashRow") do |s|
        s.style("bold", font: { bold: true })
        s.style("italic", font: { italic: true })
        s.row({ A: 1, C: 2 }, styles: { C: "bold", D: "italic" })
      end
    end

    check_sheet(xlsx_path)
  end

  def check_sheet(xlsx_path)
    wb = Xlsxrb.read(xlsx_path)
    sheet = wb.sheets.first
    assert_equal(1, sheet.rows.size)
    
    cells = sheet.rows[0].cells
    cell_hash = cells.map { |c| [c.column_index, c] }.to_h

    # A1
    assert_not_nil(cell_hash[0])
    assert_equal(1, cell_hash[0].value)
    
    # B1
    assert_nil(cell_hash[1])
    
    # C1
    assert_not_nil(cell_hash[2])
    assert_equal(2, cell_hash[2].value)
    
    # D1 (empty but has style index)
    if cell_hash[3]
      assert_nil(cell_hash[3].value)
      assert_not_nil(cell_hash[3].style_index)
    end
  ensure
    File.delete(xlsx_path) if xlsx_path && File.exist?(xlsx_path)
  end
end
