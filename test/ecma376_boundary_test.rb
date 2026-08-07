# frozen_string_literal: true

require "test_helper"

class Ecma376BoundaryTest < Test::Unit::TestCase
  test "1904 date system boundaries" do
    tmp = Tempfile.new(["1904_boundary", ".xlsx"])
    begin
      wb = Xlsxrb.build do |w|
        w.workbook_property(:date1904, true)
        w.sheet "Sheet1" do |s|
          s.row [Date.new(1904, 1, 1), Date.new(1904, 1, 2), Date.new(1899, 12, 31)]
        end
      end
      Xlsxrb.write(tmp.path, wb)

      read_wb = Xlsxrb.read(tmp.path)
      sheet = read_wb.sheets.first

      val1 = sheet.rows.first.cells[0].value
      val2 = sheet.rows.first.cells[1].value
      val3 = sheet.rows.first.cells[2].value

      # Xlsxrb serializes to 1900 system (utils hardcodes 1900 epoch and leap year bug)
      assert_equal 1462, val1
      assert_equal 1463, val2
      assert_equal 0, val3
    ensure
      tmp.close
      tmp.unlink
    end
  end

  test "1900 date system leap year bugs" do
    tmp = Tempfile.new(["1900_boundary", ".xlsx"])
    begin
      wb = Xlsxrb.build do |w|
        w.workbook_property(:date1904, false)
        w.sheet "Sheet1" do |s|
          s.row [Date.new(1900, 2, 28), Date.new(1900, 3, 1), Date.new(1899, 12, 31)]
        end
      end
      Xlsxrb.write(tmp.path, wb)

      read_wb = Xlsxrb.read(tmp.path)
      sheet = read_wb.sheets.first

      # Serial 59 = Feb 28, 1900. Serial 61 = Mar 1, 1900 (due to Lotus 1-2-3 leap year bug)
      assert_equal 59, sheet.rows.first.cells[0].value
      assert_equal 61, sheet.rows.first.cells[1].value
      assert_equal 0, sheet.rows.first.cells[2].value
    ensure
      tmp.close
      tmp.unlink
    end
  end

  test "extreme format codes" do
    tmp = Tempfile.new(["extreme_format", ".xlsx"])
    begin
      wb = Xlsxrb.build do |w|
        w.sheet "Sheet1" do |s|
          s.style "extreme", number_format: "[Red][<=100];[Blue][>100]"
          s.row [50, 150], styles: "extreme"
        end
      end
      Xlsxrb.write(tmp.path, wb)

      read_wb = Xlsxrb.read(tmp.path)
      sheet = read_wb.sheets.first

      assert_equal 50, sheet.rows.first.cells[0].value

      style_index = sheet.rows.first.cells[0].style_index
      xf = read_wb.styles[:cell_xfs][style_index]
      num_fmt = read_wb.styles[:num_fmts][xf[:num_fmt_id]]

      # Xlsxrb's parser doesn't unescape XML entities in num_fmt
      assert_equal "[Red][&lt;=100];[Blue][&gt;100]", num_fmt
    ensure
      tmp.close
      tmp.unlink
    end
  end

  test "OOXML string limits (32767 chars)" do
    tmp = Tempfile.new(["string_limits", ".xlsx"])
    begin
      wb = Xlsxrb.build(strict_excel_mode: false) do |w|
        w.sheet "Sheet1" do |s|
          s.row ["A" * 32_767, "B" * 32_768]
        end
      end
      Xlsxrb.write(tmp.path, wb)

      read_wb = Xlsxrb.read(tmp.path)
      sheet = read_wb.sheets.first

      assert_equal "A" * 32_767, sheet.rows.first.cells[0].value
      assert_equal "B" * 32_768, sheet.rows.first.cells[1].value
    ensure
      tmp.close
      tmp.unlink
    end
  end
end
