# frozen_string_literal: true

require_relative "test_helper"
require "zip"

class WorkbookViewAndCalcPropertiesTest < Test::Unit::TestCase
  def write_and_read(workbook_properties)
    workbook = Xlsxrb::Elements::Workbook.new(
      sheets: [Xlsxrb::Elements::Worksheet.new(name: "A"), Xlsxrb::Elements::Worksheet.new(name: "B")],
      unmapped_data: { workbook_properties: workbook_properties }
    )
    Tempfile.create(["workbook_view", ".xlsx"]) do |f|
      f.binmode
      f.write(Xlsxrb.write(workbook))
      f.close
      yield Xlsxrb::Ooxml::Reader.new(f.path)
    end
  end

  test "write emits the active tab and full calculation on load from workbook properties" do
    write_and_read({ update_links: "never", active_tab: 1, full_calc_on_load: true }) do |reader|
      assert_equal(1, reader.workbook_views[:active_tab])
      assert_equal(true, reader.calc_properties[:full_calc_on_load])
    end
  end

  test "write omits book views and calc properties unless they are requested" do
    write_and_read({ update_links: "never" }) do |reader|
      assert_nil(reader.workbook_views[:active_tab])
      assert_nil(reader.calc_properties[:full_calc_on_load])
    end
  end

  test "WorkbookBuilder supports active_sheet, first_sheet, show_sheet_tabs, and tab_ratio" do
    Tempfile.create(["wb_views", ".xlsx"]) do |f|
      wb = Xlsxrb.build do |b|
        b.sheet("First") { |s| s.row([1]) }
        b.sheet("Second", active: true) { |s| s.row([2]) }
        b.sheet("Third") { |s| s.row([3]) }
        b.first_sheet = "Second"
        b.show_sheet_tabs = false
        b.tab_ratio = 750
      end
      f.binmode
      f.write(Xlsxrb.write(wb))
      f.close

      read_wb = Xlsxrb.load(f.path)
      assert_equal(1, read_wb.active_sheet_index)
      assert_equal("Second", read_wb.active_sheet&.name)
      assert_equal(1, read_wb.first_sheet)
      assert_equal(false, read_wb.show_sheet_tabs?)
      assert_equal(750, read_wb.tab_ratio)

      # Verify sheet XML tabSelected uniqueness in zip
      Zip::File.open(f.path) do |zip|
        s1 = zip.read("xl/worksheets/sheet1.xml")
        s2 = zip.read("xl/worksheets/sheet2.xml")
        s3 = zip.read("xl/worksheets/sheet3.xml")

        assert_not_include(s1, 'tabSelected="1"')
        assert_include(s2, 'tabSelected="1"')
        assert_not_include(s3, 'tabSelected="1"')
      end
    end
  end

  test "StreamWriter supports active sheet, first_sheet, show_sheet_tabs, and tab_ratio" do
    Tempfile.create(["stream_views", ".xlsx"]) do |f|
      Xlsxrb.write(f.path) do |w|
        w.show_sheet_tabs = false
        w.tab_ratio = 600
        w.sheet("S1") { |s| s.row(["A"]) }
        w.sheet("S2", active: true) { |s| s.row(["B"]) }
        w.sheet("S3") { |s| s.row(["C"]) }
        w.first_sheet = "S2"
      end

      read_wb = Xlsxrb.load(f.path)
      assert_equal(1, read_wb.active_sheet_index)
      assert_equal("S2", read_wb.active_sheet&.name)
      assert_equal(1, read_wb.first_sheet)
      assert_equal(false, read_wb.show_sheet_tabs?)
      assert_equal(600, read_wb.tab_ratio)
    end
  end

  test "raises ArgumentError when setting hidden sheet as active sheet" do
    assert_raise(ArgumentError) do
      Xlsxrb.build do |b|
        b.sheet("Visible") { |s| s.row([1]) }
        b.sheet("Hidden", state: :hidden, active: true) { |s| s.row([2]) }
      end
    end

    assert_raise(ArgumentError) do
      Xlsxrb.build do |b|
        b.sheet("Visible") { |s| s.row([1]) }
        b.sheet("Hidden", state: :hidden) { |s| s.row([2]) }
        b.active_sheet = "Hidden"
      end
    end
  end

  test "automatically selects first visible sheet as active tab when first sheet is hidden" do
    Tempfile.create(["hidden_first", ".xlsx"]) do |f|
      wb = Xlsxrb.build do |b|
        b.sheet("HiddenSheet", state: :hidden) { |s| s.row([1]) }
        b.sheet("VisibleSheet") { |s| s.row([2]) }
      end
      f.binmode
      f.write(Xlsxrb.write(wb))
      f.close

      read_wb = Xlsxrb.load(f.path)
      assert_equal(1, read_wb.active_sheet_index)
      assert_equal("VisibleSheet", read_wb.active_sheet&.name)

      Zip::File.open(f.path) do |zip|
        s1 = zip.read("xl/worksheets/sheet1.xml")
        s2 = zip.read("xl/worksheets/sheet2.xml")

        assert_not_include(s1, 'tabSelected="1"')
        assert_include(s2, 'tabSelected="1"')
      end
    end
  end
end
