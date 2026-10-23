# frozen_string_literal: true

require_relative "test_helper"

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
end
