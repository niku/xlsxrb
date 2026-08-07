# frozen_string_literal: true

require "test_helper"
require "nokogiri"
require "zip"

class XsdValidationTest < Test::Unit::TestCase
  def setup
    @sml_xsd = nil
    Dir.chdir(File.expand_path("fixtures/xsd", __dir__)) do
      @sml_xsd = Nokogiri::XML::Schema(File.read("sml.xsd"))
    end
  end

  def test_workbook_and_sheet_validation
    file_path = File.join(__dir__, "tmp_validation.xlsx")
    wb = Xlsxrb.build do |w|
      w.sheet("Sheet1") do |s|
        s.row ["Test", 123]
      end
    end
    Xlsxrb.write(file_path, wb)

    Zip::File.open(file_path) do |zip|
      # Validate workbook.xml
      workbook_entry = zip.find_entry("xl/workbook.xml")
      assert_not_nil workbook_entry, "xl/workbook.xml should exist"
      workbook_xml = Nokogiri::XML(workbook_entry.get_input_stream.read)

      errors = @sml_xsd.validate(workbook_xml)
      assert_empty errors, "xl/workbook.xml failed XSD validation: #{errors.join(", ")}"

      # Validate sheet1.xml
      sheet_entry = zip.find_entry("xl/worksheets/sheet1.xml")
      assert_not_nil sheet_entry, "xl/worksheets/sheet1.xml should exist"
      sheet_xml = Nokogiri::XML(sheet_entry.get_input_stream.read)

      errors = @sml_xsd.validate(sheet_xml)
      assert_empty errors, "xl/worksheets/sheet1.xml failed XSD validation: #{errors.join(", ")}"
    end
  ensure
    File.delete(file_path) if file_path && File.exist?(file_path)
  end
end
