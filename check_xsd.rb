# frozen_string_literal: true

require "nokogiri"
require "zip"

sml_xsd = nil
Dir.chdir(File.expand_path("test/fixtures/xsd", __dir__)) do
  sml_xsd = Nokogiri::XML::Schema(File.read("sml.xsd"))
end

["test_adv_1.xlsx", "test_adv_2.xlsx", "test_adv_6.xlsx", "test_adv_7.xlsx"].each do |file_path|
  puts "Checking #{file_path}..."
  begin
    Zip::File.open(file_path) do |zip|
      if (workbook_entry = zip.find_entry("xl/workbook.xml"))
        workbook_xml = Nokogiri::XML(workbook_entry.get_input_stream.read)
        errors = sml_xsd.validate(workbook_xml)
        puts "  workbook.xml errors: #{errors.join(", ")}" unless errors.empty?
      else
        puts "  Missing xl/workbook.xml!"
      end

      if (sheet_entry = zip.find_entry("xl/worksheets/sheet1.xml"))
        sheet_xml = Nokogiri::XML(sheet_entry.get_input_stream.read)
        errors = sml_xsd.validate(sheet_xml)
        puts "  sheet1.xml errors: #{errors.join(", ")}" unless errors.empty?
      else
        puts "  Missing xl/worksheets/sheet1.xml!"
      end
    end
  rescue StandardError => e
    puts "  Failed to open: #{e.message}"
  end
end
