# frozen_string_literal: true

require "nokogiri"

begin
  # Change dir so relative schema locations work
  Dir.chdir("test/fixtures/xsd/sml") do
    Nokogiri::XML::Schema(File.read("sml_ECMA376_4ed_transitional.xsd"))
    puts "Schema parsed successfully"
  end
rescue StandardError => e
  puts "Error: #{e.message}"
end
