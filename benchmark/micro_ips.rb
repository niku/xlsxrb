# frozen_string_literal: true

require "benchmark/ips"
require_relative "../lib/xlsxrb"

puts "=" * 70
puts "xlsxrb Hot-Path Microbenchmarks (benchmark-ips)"
puts "Ruby: #{RUBY_DESCRIPTION}"
puts "=" * 70

ips_time = (ENV["IPS_TIME"] || "2").to_f
ips_warmup = (ENV["IPS_WARMUP"] || "1").to_f

# 1. XML escaping hot path
sample_escape_string = "Customer & Company <sales@example.com> \"Q1 '2026'\""
puts "\n--- 1. XmlBuilder.escape ---"
Benchmark.ips do |x|
  x.config(time: ips_time, warmup: ips_warmup)
  x.report("XmlBuilder.escape") do
    Xlsxrb::Ooxml::XmlBuilder.escape(sample_escape_string)
  end
  x.compare!
end

# 2. Coordinate conversion hot path
puts "\n--- 2. Cell.calculate_column_index ---"
Benchmark.ips do |x|
  x.config(time: ips_time, warmup: ips_warmup)
  x.report("Cell.calculate_column_index ('XFD1048576')") do
    Xlsxrb::Elements::Cell.calculate_column_index("XFD1048576")
  end
  x.report("Cell.calculate_column_letter (16384)") do
    Xlsxrb::Elements::Cell.calculate_column_letter(16_384)
  end
  x.compare!
end

# 3. Number formatting hot path
sample_numbers = [1234.5678, -42.0, 0.05, 1_000_000.99]
sample_format_code = "#,##0.00;(#,##0.00);\"-\""
puts "\n--- 3. NumberFormatter.format ---"
Benchmark.ips do |x|
  x.config(time: ips_time, warmup: ips_warmup)
  x.report("NumberFormatter.format (custom pattern)") do
    sample_numbers.each do |num|
      Xlsxrb::NumberFormatter.format(num, sample_format_code)
    end
  end
  x.report("NumberFormatter.format (currency)") do
    sample_numbers.each do |num|
      Xlsxrb::NumberFormatter.format(num, "$#,##0.00")
    end
  end
  x.compare!
end

# 4. WorksheetParser row scanning hot path
sample_row_xml = '<row r="1" spans="1:4"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>123.45</v></c><c r="C1" s="1"><v>44927</v></c><c r="D1" t="inlineStr"><is><t>Inline text</t></is></c></row>'
shared_strings = ["First String", "Second String"]
styles = {}
puts "\n--- 4. WorksheetParser Row Scanning ---"
Benchmark.ips do |x|
  x.config(time: ips_time, warmup: ips_warmup)
  x.report("WorksheetParser.fast_scan_row_values") do
    Xlsxrb::Ooxml::WorksheetParser.fast_scan_row_values(sample_row_xml, 0, sample_row_xml.bytesize, shared_strings, styles)
  end
  x.compare!
end

puts "\nMicrobenchmarks complete."
