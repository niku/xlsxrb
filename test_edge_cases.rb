require_relative "lib/xlsxrb"
require "date"

puts "Testing xlsxrb for edge cases and broken behavior..."

def assert_raises(exception_class)
  yield
  puts "FAIL: Expected #{exception_class} but nothing was raised!"
rescue exception_class => e
  puts "PASS: Raised #{exception_class} as expected: #{e.message[0..100]}"
rescue => e
  puts "FAIL: Raised #{e.class} instead of #{exception_class}: #{e.message}"
end

puts "\n--- 1. Sheet name constraints ---"
# Excel sheet names cannot exceed 31 chars and cannot contain: \ / ? * [ ]
begin
  Xlsxrb.generate("test_sheet_names.xlsx") do |wb|
    wb.sheet("ThisSheetNameIsWayTooLongForExcelToHandleOhMy") do |s|
      s.row([1, 2, 3])
    end
    wb.sheet("Invalid[Name]") do |s|
      s.row([1, 2, 3])
    end
  end
  puts "FAIL: Successfully built workbook with invalid sheet names! (xlsxrb lacks validation)"
rescue => e
  puts "PASS/FAIL? Raised #{e.class} on invalid sheet names: #{e.message}"
end

puts "\n--- 2. Exceeding row/col limits ---"
# Excel has 1,048,576 rows and 16,384 cols
begin
  Xlsxrb.generate("test_limits.xlsx") do |wb|
    wb.sheet("Limits") do |s|
      # Try to write a row at an invalid index (if xlsxrb allows specifying row indices, but it's sequential usually)
      # Let's try writing a massive row array
      massive_row = Array.new(16385, "A")
      s.row(massive_row)
    end
  end
  puts "FAIL: Successfully wrote a row with 16385 columns (exceeds Excel limits)"
rescue => e
  puts "PASS/FAIL? Raised #{e.class} on exceeding columns: #{e.message}"
end

puts "\n--- 3. Type Handling ---"
begin
  Xlsxrb.generate("test_types.xlsx") do |wb|
    wb.sheet("Types") do |s|
      # Try writing unsupported objects
      s.row([1, "text", {a: 1}, Object.new, Class, proc {}])
    end
  end
  puts "FAIL: Successfully wrote objects! Did it call to_s on them or just crash later?"
rescue => e
  puts "PASS: Raised #{e.class} when writing weird objects: #{e.message}"
end

puts "\n--- 4. Concurrency / State Leakage (assuming it's broken) ---"
# Does Xlsxrb use global state?
begin
  t1 = Thread.new do
    Xlsxrb.generate("t1.xlsx") do |wb|
      wb.sheet("S1") { |s| 100.times { s.row([1]) }; sleep 0.1 }
    end
  end
  t2 = Thread.new do
    Xlsxrb.generate("t2.xlsx") do |wb|
      wb.sheet("S2") { |s| 100.times { s.row([2]) }; sleep 0.1 }
    end
  end
  t1.join
  t2.join
  puts "PASS: Threads completed without crashing."
rescue => e
  puts "FAIL: Threads crashed: #{e.class} - #{e.message}"
end

puts "\n--- 5. Malformed API Usage ---"
begin
  Xlsxrb.generate("malformed.xlsx") do |wb|
    wb.sheet("Sheet1") do |s|
      s.row # no arguments?
      s.formula # missing args?
    end
  end
  puts "FAIL: Allowed malformed DSL calls"
rescue => e
  puts "PASS: Caught malformed DSL calls: #{e.class}"
end

puts "\n--- 6. Reading non-existent or corrupted files ---"
begin
  Xlsxrb.read("does_not_exist.xlsx") { |row| }
  puts "FAIL: Read non-existent file without error!"
rescue => e
  puts "PASS: Caught non-existent file: #{e.class}"
end

File.write("corrupt.xlsx", "This is not a zip file")
begin
  Xlsxrb.read("corrupt.xlsx") { |row| }
  puts "FAIL: Read corrupted file without error!"
rescue => e
  puts "PASS: Caught corrupted file: #{e.class}"
end
