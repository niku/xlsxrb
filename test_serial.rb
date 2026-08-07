# frozen_string_literal: true

require "date"
require "./lib/xlsxrb"

def print_serial(date, date1904)
  wb = Xlsxrb.build do |w|
    w.workbook_property(:date1904, true) if date1904
    w.sheet "S" do |s|
      s.row [date]
    end
  end
  Xlsxrb.write("tmp.xlsx", wb)
  puts Xlsxrb.read("tmp.xlsx").sheets.first.rows.first.cells.first.value
end

puts "1904: Date.new(1904, 1, 1) -> #{print_serial(Date.new(1904, 1, 1), true)}"
puts "1904: Date.new(1904, 1, 2) -> #{print_serial(Date.new(1904, 1, 2), true)}"
puts "1904: Date.new(1899, 12, 31) -> #{print_serial(Date.new(1899, 12, 31), true)}"
puts "1900: Date.new(1900, 2, 28) -> #{print_serial(Date.new(1900, 2, 28), false)}"
puts "1900: Date.new(1900, 3, 1) -> #{print_serial(Date.new(1900, 3, 1), false)}"
