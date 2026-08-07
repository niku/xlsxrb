# frozen_string_literal: true

require_relative "lib/xlsxrb"

begin
  wb = Xlsxrb.build do |w|
    # 1. Duplicate sheet names
    w.sheet("Duplicate1") do |s|
      s.row %w[A B C]
    end
    w.sheet("Duplicate2") do |s|
      s.row %w[D E F]
    end

    # 3. Control characters
    w.sheet("SpecialChars") do |s|
      s.row ["Null\x00byte", "Bell\a", "Emoji 👨‍👩‍👧‍👦"]
    end

    # 4. Overlapping/Invalid merged cells
    w.sheet("InvalidMerge") do |s|
      s.row ["Merge"]
      s.merge("A1:B1")
      s.merge("A2:B2")
    end
  end

  # Let's see what happens on write
  Xlsxrb.write("test_weird.xlsx", wb)
  puts "Write succeeded."

  # Read it back
  read_wb = Xlsxrb.read("test_weird.xlsx")
  puts "Read succeeded. Sheets: #{read_wb.sheets.map(&:name).join(", ")}"
rescue StandardError => e
  puts "Failed: #{e.class} - #{e.message}"
  puts e.backtrace.first(5)
end
