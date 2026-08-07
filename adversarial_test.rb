# frozen_string_literal: true

require_relative "lib/xlsxrb"

def try_scenario(name)
  print "Scenario '#{name}'... "
  begin
    yield
    puts "Succeeded (Unexpected - should it have failed?)"
  rescue StandardError => e
    puts "Failed gracefully: #{e.class}"
  end
end

puts "--- Adversarial Testing ---"

# 1. Negative/Invalid row heights or column widths
try_scenario("Negative row height") do
  wb = Xlsxrb.build do |w|
    w.sheet("S1") do |s|
      s.row ["A"], height: -100
    end
  end
  Xlsxrb.write("test_adv_1.xlsx", wb)
end

# 2. Extremely large numbers / Infinite
try_scenario("Infinity in cell") do
  wb = Xlsxrb.build do |w|
    w.sheet("S1") do |s|
      s.row [Float::INFINITY, Float::NAN]
    end
  end
  Xlsxrb.write("test_adv_2.xlsx", wb)
end

# 3. Invalid UTF-8 String (Bad Encoding)
try_scenario("Invalid UTF-8 Encoding") do
  wb = Xlsxrb.build do |w|
    w.sheet("S1") do |s|
      bad_string = "Invalid\xFF\xFEString".force_encoding("UTF-8")
      s.row [bad_string]
    end
  end
  Xlsxrb.write("test_adv_3.xlsx", wb)
end

# 4. Same column mapping overwritten in hash
try_scenario("Duplicate column mapping hash") do
  wb = Xlsxrb.build do |w|
    w.sheet("S1") do |s|
      # A hash with A1 mapped multiple times (Ruby hashes don't allow duplicate keys, but we can do string vs symbol)
      s.row({ "A1" => 1, :A1 => 2 })
    end
  end
  Xlsxrb.write("test_adv_4.xlsx", wb)
end

# 5. Invalid Formula Type
try_scenario("Invalid Formula Structure") do
  wb = Xlsxrb.build do |w|
    w.sheet("S1") do |s|
      # Passing something weird as formula
      s.row [{ formula: { foo: "bar" } }]
    end
  end
  Xlsxrb.write("test_adv_5.xlsx", wb)
end

# 6. Stream Writer missing yield
try_scenario("StreamWriter no yield") do
  Xlsxrb.generate("test_adv_6.xlsx") do |stream|
    # Empty block, doesn't even define a sheet
  end
end

# 7. Extremely long sheet name (but exactly 31 chars but multi-byte)
try_scenario("31-char Multi-byte Sheet Name") do
  wb = Xlsxrb.build do |w|
    w.sheet("あ" * 31) do |s|
      s.row [1]
    end
  end
  Xlsxrb.write("test_adv_7.xlsx", wb)
end
