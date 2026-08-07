# frozen_string_literal: true

require "pbt"

puts "Testing Float..."
begin
  Pbt.assert(num_runs: 50) do
    Pbt.property(Pbt.float) do |f|
      raise "Fail" if f > 0.5
    end
  end
rescue StandardError => e
  puts e.message
end
puts "Float done."

puts "Testing Time..."
begin
  Pbt.assert(num_runs: 50) do
    Pbt.property(Pbt.integer(min: 946_684_800, max: 1_893_456_000).map(->(i) { Time.at(i).utc }, lambda(&:to_i))) do |t|
      raise "Fail" if t.year > 2010
    end
  end
rescue StandardError => e
  puts e.message
end
puts "Time done."
