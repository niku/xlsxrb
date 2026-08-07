# frozen_string_literal: true

require "pbt"

puts "Testing Float NaN..."
begin
  Pbt.assert(num_runs: 50) do
    Pbt.property(Pbt.float) do |f|
      raise "Fail NaN" if f.nan?
    end
  end
rescue StandardError => e
  puts e.message
end
