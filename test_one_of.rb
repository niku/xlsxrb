# frozen_string_literal: true

require "pbt"

begin
  Pbt.assert(num_runs: 50) do
    Pbt.property(
      Pbt.one_of(
        Pbt.integer,
        Pbt.boolean,
        Pbt.constant(nil),
        Pbt.printable_ascii_string(max: 20),
        Pbt.float,
        Pbt.integer(min: 946_684_800, max: 1_893_456_000).map(->(i) { Time.at(i).utc }, lambda(&:to_i))
      )
    ) do |v|
      raise "Fail" if v.is_a?(String) && v.include?("A")
    end
  end
rescue StandardError => e
  puts e.message
end
