# frozen_string_literal: true
require_relative "lib/xlsxrb"
begin
  Ractor.make_shareable(Xlsxrb::TRACER)
rescue => e
  puts "Failed to share: #{e.class} #{e.message}"
end
Ractor.new do
  begin
    puts Xlsxrb::TRACER.in_span("test") { "hello" }
  rescue => err
    puts "Failed to use: #{err.class} #{err.message}"
  end
end.take
