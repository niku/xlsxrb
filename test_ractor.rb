# frozen_string_literal: true
require_relative "lib/xlsxrb"
Ractor.new do
  puts Xlsxrb::VERSION
  puts Xlsxrb::TRACER
end.take
