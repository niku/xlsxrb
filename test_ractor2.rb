# frozen_string_literal: true
require_relative "lib/xlsxrb"
Ractor.new do
  puts Xlsxrb::VERSION
  tracer = OpenTelemetry.tracer_provider.tracer("xlsxrb", Xlsxrb::VERSION)
  puts tracer
end.take
