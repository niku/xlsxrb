# frozen_string_literal: true
module Xlsxrb
  TRACER = String.new("hello")
  def self.in_span(name, &block)
    if defined?(Ractor) && Ractor.current != Ractor.main
      yield
    else
      puts TRACER
      yield
    end
  end
end

Ractor.new do
  Xlsxrb.in_span("test") { puts "hello from ractor" }
end.take
