# frozen_string_literal: true
module Xlsxrb
  class DummyTracer
    def in_span(*)
      yield
    end
  end
  DUMMY = DummyTracer.new.freeze

  def self.in_span(name, attributes: nil, &block)
    if defined?(Ractor) && Ractor.current != Ractor.main
      yield
    else
      unless defined?(@tracer)
        @tracer = "fake_opentelemetry"
      end
      yield
    end
  end
end

Ractor.new do
  begin
    Xlsxrb.in_span("test") { puts "hello from ractor" }
  rescue => err
    puts "Failed to use: #{err.class} #{err.message}"
  end
end.take
