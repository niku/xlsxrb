# frozen_string_literal: true
module Xlsxrb
  VERSION = "0.1.4"
  def self.in_span(name, attributes: nil, &block)
    if defined?(Ractor) && Ractor.current != Ractor.main
      yield
    else
      unless defined?(@tracer)
        require "opentelemetry/sdk"
        @tracer = OpenTelemetry.tracer_provider.tracer("xlsxrb", VERSION)
      end
      if attributes
        @tracer.in_span(name, attributes: attributes, &block)
      else
        @tracer.in_span(name, &block)
      end
    end
  end
end

Ractor.new do
  Xlsxrb.in_span("test") { puts "hello from ractor" }
end.take
