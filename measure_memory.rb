# frozen_string_literal: true

require "opentelemetry/sdk"
require "objspace"
require_relative "lib/xlsxrb"

# Outputs memory and GC stats for a given tracing span
class MemorySpanProcessor < OpenTelemetry::SDK::Trace::SpanProcessor
  def on_start(span, _parent_context)
    span.set_attribute("ruby.gc.count.start", GC.stat[:count])
    span.set_attribute("ruby.memory.bytes_allocated.start", ObjectSpace.memsize_of_all)
  end

  def on_finish(span)
    gc_start = span.attributes["ruby.gc.count.start"]
    mem_start = span.attributes["ruby.memory.bytes_allocated.start"]

    gc_diff = GC.stat[:count] - (gc_start || 0)
    mem_diff = ObjectSpace.memsize_of_all - (mem_start || 0)

    mem_mb = mem_diff.to_f / 1024 / 1024

    puts "[#{span.name}] Memory delta: #{mem_mb.round(2)} MB, GC count: #{gc_diff}"
  end
end

OpenTelemetry::SDK.configure do |c|
  c.add_span_processor(MemorySpanProcessor.new)
end

puts "Generating 100,000 cells (10,000 rows x 10 cols) to measure memory..."

# Using a smaller size first so it runs quickly
Xlsxrb.generate("test_memory.xlsx") do |w|
  w.add_sheet("Data") do
    10_000.times do |i|
      w.add_row(["A", "B", "C", "D", "E", "F", "G", "H", "I", i])
    end
  end
end

puts "Done."
