# frozen_string_literal: true

require "test_helper"
require "memory_profiler"
require "benchmark/ips"

class PerformanceTest < Test::Unit::TestCase
  def setup
    omit "Skip performance test under RBS runtime testing to prevent OOM" if defined?(RBS::Test)
  end

  def test_streaming_read_memory_usage
    filename = "test_10k_rows.xlsx"
    Xlsxrb.write(filename) do |wb|
      wb.sheet("Data") do |sheet|
        10_000.times { |i| sheet.row(["Row #{i}", i, "Status #{i}"]) }
      end
    end

    report = MemoryProfiler.report do
      Xlsxrb.read(filename) do |sheet|
        sheet.each do |row|
          # Just iterate
        end
      end
    end

    # Print the report to stdout for debugging
    # report.pretty_print

    # Streaming read should not retain memory relative to the file size
    # A generous threshold of 5MB for retained memory
    assert_operator report.total_retained_memsize, :<, 5_000_000

    FileUtils.rm_f(filename)
  end
end
