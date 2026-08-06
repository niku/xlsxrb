require_relative "test_helper"
require "memory_profiler"
require "benchmark/ips"

class BenchmarkTest < Test::Unit::TestCase
  def test_streaming_read_memory_usage
    filename = "test_10k_rows.xlsx"
    Xlsxrb.generate(filename) do |wb|
      wb.sheet("Data") do |sheet|
        10_000.times { |i| sheet.row(["Row #{i}", i, "Status #{i}"]) }
      end
    end

    report = MemoryProfiler.report do
      Xlsxrb.foreach(filename) do |sheet|
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

    File.delete(filename) if File.exist?(filename)
  end
end
