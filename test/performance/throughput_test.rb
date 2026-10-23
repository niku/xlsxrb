# frozen_string_literal: true

require "test_helper"
require "tempfile"

class ThroughputTest < Test::Unit::TestCase
  def setup
    omit "Skip performance throughput tests under RBS runtime testing to prevent slowdown" if defined?(RBS::Test)
  end

  def test_deterministic_allocation_bounds_streaming_write_and_read
    Tempfile.create(["throughput_alloc", ".xlsx"]) do |f|
      GC.start
      GC.disable
      alloc_before_write = GC.stat(:total_allocated_objects)

      Xlsxrb.write(f.path) do |wb|
        wb.sheet("Data") do |sheet|
          10_000.times do |i|
            sheet.row(["Row #{i}", i, "Status #{i}", 42.5])
          end
        end
      end

      alloc_after_write = GC.stat(:total_allocated_objects)
      GC.enable

      write_allocated = alloc_after_write - alloc_before_write
      write_per_row = write_allocated.to_f / 10_000

      # Invariant: Streaming write must allocate fewer than 25 objects per row (< 250,000 total for 10k rows)
      assert_operator write_allocated, :<, 250_000,
                      "Streaming write allocated #{write_allocated} objects (#{write_per_row.round(2)} per row), expected < 250,000"

      GC.start
      GC.disable
      alloc_before_read = GC.stat(:total_allocated_objects)
      row_count = 0

      Xlsxrb.read(f.path) do |sheet|
        sheet.each do |_row|
          row_count += 1
        end
      end

      alloc_after_read = GC.stat(:total_allocated_objects)
      GC.enable

      read_allocated = alloc_after_read - alloc_before_read
      read_per_row = read_allocated.to_f / 10_000

      assert_equal 10_000, row_count
      # Invariant: Streaming read must allocate fewer than 15 objects per row (< 150,000 total for 10k rows)
      assert_operator read_allocated, :<, 150_000,
                      "Streaming read allocated #{read_allocated} objects (#{read_per_row.round(2)} per row), expected < 150,000"
    end
  end

  def test_scaling_ratio_complexity_order_linear_growth
    measure_pipeline = lambda do |n|
      Tempfile.create(["scaling_ratio_#{n}", ".xlsx"]) do |f|
        GC.start
        alloc_before_write = GC.stat(:total_allocated_objects)
        t0_write = Process.clock_gettime(Process::CLOCK_MONOTONIC)

        Xlsxrb.write(f.path) do |wb|
          wb.sheet("Data") do |sheet|
            n.times do |i|
              sheet.row(["Row #{i}", i, "Status #{i}", 42.5])
            end
          end
        end

        t1_write = Process.clock_gettime(Process::CLOCK_MONOTONIC)
        alloc_after_write = GC.stat(:total_allocated_objects)

        alloc_before_read = GC.stat(:total_allocated_objects)
        t0_read = Process.clock_gettime(Process::CLOCK_MONOTONIC)
        read_count = 0

        Xlsxrb.read(f.path) do |sheet|
          sheet.each { |_r| read_count += 1 }
        end

        t1_read = Process.clock_gettime(Process::CLOCK_MONOTONIC)
        alloc_after_read = GC.stat(:total_allocated_objects)

        assert_equal n, read_count

        {
          write_time: [t1_write - t0_write, 0.0001].max,
          read_time: [t1_read - t0_read, 0.0001].max,
          write_alloc: alloc_after_write - alloc_before_write,
          read_alloc: alloc_after_read - alloc_before_read
        }
      end
    end

    small = measure_pipeline.call(2_000)
    large = measure_pipeline.call(20_000)

    # 10x input size increase (2,000 -> 20,000)
    # For O(N) complexity, the scaling ratio should remain close to 10x.
    # An O(N^2) regression would explode to ~100x.
    write_time_ratio = large[:write_time] / small[:write_time]
    read_time_ratio = large[:read_time] / small[:read_time]
    write_alloc_ratio = large[:write_alloc].to_f / small[:write_alloc]
    read_alloc_ratio = large[:read_alloc].to_f / small[:read_alloc]

    # Deterministic allocation ratio check: must stay <= 15.0x
    assert_operator write_alloc_ratio, :<=, 15.0,
                    "Write allocation ratio was #{write_alloc_ratio.round(2)}x, expected <= 15.0x for O(N) linear complexity"
    assert_operator read_alloc_ratio, :<=, 15.0,
                    "Read allocation ratio was #{read_alloc_ratio.round(2)}x, expected <= 15.0x for O(N) linear complexity"

    # Time ratio check with generous threshold to absorb CI load jitter (<= 25.0x)
    assert_operator write_time_ratio, :<=, 25.0,
                    "Write time ratio was #{write_time_ratio.round(2)}x, expected <= 25.0x"
    assert_operator read_time_ratio, :<=, 25.0,
                    "Read time ratio was #{read_time_ratio.round(2)}x, expected <= 25.0x"
  end
end
