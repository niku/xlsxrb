# frozen_string_literal: true

require "test_helper"

class ConcurrentTest < Test::Unit::TestCase
  def test_thread_safety
    threads = 5.times.map do |i|
      Thread.new do
        workbook = Xlsxrb.build do |w|
          w.sheet("Sheet#{i}") do |s|
            s.row(["Thread", i])
          end
        end
        [workbook.sheets.size, workbook.sheets[0].name]
      end
    end
    results = threads.map(&:value)
    results.each_with_index do |res, i|
      assert_equal [1, "Sheet#{i}"], res
    end
  end

  def test_ractor_safety
    omit "Ractors are not supported in this Ruby version" unless defined?(Ractor)
    omit "RBS::Test runtime hooks are incompatible with Ractor isolation" if defined?(RBS::Test)

    ractors = 5.times.map do |i|
      Ractor.new(i) do |idx|
        workbook = Xlsxrb.build do |w|
          w.sheet("Sheet#{idx}") do |s|
            s.row(["Ractor", idx])
          end
        end
        [workbook.sheets.size, workbook.sheets[0].name]
      rescue StandardError => e
        [e.class.name, e.message]
      end
    end

    results = ractors.map(&:value)
    results.each_with_index do |res, i|
      assert_equal [1, "Sheet#{i}"], res
    end
  end

  def test_all_xlsxrb_constants_are_ractor_shareable
    omit "Ractors are not supported in this Ruby version" unless defined?(Ractor)

    visited = Set.new
    unshareable = []

    scan_constants = lambda do |mod|
      return if visited.include?(mod)

      visited.add(mod)
      mod.constants(false).each do |c|
        val = mod.const_get(c)
        unshareable << "#{mod}::#{c} (#{val.class})" unless Ractor.shareable?(val)

        scan_constants.call(val) if val.is_a?(Module) && val.name&.start_with?("Xlsxrb")
      end
    end

    scan_constants.call(Xlsxrb)
    assert_empty unshareable, "Found non-shareable constants in Xlsxrb namespace: #{unshareable.join(", ")}"
  end

  def test_full_pipeline_ractor_parallelism
    omit "Ractors are not supported in this Ruby version" unless defined?(Ractor)
    omit "RBS::Test runtime hooks are incompatible with Ractor isolation" if defined?(RBS::Test)

    worker_count = 4
    ractors = worker_count.times.map do |idx|
      Ractor.new(idx) do |i|
        # 1. Write workbook with multiple data types and styles
        xlsx_bin = Xlsxrb.write do |wb|
          wb.sheet("Data") do |s|
            s.row(["Worker", i, 123.45, Date.new(2026, 10, 23), { formula: "B1*2", value: i * 2 }])
            s.style("A1", bold: true, color: "0000FF")
          end
          wb.sheet("Info") do |s|
            s.row(["Metadata", "Ractor-#{i}"])
          end
        end

        # 2. In-Memory read
        mem_wb = Xlsxrb.read(xlsx_bin)
        mem_sheet = mem_wb.sheets[0].load
        mem_worker_id = mem_sheet.rows[0].cells[1].value

        # 3. Streaming read
        stream_meta = nil
        Xlsxrb.read(xlsx_bin) do |sheet|
          next unless sheet.name == "Info"

          sheet.each do |row|
            stream_meta = row.values[1]
          end
        end

        # 4. Modify in-place via StringIO
        mod_io = StringIO.new(xlsx_bin)
        Xlsxrb.modify(mod_io) do |wb|
          updated_info = wb.sheets[1].update_cell("B1", value: "Success-#{i}")
          wb.with(sheets: [wb.sheets[0], updated_info])
        end

        mod_wb = Xlsxrb.read(mod_io)
        mod_val = mod_wb.sheets[1].load.rows[0].cells[1].value

        {
          success: true,
          mem_worker_id: mem_worker_id,
          stream_meta: stream_meta,
          mod_val: mod_val
        }
      rescue StandardError => e
        {
          success: false,
          error_class: e.class.name,
          error_message: e.message
        }
      end
    end

    results = ractors.map(&:value)
    results.each_with_index do |res, idx|
      assert_true res[:success], "Ractor #{idx} failed: #{res[:error_class]}: #{res[:error_message]}"
      assert_equal idx, res[:mem_worker_id]
      assert_equal "Ractor-#{idx}", res[:stream_meta]
      assert_equal "Success-#{idx}", res[:mod_val]
    end
  end
end
