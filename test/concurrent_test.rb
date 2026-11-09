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
    omit_ractor_spawning_before_ruby4

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

  def test_element_constructor_ractor_safety
    omit "Ractors are not supported in this Ruby version" unless defined?(Ractor)
    omit "RBS::Test runtime hooks are incompatible with Ractor isolation" if defined?(RBS::Test)
    omit_ractor_spawning_before_ruby4

    ractors = 4.times.map do |idx|
      Ractor.new(idx) do |i|
        col = Xlsxrb::Elements::Column.new(index: 1, width: 12.5)
        row = Xlsxrb::Elements::Row.new(index: i + 1, height: 20.0)
        img = Xlsxrb::Elements::Image.new(filename: "test.png")
        meta = Xlsxrb::Elements::CoreProperties.new(creator: "Test #{i}")
        [col.width, row.index, img.filename, meta.creator]
      end
    end

    results = ractors.map(&:value)
    results.each_with_index do |res, idx|
      assert_equal [12.5, idx + 1, "test.png", "Test #{idx}"], res
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
    omit_ractor_spawning_before_ruby4

    worker_count = 4
    ractors = worker_count.times.map do |idx|
      Ractor.new(idx) do |i|
        # 1. Write workbook with multiple data types and styles
        xlsx_bin = Xlsxrb.write do |wb|
          wb.sheet("Data") do |s|
            s.row(["Worker", i, 123.45, Date.new(2026, 10, 23), { formula: "B1*2", value: i * 2 }])
            s.style("A1", bold: true, font_color: "0000FF")
          end
          wb.sheet("Info") do |s|
            s.row(["Metadata", "Ractor-#{i}"])
          end
        end

        # 2. In-Memory read
        mem_wb = Xlsxrb.load(xlsx_bin)
        mem_sheet = mem_wb.sheets[0]
        mem_worker_id = mem_sheet.rows[0].cells[1].value

        # 3. Streaming read
        stream_meta = nil
        Xlsxrb.read(xlsx_bin) do |sheet|
          next unless sheet.name == "Info"

          sheet.each do |row|
            stream_meta = row.values[1]
          end
        end

        # 4. Streaming transform
        mod_io = StringIO.new
        Xlsxrb.write(mod_io) do |w|
          Xlsxrb.read(xlsx_bin) do |sheet|
            w.sheet(sheet.name) do |ws|
              sheet.each_row do |row|
                vals = row.values
                vals[1] = "Success-#{i}" if sheet.name == "Info" && vals[0] == "Metadata"
                ws.row(vals)
              end
            end
          end
        end

        mod_wb = Xlsxrb.load(mod_io.string)
        mod_val = mod_wb.sheets[1].rows[0].cells[1].value

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

  private

  # Ruby 3.4's experimental Ractor implementation can deadlock when Ractors run alongside
  # the parallel test threads (the process hangs in ractor_take at exit), so Ractor
  # isolation is tested on Ruby 4.0 and later only.
  def omit_ractor_spawning_before_ruby4
    omit "Spawning Ractors is tested on Ruby 4.0 and later" if Gem::Version.new(RUBY_VERSION) < Gem::Version.new("4.0")
  end
end
