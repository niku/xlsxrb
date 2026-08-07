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
end
