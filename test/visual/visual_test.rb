# frozen_string_literal: true

require "test_helper"
require_relative "support/renderer"
require_relative "support/comparator"
require "fileutils"

class VisualTest < Test::Unit::TestCase
  BASELINES_DIR = File.expand_path("baselines", __dir__)
  OUTPUT_DIR = File.expand_path("output", __dir__)

  def assert_visual_match(name, fuzz: "5%", threshold: 100)
    example_path = File.expand_path("../../examples/visual/#{name}.rb", __dir__)
    raise "Example not found: #{example_path}" unless File.exist?(example_path)

    FileUtils.mkdir_p(OUTPUT_DIR)

    # 1. Generate XLSX
    xlsx_path = File.join(OUTPUT_DIR, "#{name}.xlsx")
    system("ruby", "-Ilib", example_path, xlsx_path)
    assert(File.exist?(xlsx_path), "Failed to generate XLSX for #{name}")

    # 2. Render to PNG
    candidate_dir = File.join(OUTPUT_DIR, name)
    FileUtils.rm_rf(candidate_dir)
    candidate_pngs = Xlsxrb::Visual::Renderer.render(xlsx_path, candidate_dir)

    # 3. Compare with baselines
    baseline_dir = File.join(BASELINES_DIR, name)
    assert(File.exist?(baseline_dir), "Baseline directory not found: #{baseline_dir}. Run `rake visual:baseline` first.")

    baseline_pngs = Dir.glob(File.join(baseline_dir, "page-*.png")).sort_by do |path|
      path.match(/page-(\d+)\.png/)[1].to_i
    end

    assert_equal(baseline_pngs.size, candidate_pngs.size, "Number of rendered pages mismatch for #{name}")

    baseline_pngs.each_with_index do |baseline_path, idx|
      candidate_path = candidate_pngs[idx]
      diff_path = File.join(candidate_dir, "diff-page-#{idx + 1}.png")

      diff_pixels = Xlsxrb::Visual::Comparator.compare(baseline_path, candidate_path, diff_path, fuzz: fuzz)

      assert(diff_pixels <= threshold, "Visual mismatch in #{name} page #{idx + 1}: #{diff_pixels} pixels differ (max allowed: #{threshold})")
    end
  end

  # Dynamically define visual match test cases for all examples
  Dir.glob(File.expand_path("../../examples/visual/*.rb", __dir__)).each do |example_path|
    name = File.basename(example_path, ".rb")
    test "#{name} visual match" do
      if name.start_with?("chart_")
        assert_visual_match(name, fuzz: "10%", threshold: 500)
      else
        assert_visual_match(name)
      end
    end
  end
end
