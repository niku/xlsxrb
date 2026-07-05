# frozen_string_literal: true

require "bundler/gem_tasks"
require "rake/testtask"
require "etc"
require "fileutils"
require "open3"
require "tmpdir"

desc "Build the Open XML SDK runner"
task :build_sdk_runner do
  sh "dotnet build vendor/sdk_runner/sdk_runner.csproj -c Release"
end

def reader_fixture_dir
  File.expand_path("test/fixtures/reader_generated", __dir__)
end

def sdk_scenario_dir
  File.expand_path("test/fixtures/sdk_scenarios", __dir__)
end

def sdk_runner_dll
  File.expand_path("vendor/sdk_runner/bin/Release/net8.0/sdk_runner.dll", __dir__)
end

def reader_fixture_specs
  Dir.glob(File.join(sdk_scenario_dir, "reader_*_generated_by_sdk.cs")).map do |scenario_path|
    scenario_name = File.basename(scenario_path, ".cs")
    [scenario_name, File.join(reader_fixture_dir, "#{scenario_name}.xlsx")]
  end
end

def reader_fixture_workers
  Integer(ENV.fetch("READER_FIXTURE_WORKERS", Etc.nprocessors))
rescue ArgumentError
  Etc.nprocessors
end

desc "Ensure SDK-generated reader fixtures exist"
task ensure_reader_fixtures: :build_sdk_runner do
  missing_specs = reader_fixture_specs.reject { |_scenario_name, fixture_path| File.exist?(fixture_path) }
  next if missing_specs.empty?

  FileUtils.mkdir_p(reader_fixture_dir)

  queue = Queue.new
  missing_specs.each { |spec| queue << spec }
  failures = Queue.new

  worker_count = [reader_fixture_workers, missing_specs.size].min
  threads = Array.new(worker_count) do
    Thread.new do
      loop do
        scenario_name, fixture_path = queue.pop(true)
        scenario_path = File.join(sdk_scenario_dir, "#{scenario_name}.cs")
        FileUtils.touch(fixture_path)

        stdout, stderr, status = Open3.capture3(
          "dotnet", sdk_runner_dll, scenario_path, fixture_path
        )

        next if status.success?

        FileUtils.rm_f(fixture_path)
        message = stderr.to_s.strip.empty? ? stdout : stderr
        failures << "Failed to generate reader fixture #{scenario_name}: #{message}"
      rescue ThreadError
        break
      end
    end
  end

  threads.each(&:join)
  next if failures.empty?

  raise failures.pop
end

Rake::TestTask.new(:test) do |t|
  t.libs << "test"
  t.libs << "lib"
  t.test_files = FileList["test/**/*_test.rb"]
  workers = ENV.fetch("TEST_WORKERS", Etc.nprocessors)
  t.options = "--parallel --n-workers=#{workers}"
end

task test: %i[build_sdk_runner ensure_reader_fixtures]

namespace :test do
  Rake::TestTask.new(:unit) do |t|
    t.libs << "test"
    t.libs << "lib"
    t.test_files = FileList["test/xlsxrb/**/*_test.rb", "test/*_test.rb"]
    workers = ENV.fetch("TEST_WORKERS", Etc.nprocessors)
    t.options = "--parallel --n-workers=#{workers}"
  end

  Rake::TestTask.new(:contract) do |t|
    t.libs << "test"
    t.libs << "lib"
    t.test_files = FileList["test/contract/**/*_test.rb"]
    workers = ENV.fetch("TEST_WORKERS", Etc.nprocessors)
    t.options = "--parallel --n-workers=#{workers}"
  end

  Rake::TestTask.new(:e2e) do |t|
    t.libs << "test"
    t.libs << "lib"
    t.test_files = FileList["test/e2e/**/*_test.rb"]
    workers = ENV.fetch("TEST_WORKERS", Etc.nprocessors)
    t.options = "--parallel --n-workers=#{workers}"
  end

  Rake::TestTask.new(:visual) do |t|
    t.libs << "test"
    t.libs << "lib"
    t.test_files = FileList["test/visual/**/*_test.rb"]
    workers = ENV.fetch("TEST_WORKERS", Etc.nprocessors)
    t.options = "--parallel --n-workers=#{workers}"
  end

  namespace :fixtures do
    namespace :reader do
      desc "Generate reader fixture XLSX files from SDK scenarios"
      task generate: :ensure_reader_fixtures

      desc "Remove generated reader fixture XLSX files"
      task :clean do
        FileUtils.rm_rf(reader_fixture_dir)
      end
    end
  end
end

task "test:e2e" => %i[build_sdk_runner ensure_reader_fixtures]

require "rubocop/rake_task"

RuboCop::RakeTask.new

task default: %i[test rubocop]

namespace :visual do
  desc "Generate docs/visual/ README.md explanation gallery"
  task :gallery do
    require_relative "test/visual/support/gallery_generator"
    Xlsxrb::Visual::GalleryGenerator.generate
  end

  desc "Generate / update VRT baselines in test/visual/baselines/"
  task :baseline do
    require_relative "test/visual/support/renderer"
    require "securerandom"
    examples = Dir.glob(File.expand_path("examples/visual/*.rb", __dir__))
    baselines_dir = File.expand_path("test/visual/baselines", __dir__)

    FileUtils.rm_rf(baselines_dir)
    FileUtils.mkdir_p(baselines_dir)

    examples.each do |example_path|
      name = File.basename(example_path, ".rb")
      puts "Generating baseline for #{name}..."

      example_tmp_dir = File.join(Dir.tmpdir, "xlsxrb_baselines_build_#{name}_#{SecureRandom.hex(4)}")
      FileUtils.mkdir_p(example_tmp_dir)

      xlsx_path = File.join(example_tmp_dir, "#{name}.xlsx")
      system("ruby", "-Ilib", example_path, xlsx_path)

      dest_dir = File.join(baselines_dir, name)
      begin
        Xlsxrb::Visual::Renderer.render(xlsx_path, dest_dir)
      ensure
        FileUtils.rm_rf(example_tmp_dir)
      end
    end
    puts "Baselines generated successfully."
  end
end

desc "Generate RDoc documentation including Visual Gallery"
task :doc do
  FileUtils.rm_rf("doc")
  sh "bundle exec rdoc --title 'xlsxrb Documentation' --main README.md README.md \"docs/visual/VisualGallery.md\" lib/"

  # Copy visual gallery images and files so they are available in RDoc output
  FileUtils.mkdir_p("doc/test/visual/baselines")
  FileUtils.cp_r("test/visual/baselines", "doc/test/visual")
  FileUtils.mkdir_p("doc/test/visual/support/illustrations")
  FileUtils.cp_r(Dir.glob("test/visual/support/illustrations/*.png"), "doc/test/visual/support/illustrations")

  FileUtils.mkdir_p("doc/docs/visual/files")
  FileUtils.cp_r(Dir.glob("docs/visual/files/*"), "doc/docs/visual/files")
  puts "RDoc documentation generated successfully at doc/"
end
