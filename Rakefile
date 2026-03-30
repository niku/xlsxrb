# frozen_string_literal: true

require "bundler/gem_tasks"
require "rake/testtask"
require "etc"
require "fileutils"
require "open3"

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

task :ensure_reader_fixtures => :build_sdk_runner do
  missing_specs = reader_fixture_specs.reject { |_scenario_name, fixture_path| File.exist?(fixture_path) }
  if missing_specs.empty?
    next
  end

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

require "rubocop/rake_task"

RuboCop::RakeTask.new

task default: %i[test rubocop]
