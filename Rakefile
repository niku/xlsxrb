# frozen_string_literal: true

require "bundler/gem_tasks"
require "rake/testtask"
require "etc"

task :build_sdk_runner do
  sh "dotnet build vendor/sdk_runner/sdk_runner.csproj -c Release"
end

Rake::TestTask.new(:test) do |t|
  t.libs << "test"
  t.libs << "lib"
  t.test_files = FileList["test/**/*_test.rb"]
  workers = ENV.fetch("TEST_WORKERS", Etc.nprocessors)
  t.options = "--parallel --n-workers=#{workers}"
end

task test: :build_sdk_runner

require "rubocop/rake_task"

RuboCop::RakeTask.new

task default: %i[test rubocop]
