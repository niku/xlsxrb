# frozen_string_literal: true

$LOAD_PATH.unshift File.expand_path("../lib", __dir__)
require "xlsxrb"

require "test-unit"
require "open3"
require "fileutils"

class OpenXmlSdkScenarioRunner
  def self.sdk_runner_dll
    File.expand_path("../vendor/sdk_runner/bin/Release/net8.0/sdk_runner.dll", File.dirname(__FILE__))
  end

  def self.scenario_dir
    File.expand_path("fixtures/sdk_scenarios", __dir__)
  end

  def self.reader_fixture_dir
    File.expand_path("fixtures/reader_generated", __dir__)
  end

  def self.reader_fixture_scenario?(scenario_name)
    scenario_name.start_with?("reader_") && scenario_name.end_with?("_generated_by_sdk")
  end

  def self.reader_fixture_path(scenario_name)
    File.join(reader_fixture_dir, "#{scenario_name}.xlsx")
  end

  def self.scenario_path(scenario_name)
    File.join(scenario_dir, "#{scenario_name}.cs")
  end

  def self.ensure_reader_fixture!(scenario_name)
    return nil unless reader_fixture_scenario?(scenario_name)

    fixture_path = reader_fixture_path(scenario_name)
    return fixture_path if File.exist?(fixture_path)

    FileUtils.mkdir_p(reader_fixture_dir)
    FileUtils.touch(fixture_path)

    result = run_single_mode(scenario_path(scenario_name), fixture_path)
    raise "Failed to generate reader fixture for #{scenario_name}: #{result[:stderr]}" unless result[:success]

    fixture_path
  end

  def self.copy_reader_fixture(scenario_name, xlsx_path)
    fixture_path = ensure_reader_fixture!(scenario_name)
    return false unless fixture_path

    FileUtils.cp(fixture_path, xlsx_path)
    true
  rescue StandardError
    false
  end

  def self.run_single_scenario(scenario_path, xlsx_path)
    command = ["dotnet", sdk_runner_dll, scenario_path, xlsx_path]
    stdout, stderr, status = Open3.capture3(*command)
    {
      success: status.success?,
      stdout:,
      stderr:,
      command: command.join(" ")
    }
  end

  def self.run_single_mode(scenario_path, xlsx_path)
    run_single_scenario(scenario_path, xlsx_path)
  end
end
