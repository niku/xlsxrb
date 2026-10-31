# frozen_string_literal: true

if ENV["COVERAGE"] == "1" || ENV["CI"] == "true"
  require "simplecov"
  require "simplecov_json_formatter"
  SimpleCov.formatter = SimpleCov::Formatter::JSONFormatter
  SimpleCov.start
end
