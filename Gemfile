# frozen_string_literal: true

source "https://rubygems.org"

# Runtime dependencies must be specified in xlsxrb.gemspec.
# End-user production environments only resolve gemspec dependencies; this Gemfile
# is strictly for development/testing and is never packaged or loaded by consumers.
gemspec

# Development and test dependencies are listed flatly in alphabetical order:
# - No `group`: Gemfile is dev-only; grouping does not affect consumers or CI.
# - No `require: false`: Bundler.require is not used; files are required explicitly.
# - No version constraints: Gemfile.lock pins versions, and Dependabot manages updates.
gem "benchmark-ips"
gem "bundler-audit"
gem "irb"
gem "memory_profiler"
gem "mutant"
gem "mutant-test-unit"
gem "nokogiri"
gem "pbt"
gem "rake"
gem "rbs-inline"
gem "rubocop"
gem "rubocop-rake"
gem "ruby-lsp"
gem "ruby_wasm"
gem "rubyzip"
gem "simplecov"
gem "simplecov_json_formatter"
gem "steep"
gem "test-unit"
gem "webrick"
