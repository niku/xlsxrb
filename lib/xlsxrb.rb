# frozen_string_literal: true

require_relative "xlsxrb/version"
require_relative "xlsxrb/zip_generator"
require_relative "xlsxrb/writer"
require_relative "xlsxrb/reader"

module Xlsxrb
  class Error < StandardError; end

  # Represents a formula with an optional cached value.
  Formula = Data.define(:expression, :cached_value)
end
