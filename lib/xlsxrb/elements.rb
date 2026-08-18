# frozen_string_literal: true

# rbs_inline: enabled

require_relative "elements/types"
require_relative "elements/cell"
require_relative "elements/row"
require_relative "elements/column"
require_relative "elements/coordinate_access"
require_relative "elements/worksheet"
require_relative "elements/workbook"

module Xlsxrb
  # High-level domain model layer.
  # Provides immutable Data classes representing Excel concepts.
  module Elements
  end
end
