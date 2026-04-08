# frozen_string_literal: true

require_relative "ooxml/utils"
require_relative "ooxml/zip_reader"
require_relative "ooxml/zip_writer"
require_relative "ooxml/xml_parser"
require_relative "ooxml/xml_builder"
require_relative "ooxml/shared_strings_parser"
require_relative "ooxml/styles_parser"
require_relative "ooxml/workbook_parser"
require_relative "ooxml/worksheet_parser"
require_relative "ooxml/worksheet_writer"
require_relative "ooxml/workbook_writer"

module Xlsxrb
  # Low-level OOXML infrastructure layer.
  # Handles ZIP extraction, SAX XML parsing, and XML generation
  # in strict accordance with ECMA-376.
  module Ooxml
  end
end
