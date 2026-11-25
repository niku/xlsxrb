# frozen_string_literal: true

require "nokogiri"
require "stringio"

# Validates XML parts inside an XLSX package against the ECMA-376 XSD schemas
# stored under test/fixtures/xsd. Lazily loads and caches compiled schemas so
# test suites incur the schema compilation cost at most once per process.
module XsdValidator
  XSD_DIR = File.expand_path("../fixtures/xsd", __dir__)
  SCHEMA_MUTEX = Mutex.new
  SCHEMA_FILES = {
    sml: "sml.xsd",
    drawing: "dml-spreadsheetDrawing.xsd",
    chart: "dml-chart.xsd",
    app_props: "shared-documentPropertiesExtended.xsd",
    custom_props: "shared-documentPropertiesCustom.xsd"
  }.freeze

  module_function

  def schemas
    return @schemas if @schemas

    SCHEMA_MUTEX.synchronize do
      @schemas ||= SCHEMA_FILES.transform_values do |file|
        path = File.join(XSD_DIR, file)
        Nokogiri::XML::Schema.from_document(Nokogiri::XML(File.read(path), path))
      end
    end
  end

  def schema_key_for(entry_name)
    case entry_name
    when %r{\Axl/workbook\.xml\z},
         %r{\Axl/worksheets/sheet\d+\.xml\z},
         %r{\Axl/styles\.xml\z},
         %r{\Axl/sharedStrings\.xml\z},
         %r{\Axl/tables/table\d+\.xml\z},
         %r{\Axl/comments\d+\.xml\z},
         %r{\Axl/pivotTables/pivotTable\d+\.xml\z},
         %r{\Axl/pivotCache/pivotCacheDefinition\d+\.xml\z}
      :sml
    when %r{\Axl/drawings/drawing\d+\.xml\z}
      :drawing
    when %r{\Axl/charts/chart\d+\.xml\z}
      :chart
    when %r{\AdocProps/app\.xml\z}
      :app_props
    when %r{\AdocProps/custom\.xml\z}
      :custom_props
    end
  end

  # Returns an Array of XSD validation error strings for all known XML parts in +data+.
  def check(data)
    parts = Xlsxrb::Ooxml::ZipReader.new(StringIO.new(data)).read_all
    loaded_schemas = schemas
    errors = []

    parts.each do |name, xml_content|
      next unless name.end_with?(".xml")

      schema_key = schema_key_for(name)
      next unless schema_key

      doc = Nokogiri::XML(xml_content)
      loaded_schemas[schema_key].validate(doc).each do |err|
        errors << "#{name}: #{err}"
      end
    end

    errors
  end
end
