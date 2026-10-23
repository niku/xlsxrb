# frozen_string_literal: true

require "test_helper"

class ApiSymmetryTest < Test::Unit::TestCase
  # Non-authoring or builder-lifecycle methods that are intentionally specific to WorksheetBuilder
  BUILDER_ONLY_METHODS = %i[
    build
    name
    styles
  ].freeze

  # Workbook-level delegators or stream-lifecycle methods intentionally on WorksheetProxy
  PROXY_ONLY_METHODS = %i[
    app_property
    cleanup!
    core_property
    custom_property
    defined_name
    internal_sheet_setup
    properties
    protect_workbook
    workbook_property
  ].freeze

  def test_worksheet_builder_and_proxy_method_symmetry
    wb_methods = Xlsxrb::WorksheetBuilder.public_instance_methods(false).reject { |m| m.to_s.include?("__RBS_TEST_") }
    proxy_methods = Xlsxrb::StreamWriter::WorksheetProxy.public_instance_methods(false).reject { |m| m.to_s.include?("__RBS_TEST_") }

    # 1. Ensure all authoring methods in WorksheetBuilder exist in WorksheetProxy
    missing_in_proxy = (wb_methods - BUILDER_ONLY_METHODS) - proxy_methods
    assert_empty missing_in_proxy,
                 "Methods in WorksheetBuilder missing from StreamWriter::WorksheetProxy: #{missing_in_proxy.inspect}"

    # 2. Ensure all sheet authoring methods in WorksheetProxy exist in WorksheetBuilder
    missing_in_builder = (proxy_methods - PROXY_ONLY_METHODS) - wb_methods
    assert_empty missing_in_builder,
                 "Methods in StreamWriter::WorksheetProxy missing from WorksheetBuilder: #{missing_in_builder.inspect}"
  end

  def test_worksheet_builder_and_stream_writer_parameter_symmetry
    # Compare method signatures between WorksheetBuilder and StreamWriter for common authoring methods
    common_methods = (Xlsxrb::WorksheetBuilder.public_instance_methods(false).reject { |m| m.to_s.include?("__RBS_TEST_") } &
                      Xlsxrb::StreamWriter.public_instance_methods(false).reject { |m| m.to_s.include?("__RBS_TEST_") })

    # Methods that accept an optional :sheet keyword on StreamWriter because it operates at workbook scope
    workbook_scoped_methods = %i[
      print_area
      print_titles
      repeat_rows
      repeat_columns
    ]

    common_methods.each do |method_name|
      next if workbook_scoped_methods.include?(method_name)

      builder_params = Xlsxrb::WorksheetBuilder.instance_method(method_name).parameters
      stream_params = Xlsxrb::StreamWriter.instance_method(method_name).parameters

      # Ensure required parameters match exactly
      builder_req = builder_params.filter_map { |type, name| name if type == :req }
      stream_req = stream_params.filter_map { |type, name| name if type == :req }

      assert_equal builder_req, stream_req,
                   "Required parameter mismatch for method :#{method_name} (Builder: #{builder_req.inspect}, Stream: #{stream_req.inspect})"
    end
  end

  def test_sheet_reader_query_interface_symmetry
    # Expected common query methods on both in-memory Elements::Worksheet and StreamSheet
    expected_query_methods = %i[
      cell
      each
      name
      state
      hidden?
      visible?
      very_hidden?
    ]

    elem_methods = Xlsxrb::Elements::Worksheet.public_instance_methods(true)
    stream_methods = Xlsxrb::StreamSheet.public_instance_methods(true)

    expected_query_methods.each do |m|
      assert_include elem_methods, m, "Elements::Worksheet must implement #{m}"
      assert_include stream_methods, m, "StreamSheet must implement #{m}"
    end

    # Architectural invariant: StreamSheet intentionally does not support []
    # to avoid misleading users about O(1) random coordinate access on streaming sheets.
    # Users must explicitly call #load to get an Elements::Worksheet for [] indexing.
    assert_include elem_methods, :[], "Elements::Worksheet provides random coordinate access with []"
    assert_not_include stream_methods, :[], "StreamSheet intentionally does not define [] without explicit #load"
  end

  def test_roundtrip_lossless_contract_inmemory_and_streaming
    Tempfile.create(["contract_symmetry", ".xlsx"]) do |f|
      test_data = [
        ["Product", "Quantity", "Price", "Available"],
        ["Espresso Machine", 12, 599.95, true],
        ["Coffee Beans", 150, 18.50, false]
      ]

      # 1. Write via streaming
      Xlsxrb.write(f.path) do |w|
        w.sheet("Inventory") do |s|
          test_data.each { |r| s.row(r) }
        end
      end

      # 2. Read via in-memory
      wb_mem = Xlsxrb.read(f.path)
      ws_mem = wb_mem.sheets.first.load
      mem_rows = ws_mem.rows.map { |r| r.cells.map(&:value) }
      assert_equal test_data, mem_rows

      # 3. Read via streaming
      stream_rows = []
      Xlsxrb.read(f.path) do |sheet|
        sheet.each { |r| stream_rows << r.values }
      end
      assert_equal test_data, stream_rows
    end
  end
end
