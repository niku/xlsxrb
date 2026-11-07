# frozen_string_literal: true

require "test_helper"

class ApiSymmetryTest < Test::Unit::TestCase
  # Non-authoring or builder-lifecycle methods that are intentionally specific to WorksheetBuilder
  BUILDER_ONLY_METHODS = %i[
    build
    defined_names
    name
    styles
  ].freeze

  # Workbook-level delegators or stream-lifecycle methods intentionally on WorksheetProxy
  PROXY_ONLY_METHODS = %i[
    active_sheet
    active_tab
    app_property
    cleanup!
    core_property
    custom_property
    default_font
    first_sheet
    internal_sheet_setup
    properties
    protect_workbook
    set_default_font
    sheet_names
    show_sheet_tabs
    tab_ratio
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
      defined_name
    ]

    common_methods.each do |method_name|
      builder_params = Xlsxrb::WorksheetBuilder.instance_method(method_name).parameters
      stream_params = Xlsxrb::StreamWriter.instance_method(method_name).parameters

      if workbook_scoped_methods.include?(method_name)
        # Ensure all keyword parameters supported by WorksheetBuilder are also supported by StreamWriter
        builder_keys = builder_params.filter_map { |type, name| name if %i[key keyreq].include?(type) }
        stream_keys = stream_params.filter_map { |type, name| name if %i[key keyreq].include?(type) }
        missing_keys = builder_keys - stream_keys
        assert_empty missing_keys,
                     "Keyword parameters on WorksheetBuilder##{method_name} missing from StreamWriter##{method_name}: #{missing_keys.inspect}"
        next
      end

      # Ensure required, optional, and keyword parameters match
      builder_sig = builder_params.select { |param| %i[req opt key keyreq].include?(param.first) }
      stream_sig = stream_params.select { |param| %i[req opt key keyreq].include?(param.first) }

      assert_equal builder_sig, stream_sig,
                   "Parameter signature mismatch for method :#{method_name} (Builder: #{builder_sig.inspect}, Stream: #{stream_sig.inspect})"
    end

    # Ensure WorksheetProxy methods that define explicit parameters match WorksheetBuilder
    common_proxy_methods = (Xlsxrb::WorksheetBuilder.public_instance_methods(false).reject { |m| m.to_s.include?("__RBS_TEST_") } &
                            Xlsxrb::StreamWriter::WorksheetProxy.public_instance_methods(false).reject { |m| m.to_s.include?("__RBS_TEST_") })
    common_proxy_methods.each do |method_name|
      proxy_params = Xlsxrb::StreamWriter::WorksheetProxy.instance_method(method_name).parameters
      # Skip generic `...` delegators ([[:rest, :*], [:keyrest, :**], [:block, :&]])
      next if proxy_params.any? { |type, _| type == :rest }

      builder_params = Xlsxrb::WorksheetBuilder.instance_method(method_name).parameters
      builder_sig = builder_params.select { |param| %i[req opt key keyreq].include?(param.first) }
      proxy_sig = proxy_params.select { |param| %i[req opt key keyreq].include?(param.first) }

      assert_equal builder_sig, proxy_sig,
                   "WorksheetProxy##{method_name} signature mismatch with WorksheetBuilder##{method_name}"
    end
  end

  def test_workbook_builder_and_stream_writer_symmetry
    workbook_only_methods = %i[[] build]
    wb_methods = Xlsxrb::WorkbookBuilder.public_instance_methods(false).reject { |m| m.to_s.include?("__RBS_TEST_") } - workbook_only_methods
    sw_methods = Xlsxrb::StreamWriter.public_instance_methods(false).reject { |m| m.to_s.include?("__RBS_TEST_") }

    missing_in_stream = wb_methods - sw_methods
    assert_empty missing_in_stream,
                 "Workbook-level methods in WorkbookBuilder missing from StreamWriter: #{missing_in_stream.inspect}"

    wb_methods.each do |method_name|
      wb_params = Xlsxrb::WorkbookBuilder.instance_method(method_name).parameters
      sw_params = Xlsxrb::StreamWriter.instance_method(method_name).parameters

      wb_req_count = wb_params.count { |type, _| type == :req }
      sw_req_count = sw_params.count { |type, _| type == :req }
      assert_equal wb_req_count, sw_req_count,
                   "Required parameter count mismatch on WorkbookBuilder##{method_name} vs StreamWriter##{method_name}"

      wb_keys = wb_params.filter_map { |type, name| name if %i[key keyreq].include?(type) }.sort
      sw_keys = sw_params.filter_map { |type, name| name if %i[key keyreq].include?(type) }.sort
      assert_equal wb_keys, sw_keys,
                   "Keyword parameter mismatch on WorkbookBuilder##{method_name} vs StreamWriter##{method_name}"
    end
  end

  def test_defined_name_inmemory_and_streaming_xml_parity
    author_defined_names = lambda do |w|
      w.sheet("Master's Data") do |s|
        s.row(["Category"])
        s.row(["Alpha"])
        s.defined_name("GLOBAL_RANGE", range: "A2:A10")
        s.defined_name("COORDS_RANGE", col: 0, rows: 2..5)
        s.defined_name("LOCAL_RANGE", "A2:A10", local: true)
      end
      # Workbook-level defined name after sheet block must remain global (no localSheetId leak)
      w.defined_name("WORKBOOK_CONST", "0.08")
    end

    mem_doc = Xlsxrb.build(&author_defined_names)
    mem_bytes = Xlsxrb.write(mem_doc)

    stream_io = StringIO.new
    Xlsxrb.write(stream_io, &author_defined_names)
    stream_bytes = stream_io.string.b

    mem_wb_xml = Xlsxrb::Ooxml::Reader.new(mem_bytes).raw_entry("xl/workbook.xml")
    stream_wb_xml = Xlsxrb::Ooxml::Reader.new(stream_bytes).raw_entry("xl/workbook.xml")

    mem_dn = mem_wb_xml[%r{<definedNames>.*</definedNames>}m]
    stream_dn = stream_wb_xml[%r{<definedNames>.*</definedNames>}m]

    assert_not_nil mem_dn
    assert_equal mem_dn, stream_dn,
                 "In-memory and streaming definedNames XML must be identical"
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
      tab_color
      ignored_errors
      phonetic_properties
      code_name
      fit_to_page?
      full_calc_on_load?
      sheet_calc_properties
      sheet_calc_pr
      background_image
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
