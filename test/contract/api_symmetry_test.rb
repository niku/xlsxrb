# frozen_string_literal: true

require "test_helper"

class ApiSymmetryTest < Test::Unit::TestCase
  # Non-authoring or builder-lifecycle methods that are intentionally specific to WorksheetBuilder
  BUILDER_ONLY_METHODS = %i[
    build
    defined_names
    styles
  ].freeze

  # Workbook-level delegators or stream-lifecycle methods intentionally on WorksheetProxy
  PROXY_ONLY_METHODS = %i[
    active_sheet
    active_sheet=
    active_tab
    active_tab=
    app_property
    cleanup!
    core_property
    custom_property
    default_font
    first_sheet
    first_sheet=
    internal_sheet_setup
    properties
    protect_workbook
    workbook_protection
    set_default_font
    sheet_names
    show_sheet_tabs
    show_sheet_tabs=
    tab_ratio
    tab_ratio=
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
    normalize_sig = lambda do |sig|
      sig.map do |type, name|
        if %i[req opt].include?(type)
          name.to_s.start_with?("__") ? [type, :arg] : [type, name]
        else
          [type, name]
        end
      end
    end

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

      assert_equal normalize_sig.call(builder_sig), normalize_sig.call(stream_sig),
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

      assert_equal normalize_sig.call(builder_sig), normalize_sig.call(proxy_sig),
                   "WorksheetProxy##{method_name} signature mismatch with WorksheetBuilder##{method_name}"
    end
  end

  def test_workbook_builder_and_stream_writer_symmetry
    wb_lifecycle_methods = %i[build sheets]
    sw_lifecycle_methods = %i[
      <<
      cleanup!
      close
      compression_level
      current_sheet
      internal_sheet_setup
      shared_strings
      shared_strings?
      start_sheet_entry
    ]
    # Worksheet-delegated methods on StreamWriter (authoring operations that operate on active sheet)
    ws_delegated_methods = Xlsxrb::WorksheetBuilder.public_instance_methods(false).reject { |m| m.to_s.include?("__RBS_TEST_") } - %i[
      build
      name
      styles
      escape_formulas
      escape_formulas=
      escape_formulas?
    ]
    sw_exempt = sw_lifecycle_methods + ws_delegated_methods

    wb_methods = Xlsxrb::WorkbookBuilder.public_instance_methods(false).reject { |m| m.to_s.include?("__RBS_TEST_") } - wb_lifecycle_methods
    sw_methods = Xlsxrb::StreamWriter.public_instance_methods(false).reject { |m| m.to_s.include?("__RBS_TEST_") }

    missing_in_stream = wb_methods - sw_methods
    assert_empty missing_in_stream,
                 "Workbook-level methods in WorkbookBuilder missing from StreamWriter: #{missing_in_stream.inspect}"

    missing_in_builder = (sw_methods - sw_exempt) - wb_methods
    assert_empty missing_in_builder,
                 "Workbook-level methods in StreamWriter missing from WorkbookBuilder: #{missing_in_builder.inspect}"

    (wb_methods - %i[hyperlink comment]).each do |method_name|
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
        s.defined_name("COORDS_RANGE", columns: 0, rows: 2..5)
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

  def test_sheet_builder_explicit_keyword_arguments_symmetry
    # 1. In-memory build
    mem_wb = Xlsxrb.build do |b|
      b.sheet("Sheet1", state: :visible, tab_color: "FF0000", code_name: "Code1") do |s|
        s.row([1, 2])
      end
      b.sheet("Sheet2", state: :hidden, tab_color: "00FF00", code_name: "Code2") do |s|
        s.row([3, 4])
      end
    end
    assert_equal :visible, mem_wb.sheet(0).state
    assert_equal :hidden, mem_wb.sheet(1).state

    # 2. Streaming write
    io = StringIO.new
    Xlsxrb.write(io) do |w|
      w.sheet("Sheet1", state: :visible, tab_color: "FF0000", code_name: "Code1") do |s|
        s.row([1, 2])
      end
      w.sheet("Sheet2", state: :hidden, tab_color: "00FF00", code_name: "Code2") do |s|
        s.row([3, 4])
      end
    end
    stream_wb = Xlsxrb.load(StringIO.new(io.string))
    assert_equal :visible, stream_wb.sheet(0).state
    assert_equal :hidden, stream_wb.sheet(1).state
  end

  def test_builder_deprecation_warnings_symmetry
    # 1. WorkbookBuilder sheet hidden: and defined_name column:/row:
    err_wb = capture_stderr do
      Xlsxrb.build do |b|
        b.sheet("VisibleSheet")
        b.sheet("HiddenSheet", hidden: true)
        b.defined_name("COL_NAME", range: "A1", column: 0)
        b.defined_name("ROW_NAME", range: "A1", row: 1)
      end
    end
    assert_match(/\[Xlsxrb DEPRECATION\] sheet keyword hidden: is deprecated; use state: :hidden instead\./, err_wb)
    assert_match(/\[Xlsxrb DEPRECATION\] defined_name keyword column: is deprecated; use columns: instead\./, err_wb)
    assert_match(/\[Xlsxrb DEPRECATION\] defined_name keyword row: is deprecated; use rows: instead\./, err_wb)

    # 2. WorksheetBuilder defined_name column:/row:, outline_properties show_symbols:, fit_to_page?
    ws_builder = Xlsxrb::WorksheetBuilder.new("WS")
    err_ws = capture_stderr do
      ws_builder.defined_name("WS_COL", range: "A1", column: 0)
      ws_builder.defined_name("WS_ROW", range: "A1", row: 1)
      ws_builder.outline_properties(show_symbols: true)
      ws_builder.fit_to_page?
    end
    assert_match(/\[Xlsxrb DEPRECATION\] defined_name keyword column: is deprecated; use columns: instead\./, err_ws)
    assert_match(/\[Xlsxrb DEPRECATION\] defined_name keyword row: is deprecated; use rows: instead\./, err_ws)
    assert_match(/\[Xlsxrb DEPRECATION\] outline_properties keyword show_symbols: is deprecated; use show_outline_symbols: instead\./, err_ws)
    assert_match(/\[Xlsxrb DEPRECATION\] fit_to_page\? is deprecated; use fit_to_page instead\./, err_ws)

    # 3. StreamWriter sheet hidden:, defined_name column:/row:, outline_properties show_symbols:, fit_to_page?
    io = StringIO.new
    err_sw = capture_stderr do
      Xlsxrb.write(io) do |w|
        w.sheet("StreamVisible")
        w.sheet("StreamHidden", hidden: true)
        w.defined_name("SW_COL", range: "A1", column: 0)
        w.defined_name("SW_ROW", range: "A1", row: 1)
        w.outline_properties(show_symbols: true)
        w.fit_to_page?
      end
    end
    assert_match(/\[Xlsxrb DEPRECATION\] sheet keyword hidden: is deprecated; use state: :hidden instead\./, err_sw)
    assert_match(/\[Xlsxrb DEPRECATION\] defined_name keyword column: is deprecated; use columns: instead\./, err_sw)
    assert_match(/\[Xlsxrb DEPRECATION\] defined_name keyword row: is deprecated; use rows: instead\./, err_sw)
    assert_match(/\[Xlsxrb DEPRECATION\] outline_properties keyword show_symbols: is deprecated; use show_outline_symbols: instead\./, err_sw)
    assert_match(/\[Xlsxrb DEPRECATION\] fit_to_page\? is deprecated; use fit_to_page instead\./, err_sw)

    # 4. WorksheetProxy defined_name column:/row:, outline_properties show_symbols:, fit_to_page?
    io2 = StringIO.new
    err_proxy = capture_stderr do
      Xlsxrb.write(io2) do |w|
        w.sheet("ProxySheet") do |s|
          s.defined_name("PROXY_COL", range: "A1", column: 0)
          s.defined_name("PROXY_ROW", range: "A1", row: 1)
          s.outline_properties(show_symbols: true)
          s.fit_to_page?
        end
      end
    end
    assert_match(/\[Xlsxrb DEPRECATION\] defined_name keyword column: is deprecated; use columns: instead\./, err_proxy)
    assert_match(/\[Xlsxrb DEPRECATION\] defined_name keyword row: is deprecated; use rows: instead\./, err_proxy)
    assert_match(/\[Xlsxrb DEPRECATION\] outline_properties keyword show_symbols: is deprecated; use show_outline_symbols: instead\./, err_proxy)
    assert_match(/\[Xlsxrb DEPRECATION\] fit_to_page\? is deprecated; use fit_to_page instead\./, err_proxy)

    # 5. DataValidationBuilder #[] and #[]=
    dv = Xlsxrb::DataValidationBuilder.new("A1")
    err_dv = capture_stderr do
      dv[:type] = :list
      assert_equal(:list, dv[:type])
    end
    assert_match(/\[Xlsxrb DEPRECATION\] DataValidationBuilder#\[\]= is deprecated; use dedicated property accessors instead\./, err_dv)
    assert_match(/\[Xlsxrb DEPRECATION\] DataValidationBuilder#\[\] is deprecated; use dedicated property accessors instead\./, err_dv)

    # 6. ConditionalFormatBuilder #[], #[]=, #formulas, #formulas=
    cf = Xlsxrb::ConditionalFormatBuilder.new("A1:B2")
    err_cf = capture_stderr do
      cf[:type] = :cell_is
      assert_equal(:cell_is, cf[:type])
      cf.formulas = ["A1>0"]
      assert_equal(["A1>0"], cf.formulas)
    end
    assert_match(/\[Xlsxrb DEPRECATION\] ConditionalFormatBuilder#\[\]= is deprecated; use dedicated property accessors instead\./, err_cf)
    assert_match(/\[Xlsxrb DEPRECATION\] ConditionalFormatBuilder#\[\] is deprecated; use dedicated property accessors instead\./, err_cf)
    assert_match(/\[Xlsxrb DEPRECATION\] ConditionalFormatBuilder#formulas= is deprecated; use #formula= instead\./, err_cf)
    assert_match(/\[Xlsxrb DEPRECATION\] ConditionalFormatBuilder#formulas is deprecated; use #formula instead\./, err_cf)
  end
end
