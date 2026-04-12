# frozen_string_literal: true

require "test_helper"
require "tempfile"

# Tests for all newly promoted Facade DSL features.
# Each feature is tested in both streaming (Xlsxrb.generate) and in-memory (Xlsxrb.build) modes.
class FacadeFeaturesTest < Test::Unit::TestCase
  # =====================================================
  # Hyperlinks
  # =====================================================

  test "add_hyperlink options form in build API" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("Links") do |s|
        s.add_row(["Click here", "Internal"])
        s.add_hyperlink("A1", "https://example.com", display: "Example", tooltip: "Visit")
        s.add_hyperlink("B1", location: "Sheet1!A1")
      end
    end

    tmp = Tempfile.new(["facade_hyperlink_build", ".xlsx"])
    Xlsxrb.write(tmp.path, workbook)

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    links = reader.hyperlinks
    assert(links.key?("A1"), "Hyperlink on A1 should exist")
    link = links["A1"]
    url = link.is_a?(Hash) ? link[:url] : link
    assert_equal("https://example.com", url)
  ensure
    tmp&.close!
  end

  test "add_hyperlink options form in generate API" do
    tmp = Tempfile.new(["facade_hyperlink_stream", ".xlsx"])
    Xlsxrb.generate(tmp.path) do |w|
      w.add_sheet("Links") do |s|
        s.add_row(["Click here"])
        s.add_hyperlink("A1", "https://example.com", display: "Example")
      end
    end

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    links = reader.hyperlinks
    assert(links.key?("A1"), "Hyperlink on A1 should exist")
    link = links["A1"]
    url = link.is_a?(Hash) ? link[:url] : link
    assert_equal("https://example.com", url)
  ensure
    tmp&.close!
  end

  test "add_hyperlink keyword args in build API" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("Links") do |s|
        s.add_row(["Link"])
        s.add_hyperlink("A1", "https://example.com", display: "Example")
      end
    end

    tmp = Tempfile.new(["facade_hyperlink_block", ".xlsx"])
    Xlsxrb.write(tmp.path, workbook)

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    links = reader.hyperlinks
    assert(links.key?("A1"))
  ensure
    tmp&.close!
  end

  # =====================================================
  # Auto Filter
  # =====================================================

  test "set_auto_filter in build API" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("Data") do |s|
        s.add_row(["Name", "Score"])
        s.add_row(["Alice", 95])
        s.add_row(["Bob", 87])
        s.set_auto_filter("A1:B3")
      end
    end

    tmp = Tempfile.new(["facade_autofilter_build", ".xlsx"])
    Xlsxrb.write(tmp.path, workbook)

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    af = reader.auto_filter
    assert_equal("A1:B3", af)
  ensure
    tmp&.close!
  end

  test "set_auto_filter in generate API" do
    tmp = Tempfile.new(["facade_autofilter_stream", ".xlsx"])
    Xlsxrb.generate(tmp.path) do |w|
      w.add_sheet("Data") do |s|
        s.add_row(["Name", "Score"])
        s.add_row(["Alice", 95])
        s.set_auto_filter("A1:B2")
      end
    end

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    af = reader.auto_filter
    assert_equal("A1:B2", af)
  ensure
    tmp&.close!
  end

  # =====================================================
  # Data Validation
  # =====================================================

  test "add_data_validation options form in build API" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("DV") do |s|
        s.add_row(["Value"])
        s.add_data_validation("A2:A100", type: :whole, operator: :between,
                              formula1: "1", formula2: "100",
                              show_error_message: true, error: "Enter 1-100")
      end
    end

    tmp = Tempfile.new(["facade_dv_build", ".xlsx"])
    Xlsxrb.write(tmp.path, workbook)

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    dvs = reader.data_validations
    assert_equal(1, dvs.size)
    assert_equal("A2:A100", dvs[0][:sqref])
  ensure
    tmp&.close!
  end

  test "add_data_validation keyword args in generate API" do
    tmp = Tempfile.new(["facade_dv_stream", ".xlsx"])
    Xlsxrb.generate(tmp.path) do |w|
      w.add_sheet("DV") do |s|
        s.add_row(["Value"])
        s.add_data_validation("A2:A100", type: :list, formula1: '"A,B,C"', show_error_message: true)
      end
    end

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    dvs = reader.data_validations
    assert_equal(1, dvs.size)
    assert_equal("A2:A100", dvs[0][:sqref])
  ensure
    tmp&.close!
  end

  # =====================================================
  # Conditional Formatting
  # =====================================================

  test "add_conditional_format in build API" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("CF") do |s|
        s.add_row([10, 20, 30])
        s.add_conditional_format("A1:C1", type: :cell_is, operator: :greaterThan, formula: "15", priority: 1)
      end
    end

    tmp = Tempfile.new(["facade_cf_build", ".xlsx"])
    Xlsxrb.write(tmp.path, workbook)

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    cfs = reader.conditional_formats
    assert_equal(1, cfs.size)
    assert_equal("A1:C1", cfs[0][:sqref])
  ensure
    tmp&.close!
  end

  test "add_conditional_format in generate API" do
    tmp = Tempfile.new(["facade_cf_stream", ".xlsx"])
    Xlsxrb.generate(tmp.path) do |w|
      w.add_sheet("CF") do |s|
        s.add_row([10, 20, 30])
        s.add_conditional_format("A1:C1", type: :cell_is, operator: :greaterThan, formula: "15", priority: 1)
      end
    end

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    cfs = reader.conditional_formats
    assert_equal(1, cfs.size)
  ensure
    tmp&.close!
  end

  # =====================================================
  # Tables
  # =====================================================

  test "add_table in build API" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("Tables") do |s|
        s.add_row(["Name", "Score"])
        s.add_row(["Alice", 95])
        s.add_row(["Bob", 87])
        s.add_table("A1:B3", columns: ["Name", "Score"], name: "ScoreTable")
      end
    end

    tmp = Tempfile.new(["facade_table_build", ".xlsx"])
    Xlsxrb.write(tmp.path, workbook)

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    tables = reader.tables
    assert_equal(1, tables.size)
    assert_equal("ScoreTable", tables[0][:name])
    assert_equal("A1:B3", tables[0][:ref])
  ensure
    tmp&.close!
  end

  test "add_table in generate API" do
    tmp = Tempfile.new(["facade_table_stream", ".xlsx"])
    Xlsxrb.generate(tmp.path) do |w|
      w.add_sheet("Tables") do |s|
        s.add_row(["Name", "Score"])
        s.add_row(["Alice", 95])
        s.add_table("A1:B2", columns: ["Name", "Score"], name: "MyTable")
      end
    end

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    tables = reader.tables
    assert_equal(1, tables.size)
    assert_equal("MyTable", tables[0][:name])
  ensure
    tmp&.close!
  end

  # =====================================================
  # Comments
  # =====================================================

  test "add_comment in build API" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("Comments") do |s|
        s.add_row(["Value"])
        s.add_comment("A1", "This is a comment", author: "Test")
      end
    end

    tmp = Tempfile.new(["facade_comment_build", ".xlsx"])
    Xlsxrb.write(tmp.path, workbook)

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    comments = reader.comments
    assert_equal(1, comments.size)
    assert_equal("A1", comments[0][:ref])
    assert_equal("This is a comment", comments[0][:text])
  ensure
    tmp&.close!
  end

  test "add_comment in generate API" do
    tmp = Tempfile.new(["facade_comment_stream", ".xlsx"])
    Xlsxrb.generate(tmp.path) do |w|
      w.add_sheet("Comments") do |s|
        s.add_row(["Value"])
        s.add_comment("A1", "Stream comment", author: "Author")
      end
    end

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    comments = reader.comments
    assert_equal(1, comments.size)
    assert_equal("A1", comments[0][:ref])
  ensure
    tmp&.close!
  end

  # =====================================================
  # Merge Cells
  # =====================================================

  test "merge_cells in build API" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("Merge") do |s|
        s.add_row(["Merged title", nil, nil])
        s.add_row([1, 2, 3])
        s.merge_cells("A1:C1")
      end
    end

    tmp = Tempfile.new(["facade_merge_build", ".xlsx"])
    Xlsxrb.write(tmp.path, workbook)

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    merged = reader.merged_cells
    assert_equal(["A1:C1"], merged)
  ensure
    tmp&.close!
  end

  test "merge_cells in generate API" do
    tmp = Tempfile.new(["facade_merge_stream", ".xlsx"])
    Xlsxrb.generate(tmp.path) do |w|
      w.add_sheet("Merge") do |s|
        s.add_row(["Merged"])
        s.merge_cells("A1:B1")
      end
    end

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    merged = reader.merged_cells
    assert_equal(["A1:B1"], merged)
  ensure
    tmp&.close!
  end

  # =====================================================
  # Freeze Panes
  # =====================================================

  test "set_freeze_pane in build API" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("Frozen") do |s|
        s.add_row(["Header"])
        s.add_row([1])
        s.set_freeze_pane(row: 1, col: 0)
      end
    end

    tmp = Tempfile.new(["facade_freeze_build", ".xlsx"])
    Xlsxrb.write(tmp.path, workbook)

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    pane = reader.freeze_pane
    assert_not_nil(pane, "Freeze pane should exist")
    # Reader returns row/col for freeze pane
    assert_equal(1, pane[:row])
  ensure
    tmp&.close!
  end

  test "set_freeze_pane in generate API" do
    tmp = Tempfile.new(["facade_freeze_stream", ".xlsx"])
    Xlsxrb.generate(tmp.path) do |w|
      w.add_sheet("Frozen") do |s|
        s.add_row(["Header"])
        s.add_row([1])
        s.set_freeze_pane(row: 1, col: 1)
      end
    end

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    pane = reader.freeze_pane
    assert_not_nil(pane)
  ensure
    tmp&.close!
  end

  # =====================================================
  # Page Margins
  # =====================================================

  test "set_page_margins in build API" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("Margins") do |s|
        s.add_row(["Data"])
        s.set_page_margins(left: 1.0, right: 1.0, top: 1.5, bottom: 1.5)
      end
    end

    tmp = Tempfile.new(["facade_margins_build", ".xlsx"])
    Xlsxrb.write(tmp.path, workbook)

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    margins = reader.page_margins
    assert_not_nil(margins)
    assert_in_delta(1.0, margins[:left])
    assert_in_delta(1.5, margins[:top])
  ensure
    tmp&.close!
  end

  test "set_page_margins in generate API" do
    tmp = Tempfile.new(["facade_margins_stream", ".xlsx"])
    Xlsxrb.generate(tmp.path) do |w|
      w.add_sheet("Margins") do |s|
        s.add_row(["Data"])
        s.set_page_margins(left: 0.5, right: 0.5)
      end
    end

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    margins = reader.page_margins
    assert_not_nil(margins)
    assert_in_delta(0.5, margins[:left])
  ensure
    tmp&.close!
  end

  # =====================================================
  # Page Setup
  # =====================================================

  test "set_page_setup in build API" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("Setup") do |s|
        s.add_row(["Data"])
        s.set_page_setup(orientation: :landscape, paper_size: 9)
      end
    end

    tmp = Tempfile.new(["facade_pagesetup_build", ".xlsx"])
    Xlsxrb.write(tmp.path, workbook)

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    ps = reader.page_setup
    assert_not_nil(ps)
    assert_equal("landscape", ps[:orientation])
  ensure
    tmp&.close!
  end

  test "set_page_setup in generate API" do
    tmp = Tempfile.new(["facade_pagesetup_stream", ".xlsx"])
    Xlsxrb.generate(tmp.path) do |w|
      w.add_sheet("Setup") do |s|
        s.add_row(["Data"])
        s.set_page_setup(orientation: :portrait, scale: 80)
      end
    end

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    ps = reader.page_setup
    assert_not_nil(ps)
    assert_equal("portrait", ps[:orientation])
  ensure
    tmp&.close!
  end

  # =====================================================
  # Header / Footer
  # =====================================================

  test "set_header_footer in build API" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("HF") do |s|
        s.add_row(["Data"])
        s.set_header_footer(odd_header: "&CReport Title", odd_footer: "&CPage &P")
      end
    end

    tmp = Tempfile.new(["facade_hf_build", ".xlsx"])
    Xlsxrb.write(tmp.path, workbook)

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    hf = reader.header_footer
    assert_not_nil(hf)
    assert_equal("&CReport Title", hf[:odd_header])
    assert_equal("&CPage &P", hf[:odd_footer])
  ensure
    tmp&.close!
  end

  test "set_header_footer in generate API" do
    tmp = Tempfile.new(["facade_hf_stream", ".xlsx"])
    Xlsxrb.generate(tmp.path) do |w|
      w.add_sheet("HF") do |s|
        s.add_row(["Data"])
        s.set_header_footer(odd_header: "&LLeft Header")
      end
    end

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    hf = reader.header_footer
    assert_not_nil(hf)
    assert_equal("&LLeft Header", hf[:odd_header])
  ensure
    tmp&.close!
  end

  # =====================================================
  # Print Options
  # =====================================================

  test "set_print_option in build API" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("PO") do |s|
        s.add_row(["Data"])
        s.set_print_option(:grid_lines, true)
        s.set_print_option(:horizontal_centered, true)
      end
    end

    tmp = Tempfile.new(["facade_po_build", ".xlsx"])
    Xlsxrb.write(tmp.path, workbook)

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    opts = reader.print_options
    assert(opts[:grid_lines], "Grid lines should be set")
  ensure
    tmp&.close!
  end

  # =====================================================
  # Sheet Protection
  # =====================================================

  test "set_sheet_protection in build API" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("Protected") do |s|
        s.add_row(["Data"])
        s.set_sheet_protection(sheet: true, objects: true)
      end
    end

    tmp = Tempfile.new(["facade_prot_build", ".xlsx"])
    Xlsxrb.write(tmp.path, workbook)

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    prot = reader.sheet_protection
    assert_not_nil(prot, "Sheet protection should exist")
  ensure
    tmp&.close!
  end

  test "set_sheet_protection keyword args in generate API" do
    tmp = Tempfile.new(["facade_prot_stream", ".xlsx"])
    Xlsxrb.generate(tmp.path) do |w|
      w.add_sheet("Protected") do |s|
        s.add_row(["Data"])
        s.set_sheet_protection(sheet: true, scenarios: true)
      end
    end

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    prot = reader.sheet_protection
    assert_not_nil(prot)
  ensure
    tmp&.close!
  end

  # =====================================================
  # Row / Column Breaks
  # =====================================================

  test "add_row_break and add_col_break in build API" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("Breaks") do |s|
        10.times { |i| s.add_row(["Row #{i}"]) }
        s.add_row_break(5)
        s.add_col_break(3)
      end
    end

    tmp = Tempfile.new(["facade_breaks_build", ".xlsx"])
    Xlsxrb.write(tmp.path, workbook)

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    rbs = reader.row_breaks
    cbs = reader.col_breaks
    # Reader may return hashes or integers depending on implementation
    rb_ids = rbs.map { |b| b.is_a?(Hash) ? b[:id] : b }
    cb_ids = cbs.map { |b| b.is_a?(Hash) ? b[:id] : b }
    assert_include(rb_ids, 5)
    assert_include(cb_ids, 3)
  ensure
    tmp&.close!
  end

  # =====================================================
  # Defined Names / Print Area / Print Titles
  # =====================================================

  test "add_defined_name in build API" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("Data") do |s|
        s.add_row([100])
      end
      w.add_defined_name("MyRange", "'Data'!$A$1:$A$1", sheet: "Data")
    end

    tmp = Tempfile.new(["facade_defname_build", ".xlsx"])
    Xlsxrb.write(tmp.path, workbook)

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    names = reader.defined_names
    assert(names.any? { |dn| dn[:name] == "MyRange" }, "Defined name should exist")
  ensure
    tmp&.close!
  end

  test "add_defined_name in generate API" do
    tmp = Tempfile.new(["facade_defname_stream", ".xlsx"])
    Xlsxrb.generate(tmp.path) do |w|
      w.add_sheet("Data") do |s|
        s.add_row([100])
      end
      w.add_defined_name("Total", "'Data'!$A$1", sheet: "Data")
    end

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    names = reader.defined_names
    assert(names.any? { |dn| dn[:name] == "Total" })
  ensure
    tmp&.close!
  end

  test "set_print_area in build API" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("Data") do |s|
        s.add_row([1, 2, 3])
        s.add_row([4, 5, 6])
      end
      w.set_print_area("A1:C2", sheet: "Data")
    end

    tmp = Tempfile.new(["facade_printarea_build", ".xlsx"])
    Xlsxrb.write(tmp.path, workbook)

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    names = reader.defined_names
    pa = names.find { |dn| dn[:name] == "_xlnm.Print_Area" }
    assert_not_nil(pa, "Print area defined name should exist")
  ensure
    tmp&.close!
  end

  test "set_print_titles in generate API" do
    tmp = Tempfile.new(["facade_printtitles_stream", ".xlsx"])
    Xlsxrb.generate(tmp.path) do |w|
      w.add_sheet("Data") do |s|
        s.add_row(["Header1", "Header2"])
        s.add_row([1, 2])
      end
      w.set_print_titles(rows: "1:1", sheet: "Data")
    end

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    names = reader.defined_names
    pt = names.find { |dn| dn[:name] == "_xlnm.Print_Titles" }
    assert_not_nil(pt, "Print titles defined name should exist")
  ensure
    tmp&.close!
  end

  # =====================================================
  # Workbook Protection
  # =====================================================

  test "set_workbook_protection in build API" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("S") { |s| s.add_row(["Data"]) }
      w.set_workbook_protection(lock_structure: true)
    end

    tmp = Tempfile.new(["facade_wbprot_build", ".xlsx"])
    Xlsxrb.write(tmp.path, workbook)

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    prot = reader.workbook_protection
    assert_not_nil(prot, "Workbook protection should exist")
  ensure
    tmp&.close!
  end

  test "set_workbook_protection in generate API" do
    tmp = Tempfile.new(["facade_wbprot_stream", ".xlsx"])
    Xlsxrb.generate(tmp.path) do |w|
      w.add_sheet("S") { |s| s.add_row(["Data"]) }
      w.set_workbook_protection(lock_structure: true, lock_windows: true)
    end

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    prot = reader.workbook_protection
    assert_not_nil(prot)
  ensure
    tmp&.close!
  end

  # =====================================================
  # Document Properties
  # =====================================================

  test "set_core_property in build API" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("S") { |s| s.add_row(["Data"]) }
      w.set_core_property(:creator, "Test Author")
      w.set_core_property(:title, "Test Title")
    end

    tmp = Tempfile.new(["facade_coreprop_build", ".xlsx"])
    Xlsxrb.write(tmp.path, workbook)

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    props = reader.core_properties
    assert_equal("Test Author", props[:creator])
    assert_equal("Test Title", props[:title])
  ensure
    tmp&.close!
  end

  test "set_core_property in generate API" do
    tmp = Tempfile.new(["facade_coreprop_stream", ".xlsx"])
    Xlsxrb.generate(tmp.path) do |w|
      w.add_sheet("S") { |s| s.add_row(["Data"]) }
      w.set_core_property(:creator, "Stream Author")
    end

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    props = reader.core_properties
    assert_equal("Stream Author", props[:creator])
  ensure
    tmp&.close!
  end

  test "set_app_property in build API" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("S") { |s| s.add_row(["Data"]) }
      w.set_app_property(:company, "Acme Corp")
    end

    tmp = Tempfile.new(["facade_appprop_build", ".xlsx"])
    Xlsxrb.write(tmp.path, workbook)

    # Verify the app.xml file contains the property
    entries = Xlsxrb::Ooxml::ZipReader.open(tmp.path, &:read_all)
    app_xml = entries["docProps/app.xml"]
    assert_not_nil(app_xml, "docProps/app.xml should exist")
    assert_match(/Acme Corp/, app_xml, "Company should be in app.xml")
  ensure
    tmp&.close!
  end

  test "add_custom_property in build API" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("S") { |s| s.add_row(["Data"]) }
      w.add_custom_property("Department", "Engineering", type: :string)
    end

    tmp = Tempfile.new(["facade_custprop_build", ".xlsx"])
    Xlsxrb.write(tmp.path, workbook)

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    props = reader.custom_properties
    assert(props.any? { |p| p[:name] == "Department" && p[:value] == "Engineering" })
  ensure
    tmp&.close!
  end

  # =====================================================
  # Selection
  # =====================================================

  test "set_selection in build API" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("Sel") do |s|
        s.add_row(["A", "B"])
        s.set_selection("B1")
      end
    end

    tmp = Tempfile.new(["facade_sel_build", ".xlsx"])
    Xlsxrb.write(tmp.path, workbook)

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    sel = reader.selection
    assert_not_nil(sel)
    assert_equal("B1", sel[:active_cell])
  ensure
    tmp&.close!
  end

  # =====================================================
  # Images
  # =====================================================

  # Minimal 1x1 white pixel PNG for image tests.
  MINIMAL_PNG = [
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
    0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
    0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
    0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
    0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
    0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC,
    0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
    0x44, 0xAE, 0x42, 0x60, 0x82
  ].pack("C*").freeze

  test "add_image in build API" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("Img") do |s|
        s.add_row(["With image"])
        s.add_image(MINIMAL_PNG, ext: "png", from_col: 0, from_row: 0)
      end
    end

    tmp = Tempfile.new(["facade_image_build", ".xlsx"])
    Xlsxrb.write(tmp.path, workbook)

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    images = reader.images
    assert_equal(1, images.size, "Image should be present")
  ensure
    tmp&.close!
  end

  test "add_image in generate API" do
    tmp = Tempfile.new(["facade_image_stream", ".xlsx"])
    Xlsxrb.generate(tmp.path) do |w|
      w.add_sheet("Img") do |s|
        s.add_row(["With image"])
        s.add_image(MINIMAL_PNG, ext: "png")
      end
    end

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    images = reader.images
    assert_equal(1, images.size)
  ensure
    tmp&.close!
  end

  # =====================================================
  # Shapes
  # =====================================================

  test "add_shape in build API" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("Shapes") do |s|
        s.add_row(["Data"])
        s.add_shape(preset: "rect", text: "Hello", from_col: 1, from_row: 1, to_col: 3, to_row: 4)
      end
    end

    tmp = Tempfile.new(["facade_shape_build", ".xlsx"])
    Xlsxrb.write(tmp.path, workbook)

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    shapes = reader.shapes
    assert_equal(1, shapes.size)
  ensure
    tmp&.close!
  end

  test "add_shape in generate API" do
    tmp = Tempfile.new(["facade_shape_stream", ".xlsx"])
    Xlsxrb.generate(tmp.path) do |w|
      w.add_sheet("Shapes") do |s|
        s.add_row(["Data"])
        s.add_shape(preset: "ellipse", text: "Circle")
      end
    end

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    shapes = reader.shapes
    assert_equal(1, shapes.size)
  ensure
    tmp&.close!
  end

  # =====================================================
  # Sheet View (zoom, gridlines)
  # =====================================================

  test "set_sheet_view in build API" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("View") do |s|
        s.add_row(["Data"])
        s.set_sheet_view(:zoom_scale, 150)
        s.set_sheet_view(:show_grid_lines, false)
      end
    end

    tmp = Tempfile.new(["facade_view_build", ".xlsx"])
    Xlsxrb.write(tmp.path, workbook)

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    sv = reader.sheet_view
    assert_not_nil(sv)
    assert_equal(150, sv[:zoom_scale])
  ensure
    tmp&.close!
  end

  # =====================================================
  # Sort State
  # =====================================================

  test "set_sort_state in build API" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("Sort") do |s|
        s.add_row(["Name", "Score"])
        s.add_row(["Alice", 95])
        s.add_row(["Bob", 87])
        s.set_auto_filter("A1:B3")
        s.set_sort_state("A2:B3", [{ ref: "B2:B3", descending: true }])
      end
    end

    tmp = Tempfile.new(["facade_sort_build", ".xlsx"])
    Xlsxrb.write(tmp.path, workbook)

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    ss = reader.sort_state
    assert_not_nil(ss, "Sort state should exist")
    assert_equal("A2:B3", ss[:ref])
  ensure
    tmp&.close!
  end

  # =====================================================
  # Combined features test
  # =====================================================

  test "multiple features on same sheet in build API" do
    workbook = Xlsxrb.build do |w|
      w.add_sheet("Combined") do |s|
        s.add_row(["Name", "Score", "Grade"])
        s.add_row(["Alice", 95, "A"])
        s.add_row(["Bob", 87, "B"])
        s.set_auto_filter("A1:C3")
        s.merge_cells("A1:A1")
        s.set_freeze_pane(row: 1)
        s.set_page_margins(left: 1.0, right: 1.0)
        s.set_page_setup(orientation: :landscape)
        s.add_hyperlink("A2", "https://example.com")
        s.add_data_validation("B2:B3", type: :whole, formula1: "0", formula2: "100")
      end
      w.set_core_property(:creator, "Combined Test")
    end

    tmp = Tempfile.new(["facade_combined_build", ".xlsx"])
    Xlsxrb.write(tmp.path, workbook)

    # Read back and verify everything
    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    assert_equal("A1:C3", reader.auto_filter)
    assert_equal(["A1:A1"], reader.merged_cells)
    assert_not_nil(reader.freeze_pane)
    assert_not_nil(reader.page_margins)
    assert_not_nil(reader.page_setup)
    assert(reader.hyperlinks.key?("A2"))
    assert_equal(1, reader.data_validations.size)
    assert_equal("Combined Test", reader.core_properties[:creator])
  ensure
    tmp&.close!
  end

  test "multiple features on same sheet in generate API" do
    tmp = Tempfile.new(["facade_combined_stream", ".xlsx"])
    Xlsxrb.generate(tmp.path) do |w|
      w.add_sheet("Combined") do |s|
        s.add_row(["Name", "Score"])
        s.add_row(["Alice", 95])
        s.set_auto_filter("A1:B2")
        s.merge_cells("A1:B1")
        s.set_freeze_pane(row: 1)
        s.set_page_margins(left: 0.5, right: 0.5)
        s.set_header_footer(odd_header: "&CTest")
        s.add_conditional_format("B2", type: :cell_is, operator: :greaterThan, formula: "90")
      end
      w.set_core_property(:title, "Combined Stream")
    end

    reader = Xlsxrb::Ooxml::Reader.new(tmp.path)
    assert_equal("A1:B2", reader.auto_filter)
    assert_equal(["A1:B1"], reader.merged_cells)
    assert_not_nil(reader.freeze_pane)
    assert_not_nil(reader.page_margins)
    assert_not_nil(reader.header_footer)
    assert_equal(1, reader.conditional_formats.size)
    assert_equal("Combined Stream", reader.core_properties[:title])
  ensure
    tmp&.close!
  end
end
