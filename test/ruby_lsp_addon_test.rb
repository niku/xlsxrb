# frozen_string_literal: true

require "test_helper"
require "ruby_lsp/internal"
require "ruby_lsp/addon"
require "ruby_lsp/xlsxrb/addon"

class RubyLspAddonTest < Test::Unit::TestCase
  setup do
    @addon = RubyLsp::Xlsxrb::Addon.new
    @global_state = RubyLsp::GlobalState.new
    @addon.activate(@global_state, Thread::Queue.new)
    RubyLsp::Addon.addons << @addon
  end

  teardown do
    RubyLsp::Addon.addons.delete(@addon)
  end

  def test_addon_metadata
    assert_equal("xlsxrb", @addon.name)
    assert_equal(Xlsxrb::VERSION, @addon.version)
  end

  def test_auto_discovery_of_addon
    global_state = RubyLsp::GlobalState.new
    RubyLsp::Addon.load_addons(global_state, Thread::Queue.new)
    discovered = RubyLsp::Addon.addons.find { |a| a.name == "xlsxrb" }
    assert_not_nil(discovered)
    assert_kind_of(RubyLsp::Xlsxrb::Addon, discovered)
  end

  # 1. Xlsxrb.write -> StreamWriter
  def test_stream_writer_completions_in_generate_block
    source = <<~RUBY
      Xlsxrb.write("test.xlsx") do |wb|
        wb.sheet
      end
    RUBY

    results = request_completions(source, line: 1, character: 7)
    labels = results.map(&:label)

    assert_includes(labels, "sheet")
    assert_includes(labels, "workbook_property")
    assert_includes(labels, "style")
    assert_includes(labels, "protect_workbook")

    sheet_item = results.find { |i| i.label == "sheet" }
    assert_not_nil(sheet_item)
    assert_includes(sheet_item.detail, "Xlsxrb::StreamWriter")
  end

  # 2. Xlsxrb.build -> WorkbookBuilder
  def test_workbook_builder_completions_in_build_block
    source = <<~RUBY
      Xlsxrb.build do |builder|
        builder.sheet
      end
    RUBY

    results = request_completions(source, line: 1, character: 12)
    labels = results.map(&:label)

    assert_includes(labels, "sheet")
    assert_includes(labels, "build")
    assert_includes(labels, "style")

    build_item = results.find { |i| i.label == "build" }
    assert_not_nil(build_item)
    assert_includes(build_item.detail, "Xlsxrb::WorkbookBuilder")
  end

  # 3. Xlsxrb.read -> StreamSheet
  def test_stream_sheet_completions_in_read_block
    source = <<~RUBY
      Xlsxrb.read("test.xlsx") do |sheet|
        sheet.each_row
      end
    RUBY

    results = request_completions(source, line: 1, character: 10)
    labels = results.map(&:label)

    assert_includes(labels, "each_row")
    assert_includes(labels, "each")
    assert_includes(labels, "name")

    each_row_item = results.find { |i| i.label == "each_row" }
    assert_includes(each_row_item.detail, "Xlsxrb::StreamSheet")
  end

  # 4. Xlsxrb.modify -> Elements::Workbook
  def test_workbook_element_completions_in_modify_block
    source = <<~RUBY
      Xlsxrb.modify("test.xlsx") do |doc|
        doc.sheet
      end
    RUBY

    results = request_completions(source, line: 1, character: 8)
    labels = results.map(&:label)

    assert_includes(labels, "sheet")
    assert_includes(labels, "sheet_names")
    assert_includes(labels, "update_sheet")
    assert_includes(labels, "save")

    sheet_item = results.find { |i| i.label == "sheet" }
    assert_includes(sheet_item.detail, "Xlsxrb::Elements::Workbook")
  end

  # 5. wb.sheet -> WorksheetProxy
  def test_worksheet_proxy_completions_in_sheet_block
    source = <<~RUBY
      Xlsxrb.write("test.xlsx") do |wb|
        wb.sheet("Sheet1") do |s|
          s.row
        end
      end
    RUBY

    results = request_completions(source, line: 2, character: 8)
    labels = results.map(&:label)

    assert_includes(labels, "row")
    assert_includes(labels, "column")
    assert_includes(labels, "merge")
    assert_includes(labels, "freeze_pane")
    assert_includes(labels, "hyperlink")
    assert_includes(labels, "auto_filter")
    assert_includes(labels, "conditional_format")
    assert_includes(labels, "table")
    assert_includes(labels, "pivot_table")

    row_item = results.find { |i| i.label == "row" }
    assert_not_nil(row_item)
    assert_includes(row_item.detail, "Xlsxrb::StreamWriter::WorksheetProxy")
  end

  # 6. s.chart -> ChartBuilder
  def test_chart_builder_completions_in_chart_block
    source = <<~RUBY
      s.chart(:bar) do |cb|
        cb.title
      end
    RUBY

    results = request_completions(source, line: 1, character: 8)
    labels = results.map(&:label)

    assert_includes(labels, "title")
    assert_includes(labels, "categories")
    assert_includes(labels, "series")
    assert_includes(labels, "legend")

    title_item = results.find { |i| i.label == "title" }
    assert_includes(title_item.detail, "Xlsxrb::ChartBuilder")
  end

  # 7. wb.style -> StyleBuilder
  def test_style_builder_completions_in_style_block
    source = <<~RUBY
      wb.style(:header) do |sb|
        sb.font
      end
    RUBY

    results = request_completions(source, line: 1, character: 8)
    labels = results.map(&:label)

    assert_includes(labels, "font")
    assert_includes(labels, "fill")
    assert_includes(labels, "border")
    assert_includes(labels, "number_format")

    font_item = results.find { |i| i.label == "font" }
    assert_includes(font_item.detail, "Xlsxrb::StyleBuilder")
  end

  # 8. sheet.each_row -> Elements::Row
  def test_row_completions_in_each_row_block
    source = <<~RUBY
      sheet.each_row do |r|
        r.to_a
      end
    RUBY

    results = request_completions(source, line: 1, character: 7)
    labels = results.map(&:label)

    assert_includes(labels, "to_a")
    assert_includes(labels, "cell_at")
    assert_includes(labels, "values")
    assert_includes(labels, "index")

    to_a_item = results.find { |i| i.label == "to_a" }
    assert_includes(to_a_item.detail, "Xlsxrb::Elements::Row")
  end

  # 9. sheet.each_cell -> Elements::Cell
  def test_cell_completions_in_each_cell_block
    source = <<~RUBY
      sheet.each_cell do |c|
        c.value
      end
    RUBY

    results = request_completions(source, line: 1, character: 7)
    labels = results.map(&:label)

    assert_includes(labels, "value")
    assert_includes(labels, "ref")
    assert_includes(labels, "to_i")
    assert_includes(labels, "to_date")

    value_item = results.find { |i| i.label == "value" }
    assert_includes(value_item.detail, "Xlsxrb::Elements::Cell")
  end

  # 10. workbook.each -> Elements::Worksheet
  def test_worksheet_completions_in_workbook_each_block
    source = <<~RUBY
      workbook.each do |sheet|
        sheet.row_at
      end
    RUBY

    results = request_completions(source, line: 1, character: 10)
    labels = results.map(&:label)

    assert_includes(labels, "row_at")
    assert_includes(labels, "cell_value")
    assert_includes(labels, "cells")

    row_at_item = results.find { |i| i.label == "row_at" }
    assert_includes(row_at_item.detail, "Xlsxrb::Elements::Worksheet")
  end

  # 11. Fallback heuristic by variable name
  def test_fallback_heuristics_by_variable_names
    # Row heuristic
    res_r = request_completions("row.to_a", line: 0, character: 6)
    assert_includes(res_r.map(&:label), "to_a")

    # Cell heuristic
    res_c = request_completions("cell.value", line: 0, character: 7)
    assert_includes(res_c.map(&:label), "value")

    # Chart heuristic
    res_cb = request_completions("chart.title", line: 0, character: 8)
    assert_includes(res_cb.map(&:label), "title")

    # Style heuristic
    res_sb = request_completions("style.font", line: 0, character: 8)
    assert_includes(res_sb.map(&:label), "font")
  end

  # 12. SortText Priority Ordering
  def test_sort_text_priority_ordering
    source = <<~RUBY
      Xlsxrb.write("test.xlsx") do |wb|
        wb.sheet("S1") do |s|
          s.row
        end
      end
    RUBY

    wb_results = request_completions(source, line: 1, character: 7)
    assert_equal("000_sheet", wb_results.first.sort_text)
    assert_equal("sheet", wb_results.first.label)

    s_results = request_completions(source, line: 2, character: 8)
    assert_equal("000_row", s_results.first.sort_text)
    assert_equal("row", s_results.first.label)
    assert_equal("001_column", s_results[1].sort_text)
    assert_equal("column", s_results[1].label)
  end

  # 13. Negative test
  def test_unrelated_variable_receives_no_xlsxrb_completions
    source = <<~RUBY
      User.find_each do |user|
        user.name
      end
    RUBY

    results = request_completions(source, line: 1, character: 9)
    labels = results.map(&:label)

    assert_empty(labels)
  end

  private

  def request_completions(source, line:, character:)
    doc = RubyLsp::RubyDocument.new(source: source, version: 1, uri: URI("file:///test.rb"), global_state: @global_state)
    dispatcher = Prism::Dispatcher.new
    params = {
      textDocument: { uri: "file:///test.rb" },
      position: { line: line, character: character }
    }

    req = RubyLsp::Requests::Completion.new(
      doc,
      @global_state,
      params,
      RubyLsp::SorbetLevel.new("ignore"),
      dispatcher
    )

    req.perform
  end
end
