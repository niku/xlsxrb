# frozen_string_literal: true

require "test_helper"
require "rbconfig"
require "tmpdir"
require_relative "../support/excel_rules"

# Checks written workbooks against rules that Excel enforces beyond the ECMA-376 schemas
# (see test/support/excel_rules.rb).
class ExcelRulesTest < Test::Unit::TestCase
  GALLERY = File.expand_path("../../examples/visual", __dir__)

  # Runs every gallery example in a subprocess, through StreamWriter or through Xlsxrb.build,
  # and writes the workbooks to out_dir.
  GALLERY_RUNNER = <<~RUBY
    require "xlsxrb"
    gallery, out_dir, mode = ARGV.shift(3)
    if mode == "build"
      Xlsxrb.singleton_class.prepend(Module.new do
        define_method(:write) do |target = Xlsxrb::OMITTED_TARGET, workbook = nil, **opts, &blk|
          return super(target, workbook, **opts, &blk) unless blk && workbook.nil?

          super(target, Xlsxrb.build(&blk), **opts.except(:shared_strings))
        end
      end)
    end
    Dir[File.join(gallery, "*.rb")].sort.each do |example|
      ARGV.replace([File.join(out_dir, File.basename(example, ".rb") + ".xlsx")])
      load example, true
    end
  RUBY

  def assert_no_violations(data, label)
    violations = ExcelRules.check(data)
    assert_empty(violations.map(&:to_s), label)
  end

  def base_workbook
    Xlsxrb.write(Xlsxrb.build do |b|
      b.style(:bold, bold: true)
      b.sheet("Data") do |s|
        s.row(%w[Region Quarter Sales], cell_styles: %i[bold bold bold])
        s.row(["East", "Q1", 10])
        s.row(["West", "Q2", 20])
        s.table("A1:C3", name: "Sales", columns: %w[Region Quarter Sales])
        s.comment("A2", "note")
        s.pivot_table("Data!A1:C3", field_names: %w[Region Quarter Sales], row_fields: ["Region"], column_fields: ["Quarter"],
                                    data_fields: [{ field: "Sales" }], items: { 0 => %w[East West], 1 => %w[Q1 Q2] }, ref: "E1")
      end
    end)
  end

  # Returns base_workbook with the given parts replaced, or with a block applied to their XML.
  def patched(changes)
    parts = Xlsxrb::Ooxml::ZipReader.new(StringIO.new(base_workbook)).read_all
    changes.each { |name, change| parts[name] = change.respond_to?(:call) ? change.call(parts.fetch(name)) : change }
    io = StringIO.new
    Xlsxrb::Ooxml::ZipWriter.open(io) { |zip| parts.each { |name, xml| zip.add_entry(name, xml) } }
    io.string
  end

  def rules(data) = ExcelRules.check(data).map(&:rule).uniq

  test "the base workbook follows every rule" do
    assert_no_violations(base_workbook, "base workbook")
  end

  test "a row or column pivot field without items is reported" do
    data = patched("xl/pivotTables/pivotTable1.xml" => ->(xml) { xml.sub(%r{<pivotField axis="axisRow"([^>]*)>.*?</pivotField>}m, '<pivotField axis="axisRow"\1/>') })
    assert_equal([:pivot_axis_items], rules(data))
  end

  test "a default subtotal item on a field with defaultSubtotal=0 is reported" do
    data = patched("xl/pivotTables/pivotTable1.xml" => ->(xml) { xml.sub('<pivotField axis="axisRow" showAll="0"', '<pivotField axis="axisRow" showAll="0" defaultSubtotal="0"') })
    assert_equal([:pivot_default_subtotal], rules(data))
  end

  test "pivot items past the cache's shared items are reported" do
    data = patched("xl/pivotTables/pivotTable1.xml" => ->(xml) { xml.sub('<item x="1"/>', '<item x="5"/>') })
    assert_equal([:pivot_items], rules(data))
  end

  test "a table style that is neither built in nor defined is reported" do
    data = patched("xl/tables/table1.xml" => ->(xml) { xml.sub(/<tableStyleInfo name="[^"]*"/, '<tableStyleInfo name="{name: &quot;TableStyleMedium2&quot;}"') })
    assert_equal([:style_name], rules(data))
  end

  test "a style index past its list is reported" do
    data = patched("xl/worksheets/sheet1.xml" => ->(xml) { xml.sub(/<c r="A2"([^>]*)>/, '<c r="A2" s="99"\1>') })
    assert_equal([:style_index], rules(data))
  end

  test "a comment without a VML note is reported" do
    data = patched("xl/drawings/vmlDrawing1.vml" => ->(xml) { xml.sub(%r{<x:Row>1</x:Row>}, "<x:Row>7</x:Row>") })
    assert_equal([:comment_notes], rules(data))
  end

  test "a relationship to a missing part and a part without a content type are reported" do
    data = patched("xl/_rels/workbook.xml.rels" => ->(xml) { xml.sub("</Relationships>", '<Relationship Id="rId99" Type="x" Target="missing.xml"/></Relationships>') },
                   "xl/extra.bin" => "x")
    assert_equal(%i[relationship_target content_type], rules(data))
  end

  test "cell text and sheet names over Excel's limits in UTF-16 code units are reported" do
    emoji = [0x1F600].pack("U")
    within = patched("xl/workbook.xml" => ->(xml) { xml.sub('name="Data"', %(name="#{emoji * 15}")) })
    assert_equal([], rules(within).grep(:text_length))

    long_name = patched("xl/workbook.xml" => ->(xml) { xml.sub('name="Data"', %(name="#{emoji * 16}")) })
    assert_include(ExcelRules.check(long_name).map(&:to_s), "text_length xl/workbook.xml: a sheet name has 32 UTF-16 code units, over 31")

    io = StringIO.new
    Xlsxrb.write(io, shared_strings: false, strict_excel_mode: false) { |w| w.sheet("S") { |s| s.row([emoji * 16_384, "x" * 32_767]) } }
    assert_equal(["text_length xl/worksheets/sheet1.xml: an inline string has 32768 UTF-16 code units, over 32767"],
                 ExcelRules.check(io.string).select { |v| v.rule == :text_length }.map(&:to_s))
    io = StringIO.new
    Xlsxrb.write(io, strict_excel_mode: false) { |w| w.sheet("S") { |s| s.row(["x" * 32_768]) } }
    assert_equal(["text_length xl/sharedStrings.xml: shared string 0 has 32768 UTF-16 code units, over 32767"],
                 ExcelRules.check(io.string).select { |v| v.rule == :text_length }.map(&:to_s))
    # Escaped control characters count as the one character they stand for.
    io = StringIO.new
    Xlsxrb.write(io) { |w| w.sheet("S") { |s| s.row([("x" * 32_766) + 1.chr]) } }
    assert_equal([], ExcelRules.check(io.string).select { |v| v.rule == :text_length })
  end

  test "a part stored twice is reported" do
    parts = Xlsxrb::Ooxml::ZipReader.new(StringIO.new(base_workbook)).read_all
    io = StringIO.new
    Xlsxrb::Ooxml::ZipWriter.open(io) do |zip|
      parts.each { |name, xml| zip.add_entry(name, xml) }
      zip.add_entry("xl/styles.xml", parts["xl/styles.xml"])
    end
    assert_equal([:duplicate_entry], rules(io.string))
  end

  test "gallery workbooks follow the rules when streamed, built, and loaded and written again" do
    Dir.mktmpdir do |dir|
      %w[stream build].each do |mode|
        out_dir = File.join(dir, mode)
        Dir.mkdir(out_dir)
        _, err, status = Open3.capture3(RbConfig.ruby, "-I", File.expand_path("../../lib", __dir__), "-e", GALLERY_RUNNER, GALLERY, out_dir, mode)
        assert_true(status.success?, "#{mode} gallery run failed: #{err.lines.first(5).join}")

        files = Dir[File.join(out_dir, "*.xlsx")]
        assert_equal(Dir[File.join(GALLERY, "*.rb")].size, files.size, "#{mode} gallery outputs")
        files.each do |file|
          data = File.binread(file)
          assert_no_violations(data, "#{mode} #{File.basename(file)}")
          assert_no_violations(Xlsxrb.write(Xlsxrb.load(StringIO.new(data))), "#{mode} #{File.basename(file)} loaded and written")
        end
      end
    end
  end

  test "charts across multiple sheets are numbered sequentially without duplicate zip entries" do
    build_proc = lambda do |w|
      w.sheet("SheetA") do |s|
        s.row(%w[Cat Val])
        s.row(["A", 10])
        s.chart(type: :bar, title: "ChartA", from_column: 3, from_row: 0, to_column: 8, to_row: 10,
                series: [{ name: "S1", categories: "SheetA!$A$2:$A$2", values: "SheetA!$B$2:$B$2" }])
      end
      w.sheet("SheetB") do |s|
        s.row(%w[Cat Val])
        s.row(["B", 20])
        s.chart(type: :line, title: "ChartB", from_column: 3, from_row: 0, to_column: 8, to_row: 10,
                series: [{ name: "S2", categories: "SheetB!$A$2:$A$2", values: "SheetB!$B$2:$B$2" }])
      end
    end

    [
      ["stream", Xlsxrb.write(&build_proc)],
      ["build", Xlsxrb.write(Xlsxrb.build(&build_proc))]
    ].each do |mode, data|
      assert_no_violations(data, "#{mode} multi-sheet charts")
      entry_names = Xlsxrb::Ooxml::ZipReader.new(StringIO.new(data)).entry_names
      assert_include(entry_names, "xl/charts/chart1.xml", "#{mode}: missing chart1.xml")
      assert_include(entry_names, "xl/charts/chart2.xml", "#{mode}: missing chart2.xml")

      wb = Xlsxrb.load(StringIO.new(data))
      assert_equal("ChartA", wb.sheet(0).charts.first.title)
      assert_equal("ChartB", wb.sheet(1).charts.first.title)
    end
  end

  test "mixed image and chart on a single sheet preserve distinct content types and chart numbers" do
    minimal_png = [
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

    build_proc = lambda do |w|
      w.sheet("Mixed") do |s|
        s.row(%w[Cat Val])
        s.row(["A", 10])
        s.image(minimal_png, ext: "png", from_column: 0, from_row: 3, to_column: 2, to_row: 6)
        s.chart(type: :bar, title: "MixedChart", from_column: 3, from_row: 0, to_column: 8, to_row: 10,
                series: [{ name: "S1", categories: "Mixed!$A$2:$A$2", values: "Mixed!$B$2:$B$2" }])
      end
    end

    [
      ["stream", Xlsxrb.write(&build_proc)],
      ["build", Xlsxrb.write(Xlsxrb.build(&build_proc))]
    ].each do |mode, data|
      assert_no_violations(data, "#{mode} mixed image and chart")
      entry_names = Xlsxrb::Ooxml::ZipReader.new(StringIO.new(data)).entry_names
      assert_include(entry_names, "xl/media/image1.png", "#{mode}: missing image1.png")
      assert_include(entry_names, "xl/charts/chart1.xml", "#{mode}: missing chart1.xml")

      wb = Xlsxrb.load(StringIO.new(data))
      assert_equal(1, wb.sheet(0).images.size)
      assert_equal(1, wb.sheet(0).charts.size)
      assert_equal("MixedChart", wb.sheet(0).charts.first.title)
    end
  end
end
