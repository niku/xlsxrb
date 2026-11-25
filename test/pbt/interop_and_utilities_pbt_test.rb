# frozen_string_literal: true

require "test_helper"
require "pbt"
require "zip"
require "bigdecimal"
require "date"
require "time"
require "stringio"
require "tmpdir"
require_relative "../support/excel_rules"
require_relative "../support/xsd_validator"

class InteropAndUtilitiesPbtTest < Test::Unit::TestCase
  TINY_PNG = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
             .unpack1("m0").b.freeze

  # Minimal write-only sink without seek/pos/tell to force unseekable ZipWriter paths.
  class UnseekableSink
    attr_reader :buffer

    def initialize
      @buffer = +"".b
    end

    def write(chunk)
      s = chunk.b
      @buffer << s
      s.bytesize
    end

    def <<(chunk)
      @buffer << chunk.b
      self
    end

    def flush
      self
    end

    def binmode
      self
    end
  end

  # Duck-typed Time replacement simulating ActiveSupport::TimeWithZone.
  class FakeTimeWithZone
    attr_reader :year, :month, :day, :hour, :min, :sec, :subsec, :utc_offset

    def initialize(year, month, day, hour, min, sec, subsec = 0, utc_offset = 0)
      @year = year
      @month = month
      @day = day
      @hour = hour
      @min = min
      @sec = sec
      @subsec = subsec
      @utc_offset = utc_offset
    end

    def is_a?(klass)
      klass == Time || super
    end

    def to_time
      Time.new(@year, @month, @day, @hour, @min, @sec + @subsec, @utc_offset)
    end
  end

  def setup
    omit "Skip property-based tests under RBS runtime testing to prevent OOM / slowdown" if defined?(RBS::Test)
  end

  # --- Helpers for XML AST round-trip ---

  def build_unmapped_ast(depth, seed)
    tag = "node_#{seed % 7}"
    attr_samples = [
      "plain",
      "a&b",
      "1<2>3",
      'say "hi"',
      "it's_ok",
      "日本語_#{seed % 11}"
    ]
    attrs = {}
    (seed % 3).times do |i|
      attrs["attr_#{i}"] = attr_samples[(seed + i) % attr_samples.size]
    end

    if depth <= 1 || (seed % 4).zero?
      text = if (seed % 3).zero?
               nil
             elsif (seed % 3) == 1
               "text_#{seed % 13}_<&>\"'"
             else
               "日本語テキスト_#{seed % 17}"
             end
      { tag: tag, attrs: attrs, children: [], text: text }
    else
      child_count = 1 + (seed % 3)
      children = Array.new(child_count) { |i| build_unmapped_ast(depth - 1, (seed * 31) + i + 7) }
      { tag: tag, attrs: attrs, children: children, text: nil }
    end
  end

  def serialize_unmapped_nodes(nodes)
    out = +""
    builder = Xlsxrb::Ooxml::XmlBuilder.new(out)
    builder.open_tag("unmappedRoot")
    nodes.each { |n| builder.write_unmapped(n) }
    builder.close_tag("unmappedRoot")
    out
  end

  def parse_via_base_listener(xml)
    listener = Xlsxrb::Ooxml::XmlParser::BaseListener.new
    Xlsxrb::Ooxml::XmlParser.parse(xml, listener)
    listener.unmapped_data
  end

  def assert_rubyzip_and_zipreader_agree(zip_bytes, expected_entries)
    expected_names = expected_entries.map(&:first)

    # 1. Verify rubyzip can open and extract every entry identically
    Zip::File.open_buffer(StringIO.new(zip_bytes)) do |zf|
      actual_names = zf.entries.map { |e| e.name.dup.force_encoding("UTF-8") }
      assert_equal(expected_names, actual_names, "rubyzip entry names mismatch")
      expected_entries.each do |name, expected_data|
        entry = zf.find_entry(name.b)
        assert_not_nil(entry, "rubyzip missing entry #{name}")
        assert_equal(!name.ascii_only?, entry.gp_flags.anybits?(0x0800), "UTF-8 flag mismatch for #{name}")
        actual_data = entry.get_input_stream.read.b
        assert_equal(expected_data.b, actual_data, "rubyzip content mismatch for #{name}")
      end
    end

    # 2. Verify Xlsxrb::Ooxml::ZipReader agrees
    reader = Xlsxrb::Ooxml::ZipReader.new(StringIO.new(zip_bytes))
    assert_equal(expected_names, reader.entry_names, "ZipReader entry names mismatch")
    all = reader.read_all
    expected_entries.each do |name, expected_data|
      assert_equal(expected_data.b, all[name].b, "ZipReader read_all mismatch for #{name}")
      streamed = +"".b
      stream_io = reader.open_entry_io(name)
      begin
        while (chunk = stream_io.read(7))
          streamed << chunk
        end
      ensure
        stream_io.close
      end
      assert_equal(expected_data.b, streamed, "ZipReader open_entry_io mismatch for #{name}")
    end
  end

  # --- PBT 1: rubyzip Differential ZIP Oracle & ZipGenerator ---

  test "property: rubyzip differential ZIP oracle and ZipGenerator roundtrip" do
    entry_name_arb = Pbt.constant_from(
      "alpha.txt",
      "dir/sub/beta.dat",
      "日本語ファイル.bin",
      "utf8_📝_entry.dat",
      "nested/empty.txt"
    )
    payload_arb = Pbt.one_of(
      Pbt.constant("".b),
      Pbt.constant("hello world".b),
      Pbt.constant(("compressible line\n" * 20).b),
      Pbt.constant("\x00\x01\x02\xFE\xFFrandom\x89PNG".b),
      Pbt.printable_ascii_string(min: 1, max: 60).map(:b.to_proc, :itself.to_proc)
    )
    entry_spec_arb = Pbt.fixed_hash(
      name: entry_name_arb,
      payload: payload_arb,
      mode: Pbt.constant_from(:add_entry, :stream_entry, :add_binary),
      chunk_size: Pbt.integer(min: 1, max: 19)
    )

    Pbt.assert(num_runs: 25) do
      Pbt.property(
        Pbt.array(entry_spec_arb, min: 1, max: 6),
        Pbt.boolean,
        Pbt.constant_from(0, 1, 6, 9),
        Pbt.constant_from(nil, 64),
        Pbt.boolean
      ) do |raw_specs, use_unseekable, comp_level, zip64_limit, use_rubyzip_stored|
        # Deduplicate names to maintain unique archive entries
        seen = {}
        entries_spec = raw_specs.each_with_index.map do |spec, idx|
          name = seen.key?(spec[:name]) ? "e#{idx}_#{spec[:name]}" : spec[:name]
          seen[name] = true
          spec.merge(name: name)
        end

        # Direction A: Xlsxrb::Ooxml::ZipWriter -> rubyzip & ZipReader
        target = use_unseekable ? UnseekableSink.new : StringIO.new
        zw = Xlsxrb::Ooxml::ZipWriter.new(target, compression_level: comp_level)
        zw.instance_variable_set(:@zip64_limit, zip64_limit) if zip64_limit

        expected_a = []
        entries_spec.each do |spec|
          data = spec[:payload]
          case spec[:mode]
          when :add_entry
            utf8 = data.dup.force_encoding(Encoding::UTF_8)
            text = utf8.valid_encoding? ? utf8 : "utf8_fallback_#{spec[:name]}"
            zw.add_entry(spec[:name], text)
            expected_a << [spec[:name], text.b]
          when :add_binary
            zw.add_binary_entry(spec[:name], data)
            expected_a << [spec[:name], data.b]
          when :stream_entry
            zw.start_entry(spec[:name])
            offset = 0
            while offset < data.bytesize
              zw.write_data(data.byteslice(offset, spec[:chunk_size]))
              offset += spec[:chunk_size]
            end
            zw.finish_entry
            expected_a << [spec[:name], data.b]
          end
        end
        zw.close

        zip_bytes_a = use_unseekable ? target.buffer : target.string.b
        assert_rubyzip_and_zipreader_agree(zip_bytes_a, expected_a)

        # Direction B: rubyzip Zip::OutputStream -> Xlsxrb::Ooxml::ZipReader
        expected_b = entries_spec.map { |s| [s[:name], s[:payload].b] }
        comp_method = use_rubyzip_stored ? Zip::Entry::STORED : Zip::Entry::DEFLATED
        rz_buf = Zip::OutputStream.write_buffer do |zos|
          expected_b.each do |name, data|
            zos.put_next_entry(name, nil, nil, comp_method)
            zos.write(data)
          end
        end
        rz_bytes = rz_buf.string.b
        reader_b = Xlsxrb::Ooxml::ZipReader.new(StringIO.new(rz_bytes))
        all_b = reader_b.read_all
        expected_b.each do |name, data|
          assert_equal(data, all_b[name].b, "ZipReader read_all mismatch on rubyzip archive for #{name}")
          assert_equal(data, reader_b.read_entry(name).b, "ZipReader read_entry mismatch on rubyzip archive for #{name}")
          streamed = +"".b
          stream_io = reader_b.open_entry_io(name)
          begin
            while (chunk = stream_io.read(11))
              streamed << chunk
            end
          ensure
            stream_io.close
          end
          assert_equal(data, streamed, "ZipReader open_entry_io mismatch on rubyzip archive for #{name}")
        end

        # Direction C: Xlsxrb::Ooxml::ZipGenerator -> rubyzip & ZipReader + helpers
        Dir.mktmpdir do |dir|
          gen_path = File.join(dir, "gen.xlsx")
          gen = Xlsxrb::Ooxml::ZipGenerator.new(gen_path)
          bin_entries = expected_b.each_with_index.map { |(_n, d), idx| ["part_#{idx}.bin", d] }
          bin_entries.each { |name, data| gen.add_entry(name, data) }
          gen.generate
          assert_rubyzip_and_zipreader_agree(File.binread(gen_path), bin_entries)
        end

        sample16 = comp_level + 100
        assert_equal([sample16].pack("v").bytes, Xlsxrb::Ooxml::ZipGenerator.le16(sample16))
        sample32 = (comp_level * 100_000) + 123
        assert_equal([sample32].pack("V").bytes, Xlsxrb::Ooxml::ZipGenerator.le32(sample32))
        dos_dt = Xlsxrb::Ooxml::ZipGenerator.dos_datetime(Time.utc(2026, 4, 1, 10, 20, 30))
        assert_equal(4, dos_dt.size)
        assert_true(dos_dt.all? { |b| b.is_a?(Integer) && b.between?(0, 255) })
      end
    end
  end

  # --- PBT 2: Forward-Compatibility unmapped_data / Unknown XML ---

  test "property: forward-compatibility unmapped XML roundtrip via BaseListener, write_unmapped, and XLSX injection" do
    Pbt.assert(num_runs: 25) do
      Pbt.property(
        Pbt.array(Pbt.integer(min: 1, max: 10_000), min: 1, max: 4),
        Pbt.integer(min: 1, max: 4),
        Pbt.alphanumeric_string(min: 1, max: 12),
        Pbt.boolean
      ) do |seeds, depth, ext_text, use_stream_writer|
        # Part A: BaseListener <-> XmlBuilder#write_unmapped fixed-point over random XML trees
        asts = seeds.map { |seed| build_unmapped_ast(depth, seed) }
        xml1 = serialize_unmapped_nodes(asts)
        parsed1 = parse_via_base_listener(xml1).first[:children]
        assert_equal(asts, parsed1, "First BaseListener parse did not preserve AST")

        xml2 = serialize_unmapped_nodes(parsed1)
        assert_equal(xml1, xml2, "Re-serialized unmapped XML diverged")
        assert_equal(parsed1, parse_via_base_listener(xml2).first[:children])

        # Part B: Unknown extension elements & attributes injected into a real XLSX workbook
        base_bytes = if use_stream_writer
                       io = StringIO.new
                       Xlsxrb.write(io) do |w|
                         w.style("hdr", bold: true)
                         w.sheet("ForwardSheet") do |s|
                           s.row(%w[Key Val], cell_styles: %w[hdr hdr])
                           s.row(["item_#{seeds.first}", Xlsxrb.formula("B2*1", cached_value: 42)])
                         end
                       end
                       io.string.b
                     else
                       wb = Xlsxrb.build do |w|
                         w.style("hdr", bold: true)
                         w.sheet("ForwardSheet") do |s|
                           s.row(%w[Key Val], cell_styles: %w[hdr hdr])
                           s.row(["item_#{seeds.first}", Xlsxrb.formula("B2*1", cached_value: 42)])
                         end
                       end
                       Xlsxrb.write(wb)
                     end

        parts = Xlsxrb::Ooxml::ZipReader.new(StringIO.new(base_bytes)).read_all
        ext_uri = "{00000000-0000-0000-0000-#{format("%012d", seeds.first % 999_999)}}"
        sheet_ext = %(<extLst><ext uri="#{ext_uri}" xmlns:xtest="urn:xlsxrb:pbt"><xtest:tag note="#{ext_text}">#{ext_text}</xtest:tag></ext></extLst>)
        wb_ext = %(<extLst><ext uri="#{ext_uri}" xmlns:xtest="urn:xlsxrb:wb"><xtest:wbTag v="1"/></ext></extLst>)

        patched_parts = parts.dup
        patched_parts["xl/worksheets/sheet1.xml"] = patched_parts["xl/worksheets/sheet1.xml"].sub(
          "</worksheet>",
          "#{sheet_ext}</worksheet>"
        )
        patched_parts["xl/workbook.xml"] = patched_parts["xl/workbook.xml"].sub(
          "</workbook>",
          "#{wb_ext}</workbook>"
        )

        patched_io = StringIO.new
        Xlsxrb::Ooxml::ZipWriter.open(patched_io) do |zip|
          patched_parts.each { |name, content| zip.add_binary_entry(name, content) }
        end
        patched_bytes = patched_io.string.b
        assert_empty(ExcelRules.check(patched_bytes).map(&:to_s))

        # Verify both in-memory load and streaming read tolerate unknown XML and preserve known content
        loaded = Xlsxrb.load(StringIO.new(patched_bytes))
        ws = loaded.sheet(0)
        assert_equal("ForwardSheet", ws.name)
        assert_equal("Key", ws[0, 0].value)
        assert_equal("Val", ws[0, 1].value)
        assert_equal("item_#{seeds.first}", ws[1, 0].value)
        assert_equal(42, ws[1, 1].value)
        assert_equal("B2*1", ws[1, 1].formula_expression)
        assert_true(loaded.styles.fonts[loaded.styles.cell_xf(ws[0, 0].style_index).font_id.to_i].bold?)

        stream_rows = []
        Xlsxrb.read(StringIO.new(patched_bytes)) { |ss| ss.each { |r| stream_rows << r.values } }
        assert_equal([%w[Key Val], ["item_#{seeds.first}", 42]], stream_rows)

        resaved = loaded.to_data
        assert_empty(ExcelRules.check(resaved).map(&:to_s))
        assert_empty(XsdValidator.check(resaved))
        reloaded = Xlsxrb.load(StringIO.new(resaved)).sheet(0)
        assert_equal("item_#{seeds.first}", reloaded[1, 0].value)
        assert_equal("B2*1", reloaded[1, 1].formula_expression)
      end
    end
  end

  # --- PBT 3: Real-Workbook Display & Type-Cast Consistency ---

  NUMBER_FORMAT_CASES = [
    [0, "General"],
    [1, "0"],
    [2, "0.00"],
    [4, "#,##0.00"],
    [9, "0%"],
    [10, "0.00%"],
    [14, "mm-dd-yy"],
    [15, "d-mmm-yy"],
    [20, "h:mm"],
    [22, "m/d/yy h:mm"],
    ["#,##0.00;(#,##0.00);\"-\"", "#,##0.00;(#,##0.00);\"-\""],
    ["\"$\"#,##0.00", "\"$\"#,##0.00"],
    ["0.0%", "0.0%"],
    ["yyyy/mm/dd", "yyyy/mm/dd"],
    ["yyyy-mm-dd hh:mm:ss", "yyyy-mm-dd hh:mm:ss"],
    ["yyyy年m月d日", "yyyy年m月d日"]
  ].freeze

  def build_display_value(kind_idx, int_arg)
    day_offset = int_arg.abs % 400
    d = Date.new(2025, 1, 15) + day_offset
    sec_offset = int_arg.abs % 86_400
    t = Time.utc(d.year, d.month, d.day, sec_offset / 3600, (sec_offset / 60) % 60, sec_offset % 60)

    case kind_idx
    when 0 then d
    when 1 then t
    when 2 then DateTime.new(d.year, d.month, d.day, t.hour, t.min, t.sec)
    when 3 then FakeTimeWithZone.new(d.year, d.month, d.day, t.hour, t.min, t.sec, 0, [0, 32_400, -18_000][int_arg.abs % 3])
    when 4 then BigDecimal("#{int_arg % 1000}.#{(int_arg.abs % 90) + 10}")
    when 5 then int_arg % 5000
    when 6 then ((int_arg % 2000) / 7.0).round(4)
    when 7 then "Text_#{int_arg.abs % 100}"
    else (int_arg % 2).zero?
    end
  end

  test "property: real-workbook formatted_value, formatted_values, and values(type_cast: true) consistency" do
    Pbt.assert(num_runs: 25) do
      Pbt.property(
        Pbt.boolean,
        Pbt.boolean,
        Pbt.array(
          Pbt.tuple(
            Pbt.choose(0..8),
            Pbt.integer(min: -5000, max: 5000),
            Pbt.choose(0...(NUMBER_FORMAT_CASES.size))
          ),
          min: 2,
          max: 6
        )
      ) do |date1904, shared_strings, col_specs|
        row_values = col_specs.map { |kind, n, _| build_display_value(kind, n) }
        fmt_specs = col_specs.map { |_, _, fidx| NUMBER_FORMAT_CASES[fidx] }

        build_proc = lambda do |w|
          w.workbook_properties(date1904: true) if date1904
          fmt_specs.each_with_index do |(spec_val, _), idx|
            w.style("s_#{idx}", number_format: spec_val)
          end
          w.sheet("Display") do |s|
            s.row(row_values, cell_styles: fmt_specs.each_index.map { |i| "s_#{i}" })
          end
        end

        mem_bytes = Xlsxrb.write(Xlsxrb.build(&build_proc), shared_strings: shared_strings)
        stream_io = StringIO.new
        Xlsxrb.write(stream_io, shared_strings: shared_strings, &build_proc)
        stream_bytes = stream_io.string.b

        [mem_bytes, stream_bytes].each do |bytes|
          assert_empty(ExcelRules.check(bytes).map(&:to_s))
          assert_empty(XsdValidator.check(bytes))
        end

        mem_wb = Xlsxrb.load(StringIO.new(mem_bytes))
        stream_wb = Xlsxrb.load(StringIO.new(stream_bytes))
        assert_equal(date1904, mem_wb.date1904?)
        assert_equal(date1904, stream_wb.date1904?)

        mem_row = mem_wb.sheet(0).rows.first
        stream_loaded_row = stream_wb.sheet(0).rows.first
        assert_equal(mem_row.formatted_values, stream_loaded_row.formatted_values)

        [mem_bytes, stream_bytes].each do |bytes|
          sr_rows = []
          sr_typed = []
          Xlsxrb.read(StringIO.new(bytes)) do |ss|
            ss.each do |sr|
              sr_rows << sr.formatted_values
              sr_typed << sr.values(type_cast: true)
            end
          end
          assert_equal([mem_row.formatted_values], sr_rows, "StreamRow#formatted_values mismatch")

          typed_row = sr_typed.first
          mem_row.cells.each_with_index do |cell, idx|
            expected_fmt = Xlsxrb.format_value(cell.value, cell.format_code, date1904: date1904)
            assert_equal(expected_fmt, cell.formatted_value)
            assert_equal(expected_fmt, mem_row.formatted_values[idx])

            orig = row_values[idx]
            if orig.is_a?(Date) && !orig.is_a?(DateTime)
              assert_equal(orig, cell.to_date)
              if cell.date_format? && !%i[datetime time].include?(cell.format_type)
                assert_instance_of(Date, typed_row[idx])
                assert_equal(orig, typed_row[idx])
              elsif %i[datetime time].include?(cell.format_type)
                assert_equal(
                  [orig.year, orig.month, orig.day, 0, 0, 0],
                  [typed_row[idx].year, typed_row[idx].month, typed_row[idx].day, typed_row[idx].hour, typed_row[idx].min, typed_row[idx].sec]
                )
              end
            elsif orig.is_a?(Time) || orig.is_a?(DateTime)
              actual_t = cell.to_time
              exp_wall = [orig.year, orig.month, orig.day, orig.hour, orig.min, orig.sec]
              assert_equal(exp_wall, [actual_t.year, actual_t.month, actual_t.day, actual_t.hour, actual_t.min, actual_t.sec])
              if cell.date_format? && !%i[datetime time].include?(cell.format_type)
                assert_equal(Date.new(orig.year, orig.month, orig.day), typed_row[idx])
              elsif %i[datetime time].include?(cell.format_type)
                assert_equal(exp_wall, [typed_row[idx].year, typed_row[idx].month, typed_row[idx].day, typed_row[idx].hour, typed_row[idx].min, typed_row[idx].sec])
              end
            end
          end
        end
      end
    end
  end

  # --- PBT 4: Top-Level Xlsxrb.write on Unseekable IO Stateful PBT ---

  UnseekableState = Data.define(
    :shared_strings,
    :password,
    :default_font,
    :compression_level,
    :sheets,
    :active_sheet,
    :styles,
    :rows_by_sheet,
    :merges_by_sheet,
    :comments_by_sheet,
    :hyperlinks_by_sheet,
    :images_count,
    :charts_count
  )

  # Records DSL operations and replays them through top-level Xlsxrb.write against both
  # a write-only UnseekableSink and a seekable StringIO.
  class UnseekableWriteSut
    def initialize(state)
      @opts = {
        shared_strings: state.shared_strings,
        default_font: state.default_font,
        compression_level: state.compression_level
      }
      if state.password
        @opts[:password] = state.password
        @opts[:encryption_mode] = :agile
      end
      @ops = []
      @sheet_seq = 0
    end

    def record(operation)
      @ops << operation
      nil
    end

    def serialize_pair
      # Group ops per workbook level & sheet blocks for valid Xlsxrb.write block DSL
      grouped_writer = lambda do |wb|
        wb.style("s1", bold: true, number_format: "#,##0.00")
        sheet_blocks = []
        current_sheet_ops = nil

        @ops.each do |op|
          case op[0]
          when :define_style
            wb.style(op[1], **op[2])
          when :add_sheet
            current_sheet_ops = [op[1], []]
            sheet_blocks << current_sheet_ops
          else
            current_sheet_ops[1] << op if current_sheet_ops
          end
        end

        sheet_blocks.each do |sheet_name, s_ops|
          wb.sheet(sheet_name) do |s|
            s_ops.each do |sop|
              case sop[0]
              when :row
                s.row(sop[1], cell_styles: sop[2])
              when :merge
                s.merge(sop[1])
              when :comment
                s.comment(sop[1], sop[2])
              when :hyperlink
                s.hyperlink(sop[1], sop[2])
              when :image
                s.image(TINY_PNG, **sop[1])
              when :chart
                s.chart(**sop[1])
              end
            end
          end
        end
      end

      sink = UnseekableSink.new
      seekable_io = StringIO.new
      Xlsxrb.write(sink, **@opts, &grouped_writer)
      Xlsxrb.write(seekable_io, **@opts, &grouped_writer)
      [sink.buffer, seekable_io.string.b]
    end
  end

  test "stateful property: top-level Xlsxrb.write on unseekable write-only IO parity with seekable StringIO" do
    init_arb = Pbt.fixed_hash(
      shared_strings: Pbt.boolean,
      password: Pbt.constant_from(nil, "pbt-pass"),
      default_font: Pbt.constant_from(nil, { name: "Arial", size: 10 }),
      compression_level: Pbt.constant_from(0, 1, 6)
    ).map(
      lambda do |h|
        UnseekableState.new(
          shared_strings: h[:shared_strings],
          password: h[:password],
          default_font: h[:default_font],
          compression_level: h[:compression_level],
          sheets: [],
          active_sheet: nil,
          styles: { "s1" => true },
          rows_by_sheet: {},
          merges_by_sheet: {},
          comments_by_sheet: {},
          hyperlinks_by_sheet: {},
          images_count: {},
          charts_count: {}
        )
      end,
      ->(s) { { shared_strings: s.shared_strings, password: s.password, default_font: s.default_font, compression_level: s.compression_level } }
    )

    Pbt.assert(num_runs: 10) do
      Pbt.stateful(max_steps: 12) do |m|
        m.initial_state init_arb
        m.sut { |init| UnseekableWriteSut.new(init) }

        m.command :add_sheet, weight: 3 do |c|
          c.guard { |s| s.sheets.size < 3 }
          c.arguments Pbt.constant_from("Stream1", "Report's", "Sheet_3")
          c.precondition { |s, name| !s.sheets.include?(name) }
          c.next_state do |s, name|
            s.with(
              sheets: [*s.sheets, name],
              active_sheet: name,
              rows_by_sheet: s.rows_by_sheet.merge(name => []),
              merges_by_sheet: s.merges_by_sheet.merge(name => []),
              comments_by_sheet: s.comments_by_sheet.merge(name => []),
              hyperlinks_by_sheet: s.hyperlinks_by_sheet.merge(name => {}),
              images_count: s.images_count.merge(name => 0),
              charts_count: s.charts_count.merge(name => 0)
            )
          end
          c.run { |sut, name| sut.record([:add_sheet, name]) }
        end

        m.command :define_style, weight: 2 do |c|
          c.arguments(name: Pbt.constant_from("s2", "s3"), italic: Pbt.boolean)
          c.next_state { |s, args| s.with(styles: s.styles.merge(args[:name] => true)) }
          c.run { |sut, args| sut.record([:define_style, args[:name], { italic: args[:italic] }]) }
        end

        m.command :append_row, weight: 5 do |c|
          c.guard { |s| !s.active_sheet.nil? && s.rows_by_sheet.fetch(s.active_sheet).size < 6 }
          c.arguments(
            kind: Pbt.choose(0..4),
            n: Pbt.integer(min: 1, max: 500),
            txt: Pbt.alphanumeric_string(min: 1, max: 10),
            use_style: Pbt.boolean
          )
          c.next_state do |s, args|
            sname = s.active_sheet
            rows = s.rows_by_sheet.fetch(sname)
            val = case args[:kind]
                  when 0 then args[:n]
                  when 1 then BigDecimal("#{args[:n]}.25")
                  when 2 then Date.new(2026, 4, 1) + (args[:n] % 20)
                  when 3 then FakeTimeWithZone.new(2026, 4, 1, 9, 30, 0, 0, 32_400)
                  else Xlsxrb::Elements::Formula.new(expression: "A1+#{args[:n]}", cached_value: args[:n])
                  end
            row_vals = [args[:txt], val]
            s.with(rows_by_sheet: s.rows_by_sheet.merge(sname => [*rows, row_vals]))
          end
          c.run do |sut, args|
            val = case args[:kind]
                  when 0 then args[:n]
                  when 1 then BigDecimal("#{args[:n]}.25")
                  when 2 then Date.new(2026, 4, 1) + (args[:n] % 20)
                  when 3 then FakeTimeWithZone.new(2026, 4, 1, 9, 30, 0, 0, 32_400)
                  else Xlsxrb::Elements::Formula.new(expression: "A1+#{args[:n]}", cached_value: args[:n])
                  end
            cstyles = args[:use_style] ? ["s1", nil] : nil
            sut.record([:row, [args[:txt], val], cstyles])
          end
        end

        m.command :add_merge, weight: 2 do |c|
          c.guard { |s| !s.active_sheet.nil? && s.merges_by_sheet.fetch(s.active_sheet).empty? && !s.rows_by_sheet.fetch(s.active_sheet).empty? }
          c.next_state { |s, _| s.with(merges_by_sheet: s.merges_by_sheet.merge(s.active_sheet => ["C1:D1"])) }
          c.run { |sut, _| sut.record([:merge, "C1:D1"]) }
        end

        m.command :add_comment_and_link, weight: 2 do |c|
          c.guard { |s| !s.active_sheet.nil? && s.comments_by_sheet.fetch(s.active_sheet).empty? && !s.rows_by_sheet.fetch(s.active_sheet).empty? }
          c.arguments Pbt.alphanumeric_string(min: 1, max: 12)
          c.next_state do |s, txt|
            sname = s.active_sheet
            s.with(
              comments_by_sheet: s.comments_by_sheet.merge(sname => [["A1", txt]]),
              hyperlinks_by_sheet: s.hyperlinks_by_sheet.merge(sname => { "B1" => "https://example.com/#{txt}" })
            )
          end
          c.run do |sut, txt|
            sut.record([:comment, "A1", txt])
            sut.record([:hyperlink, "B1", "https://example.com/#{txt}"])
          end
        end

        m.command :add_image, weight: 2 do |c|
          c.guard { |s| !s.active_sheet.nil? && s.images_count.fetch(s.active_sheet) < 1 }
          c.next_state do |s, _|
            s.with(images_count: s.images_count.merge(s.active_sheet => s.images_count.fetch(s.active_sheet) + 1))
          end
          c.run { |sut, _| sut.record([:image, { from_column: 4, from_row: 0, width: 40, height: 40 }]) }
        end

        m.command :add_chart, weight: 2 do |c|
          c.guard { |s| !s.active_sheet.nil? && s.rows_by_sheet.fetch(s.active_sheet).size >= 2 && s.charts_count.fetch(s.active_sheet).zero? }
          c.arguments { |s| Pbt.constant(s.active_sheet) }
          c.next_state { |s, sname| s.with(charts_count: s.charts_count.merge(sname => 1)) }
          c.run do |sut, sname|
            q = Xlsxrb::DslHelpers.quote_sheet_name(sname)
            sut.record([:chart, {
                         type: :bar,
                         title: "C_#{sname}",
                         from_column: 5,
                         from_row: 5,
                         to_column: 10,
                         to_row: 14,
                         series: [{ name: "S", categories: "#{q}!$A$1:$A$2", values: "#{q}!$B$1:$B$2" }]
                       }])
          end
        end

        m.invariant(:unseekable_and_seekable_parity) do |state, sut|
          next if state.sheets.empty?

          unseek_bytes, seek_bytes = sut.serialize_pair

          if state.password.nil?
            [unseek_bytes, seek_bytes].each do |bytes|
              assert_empty(ExcelRules.check(bytes).map(&:to_s))
              assert_empty(XsdValidator.check(bytes))
              Zip::File.open_buffer(StringIO.new(bytes)) { |zf| assert_operator(zf.size, :>=, 5) }
            end
          end

          u_wb = Xlsxrb.load(StringIO.new(unseek_bytes), password: state.password)
          s_wb = Xlsxrb.load(StringIO.new(seek_bytes), password: state.password)

          assert_equal(state.sheets, u_wb.sheet_names)
          assert_equal(state.sheets, s_wb.sheet_names)
          assert_equal(u_wb.styles.fonts.first&.name, s_wb.styles.fonts.first&.name)

          state.sheets.each_with_index do |sname, idx|
            uws = u_wb.sheet(idx)
            sws = s_wb.sheet(idx)
            assert_equal(state.rows_by_sheet.fetch(sname).size, uws.rows.size)
            assert_equal(uws.rows.map(&:formatted_values), sws.rows.map(&:formatted_values))
            assert_equal(state.merges_by_sheet.fetch(sname), uws.merged_cells)
            assert_equal(uws.merged_cells, sws.merged_cells)
            assert_equal(state.comments_by_sheet.fetch(sname), uws.comments.map { |cm| [cm.ref, cm.text] })
            assert_equal(uws.comments.map { |cm| [cm.ref, cm.text] }, sws.comments.map { |cm| [cm.ref, cm.text] })
            assert_equal(state.hyperlinks_by_sheet.fetch(sname), uws.hyperlinks.transform_values(&:url))
            assert_equal(uws.hyperlinks.transform_values(&:url), sws.hyperlinks.transform_values(&:url))
            assert_equal(state.images_count.fetch(sname), uws.images.size)
            assert_equal(uws.images.map(&:data), sws.images.map(&:data))
            assert_equal(state.charts_count.fetch(sname), uws.charts.size)
            assert_equal(uws.charts.map(&:title), sws.charts.map(&:title))
          end

          u_stream = []
          s_stream = []
          Xlsxrb.read(StringIO.new(unseek_bytes), password: state.password) { |sh| u_stream << [sh.name, sh.map(&:values)] }
          Xlsxrb.read(StringIO.new(seek_bytes), password: state.password) { |sh| s_stream << [sh.name, sh.map(&:values)] }
          assert_equal(s_stream, u_stream)
        end
      end
    end
  end

  # --- PBT 5: Public Pure Utilities Algebraic Properties ---

  test "property: public pure utilities (sanitize_sheet_name, parse_range/range_to_ref, split_coordinate, palette_color, to_hex_color)" do
    raw_sheet_name_arb = Pbt.one_of(
      Pbt.constant(""),
      Pbt.constant("   "),
      Pbt.constant("History"),
      Pbt.constant("history"),
      Pbt.constant("'LeadingQuote'"),
      Pbt.constant("Sales: 2026/Q1 [Draft]?*\\"),
      Pbt.constant("Ctrl\x00\x01\x1FName"),
      Pbt.constant("😀" * 20),
      Pbt.constant("#{"あ" * 28}😀😀😀"),
      Pbt.printable_ascii_string(min: 0, max: 45),
      Pbt.alphanumeric_string(min: 1, max: 35).map(->(s) { "'#{s}/:?*[]\\'" }, :itself.to_proc)
    )

    color_input_arb = Pbt.one_of(
      Pbt.integer(min: 0, max: 63),
      Pbt.constant_from(:red, :navy, :blue, :green, :orange, :purple, :white, :black),
      Pbt.constant_from("#F00", "#0aB", "#112233", "AABBCC", "#80FF1122", "FF00AAFF"),
      Pbt.integer(min: 0, max: 0xFF_FF_FF)
    )

    Pbt.assert(num_runs: 60) do
      Pbt.property(
        raw_sheet_name_arb,
        Pbt.constant_from("_", "-", "X"),
        Pbt.integer(min: 0, max: 500),
        Pbt.integer(min: 0, max: 200),
        Pbt.integer(min: 0, max: 500),
        Pbt.integer(min: 0, max: 200),
        color_input_arb,
        Pbt.boolean
      ) do |raw_name, rep_char, r1, c1, r2, c2, color_in, with_alpha|
        # 1. sanitize_sheet_name
        sanitized = Xlsxrb.sanitize_sheet_name(raw_name, replace_char: rep_char)
        assert_true(sanitized.valid_encoding?)
        assert_operator(Xlsxrb::Ooxml::Utils.utf16_length(sanitized), :<=, 31)
        assert_true(Xlsxrb::Elements::Worksheet.valid_name?(sanitized), "Sanitized name invalid: #{sanitized.inspect}")
        assert_equal(sanitized, Xlsxrb.sanitize_sheet_name(sanitized, replace_char: rep_char), "sanitize_sheet_name must be idempotent")

        # 2. parse_range / range_to_ref / split_coordinate
        top_r, bot_r = [r1, r2].minmax
        left_c, right_c = [c1, c2].minmax
        canonical = [[top_r, left_c], [bot_r, right_c]]
        ref = Xlsxrb.range_to_ref(canonical)
        assert_equal(ref, Xlsxrb.range_to_ref(top_r, left_c, bot_r, right_c))
        assert_equal(canonical, Xlsxrb.parse_range(ref))

        reversed_ref = "#{Xlsxrb.cell_to_ref(r2, c2)}:#{Xlsxrb.cell_to_ref(r1, c1)}"
        assert_equal(canonical, Xlsxrb.parse_range(reversed_ref))

        single_ref = Xlsxrb.range_to_ref(top_r, left_c, top_r, left_c)
        assert_not_include(single_ref, ":")
        assert_equal([[top_r, left_c], [top_r, left_c]], Xlsxrb.parse_range(single_ref))
        assert_equal([Xlsxrb.column_letter(left_c), top_r + 1], Xlsxrb.split_coordinate(single_ref))

        assert_nil(Xlsxrb.parse_range(nil))
        assert_nil(Xlsxrb.parse_range(""))
        assert_nil(Xlsxrb.parse_range("INVALID!"))
        assert_nil(Xlsxrb.split_coordinate(nil))
        assert_nil(Xlsxrb.split_coordinate("123"))

        # 3. palette_color / to_hex_color
        pal_idx = r1 % 64
        pal6 = Xlsxrb.palette_color(pal_idx, alpha: false)
        pal8 = Xlsxrb.palette_color(pal_idx, alpha: true)
        assert_match(/\A\h{6}\z/, pal6)
        assert_equal("FF#{pal6}", pal8)
        assert_nil(Xlsxrb.palette_color(-1))
        assert_nil(Xlsxrb.palette_color(64))
        assert_nil(Xlsxrb.palette_color("invalid"))

        hex = Xlsxrb.to_hex_color(color_in, alpha: with_alpha)
        assert_not_nil(hex)
        pattern = with_alpha ? /\A[0-9A-F]{8}\z/ : /\A[0-9A-F]{6}\z/
        assert_match(pattern, hex)
        assert_equal(hex, Xlsxrb.to_hex_color(hex, alpha: with_alpha), "to_hex_color must be idempotent")
      end
    end
  end
end
