# frozen_string_literal: true

require "zlib"
require "rexml/parsers/sax2parser"
require "rexml/sax2listener"

module Xlsxrb
  # Reads cells from an XLSX file.
  class Reader
    def initialize(filepath)
      @filepath = filepath
    end

    # Returns cells for the given sheet (by name or 0-based index).
    # Defaults to the first sheet. Numeric cells with date numFmt are converted to Date.
    def cells(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return {} if worksheet_xml.nil? || worksheet_xml.empty?

      shared_strings = load_shared_strings
      raw_cells = parse_worksheet_cells(worksheet_xml, shared_strings)

      # Resolve date-formatted cells.
      styles = load_styles
      return raw_cells if styles.empty?

      cell_style_map = parse_cell_style_indices(worksheet_xml)
      resolve_date_cells(raw_cells, cell_style_map, styles)
    end

    # Returns column widths as { "A" => 20.0, "B" => 15.5 } for the given sheet.
    def columns(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return {} if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_columns(worksheet_xml)
    end

    # Returns column attributes as { "A" => { hidden: true, outline_level: 1 } }.
    def column_attributes(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return {} if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_column_attributes(worksheet_xml)
    end

    # Returns row attributes as { 1 => { height: 25.0 }, 3 => { hidden: true } }.
    def row_attributes(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return {} if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_row_attributes(worksheet_xml)
    end

    # Returns merged cell ranges as ["A1:B2", "C3:D4"].
    def merged_cells(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return [] if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_merge_cells(worksheet_xml)
    end

    # Returns hyperlinks as { "A1" => "https://example.com" }.
    def hyperlinks(sheet: nil)
      sheets = discover_sheets
      raise ArgumentError, "workbook has no sheets" if sheets.empty?

      target = resolve_sheet_target(sheets, sheet)
      raise ArgumentError, "sheet not found: #{sheet.inspect}" if target.nil?

      entry_path = if target.start_with?("/")
                     target.delete_prefix("/")
                   else
                     "xl/#{target}"
                   end

      worksheet_xml = extract_zip_entry(entry_path)
      return {} if worksheet_xml.nil? || worksheet_xml.empty?

      # Parse hyperlink elements from worksheet.
      parser = REXML::Parsers::SAX2Parser.new(worksheet_xml)
      listener = HyperlinksListener.new
      parser.listen(listener)
      parser.parse

      # Parse rels to resolve rId -> URL.
      rels_path = entry_path.sub(%r{([^/]+)$}, '_rels/\1.rels')
      rels_xml = extract_zip_entry(rels_path)
      rid_to_url = {}
      rid_to_url = parse_rels(rels_xml).transform_values { |v| v } if rels_xml && !rels_xml.empty?

      result = {}
      listener.links.each do |link|
        url = rid_to_url[link[:rid]]
        result[link[:ref]] = url if url
      end
      result
    end

    # Returns cell format codes as { "A1" => "0.00" } for cells with custom numFmt.
    def cell_formats(sheet: nil)
      # Load styles.
      styles = load_styles
      return {} if styles.empty?

      # Parse worksheet to get cell style indices.
      worksheet_xml = load_worksheet_xml(sheet)
      return {} if worksheet_xml.nil? || worksheet_xml.empty?

      parser = REXML::Parsers::SAX2Parser.new(worksheet_xml)
      listener = CellStyleListener.new
      parser.listen(listener)
      parser.parse

      result = {}
      listener.cell_style_indices.each do |cell_ref, xf_index|
        xf = styles[:cell_xfs][xf_index]
        next unless xf

        fmt_id = xf[:num_fmt_id]
        next unless fmt_id && fmt_id != 0

        format_code = styles[:num_fmts][fmt_id]
        result[cell_ref] = format_code if format_code
      end
      result
    end

    # Returns expanded cell style info: { "A1" => { font:, fill:, border:, num_fmt: } }.
    def cell_styles(sheet: nil)
      styles = load_styles
      return {} if styles.empty?

      worksheet_xml = load_worksheet_xml(sheet)
      return {} if worksheet_xml.nil? || worksheet_xml.empty?

      indices = parse_cell_style_indices(worksheet_xml)
      result = {}
      indices.each do |cell_ref, xf_index|
        xf = styles[:cell_xfs][xf_index]
        next unless xf

        entry = {}
        entry[:font] = styles[:fonts][xf[:font_id]] if xf[:font_id]&.positive? && styles[:fonts][xf[:font_id]]
        entry[:fill] = styles[:fills][xf[:fill_id]] if xf[:fill_id]&.positive? && styles[:fills][xf[:fill_id]]
        entry[:border] = styles[:borders][xf[:border_id]] if xf[:border_id]&.positive? && styles[:borders][xf[:border_id]]
        if xf[:num_fmt_id]&.positive?
          code = styles[:num_fmts][xf[:num_fmt_id]]
          entry[:num_fmt] = code if code
        end
        result[cell_ref] = entry unless entry.empty?
      end
      result
    end

    # Returns array of differential formats (dxfs) from the styles.
    def dxfs
      styles = load_styles
      return [] if styles.empty?

      styles[:dxfs] || []
    end

    # Returns array of cellStyleXfs entries (base style format definitions).
    def cell_style_xfs
      styles = load_styles
      return [] if styles.empty?

      styles[:cell_style_xfs] || []
    end

    # Returns array of named cell styles (cellStyle elements).
    def named_cell_styles
      styles = load_styles
      return [] if styles.empty?

      styles[:cell_styles] || []
    end

    # Returns the autoFilter range string (e.g. "A1:B10") or nil.
    def auto_filter(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return nil if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_auto_filter(worksheet_xml)
    end

    # Returns tables for the given sheet as an array of { id:, name:, display_name:, ref:, columns: }.
    def tables(sheet: nil)
      sheet_index = resolve_sheet_index(sheet)
      load_tables(sheet_index)
    end

    # Returns filter columns as { col_id => filter_hash }.
    def filter_columns(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return {} if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_filter_columns(worksheet_xml)
    end

    # Returns sort state as { ref: "A1:B10", sort_conditions: [...] } or nil.
    def sort_state(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return nil if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_sort_state(worksheet_xml)
    end

    # Returns data validations as an array of hashes.
    def data_validations(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return [] if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_data_validations(worksheet_xml)
    end

    # Returns conditional formatting rules for the given sheet.
    def conditional_formats(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return [] if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_conditional_formats(worksheet_xml)
    end

    # Returns sheet-level properties (tabColor, outlinePr) for the given sheet.
    def sheet_properties(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return {} if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_properties(worksheet_xml)
    end

    # Returns sheet protection settings as a hash, or nil if unprotected.
    def sheet_protection(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return nil if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_sheet_protection(worksheet_xml)
    end

    # Returns the dimension ref string (e.g. "A1:B10") for the given sheet.
    def dimension(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return nil if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_dimension(worksheet_xml)
    end

    # Returns sheet format properties (defaultRowHeight, defaultColWidth, baseColWidth).
    def sheet_format(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return {} if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_sheet_format(worksheet_xml)
    end

    # Returns sheet view properties for the given sheet.
    def sheet_view(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return {} if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_sheet_view(worksheet_xml)[:view]
    end

    # Returns freeze pane settings for the given sheet.
    def freeze_pane(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return nil if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_sheet_view(worksheet_xml)[:pane]
    end

    # Returns selection for the given sheet.
    def selection(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return nil if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_sheet_view(worksheet_xml)[:selection]
    end

    # Returns print options for the given sheet.
    def print_options(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return {} if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_print_page(worksheet_xml)[:print_options]
    end

    # Returns page margins for the given sheet.
    def page_margins(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return nil if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_print_page(worksheet_xml)[:page_margins]
    end

    # Returns page setup for the given sheet.
    def page_setup(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return {} if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_print_page(worksheet_xml)[:page_setup]
    end

    # Returns header/footer for the given sheet.
    def header_footer(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return {} if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_print_page(worksheet_xml)[:header_footer]
    end

    # Returns row breaks for the given sheet.
    def row_breaks(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return [] if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_print_page(worksheet_xml)[:row_breaks]
    end

    # Returns column breaks for the given sheet.
    def col_breaks(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return [] if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_print_page(worksheet_xml)[:col_breaks]
    end

    # Returns core properties as a hash (e.g. { title: "...", creator: "..." }).
    def core_properties
      # Discover core properties path from _rels/.rels
      rels_xml = extract_zip_entry("_rels/.rels")
      return {} if rels_xml.nil? || rels_xml.empty?

      rels = parse_rels_with_types(rels_xml)
      core_rel = rels.find { |r| r[:type]&.end_with?("/metadata/core-properties") }
      return {} unless core_rel

      target = core_rel[:target]
      entry_path = target.start_with?("/") ? target.delete_prefix("/") : target
      xml = extract_zip_entry(entry_path)
      return {} if xml.nil? || xml.empty?

      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = CorePropertiesListener.new
      parser.listen(listener)
      parser.parse
      listener.properties
    end

    # Returns app properties as a hash.
    def app_properties
      # Try standard path first, then discover via rels
      xml = extract_zip_entry("docProps/app.xml")
      if xml.nil? || xml.empty?
        rels_xml = extract_zip_entry("_rels/.rels")
        return {} if rels_xml.nil? || rels_xml.empty?

        rels = parse_rels_with_types(rels_xml)
        app_rel = rels.find { |r| r[:type]&.end_with?("/extended-properties") }
        return {} unless app_rel

        target = app_rel[:target]
        entry_path = target.start_with?("/") ? target.delete_prefix("/") : target
        xml = extract_zip_entry(entry_path)
      end
      return {} if xml.nil? || xml.empty?

      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = AppPropertiesListener.new
      parser.listen(listener)
      parser.parse
      listener.properties
    end

    # Returns workbook properties (e.g. { date1904: false, default_theme_version: 166925 }).
    def workbook_properties
      parse_workbook_metadata[:workbook_properties]
    end

    # Returns workbook view properties (e.g. { active_tab: 0 }).
    def workbook_views
      parse_workbook_metadata[:workbook_views]
    end

    # Returns workbook protection settings as a hash, or nil if unprotected.
    def workbook_protection
      parse_workbook_metadata[:workbook_protection]
    end

    # Returns calc properties (e.g. { calc_id: 191029 }).
    def calc_properties
      parse_workbook_metadata[:calc_properties]
    end

    # Returns the calc chain as an array of { ref:, sheet_id: } hashes, or empty array.
    def calc_chain
      xml = extract_zip_entry("xl/calcChain.xml")
      return [] if xml.nil? || xml.empty?

      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = CalcChainListener.new
      parser.listen(listener)
      parser.parse
      listener.entries
    end

    # Returns defined names as an array of hashes.
    def defined_names
      parse_workbook_metadata[:defined_names]
    end

    # Returns sheet states as { "Sheet1" => :visible, "Hidden" => :hidden }.
    def sheet_states
      sheets = discover_sheets
      result = {}
      sheets.each do |s|
        state = case s[:state]
                when "hidden" then :hidden
                when "veryHidden" then :very_hidden
                else :visible
                end
        result[s[:name]] = state
      end
      result
    end

    # Returns ordered sheet names.
    def sheet_names
      discover_sheets.map { |s| s[:name] }
    end

    # Returns all ZIP entry paths in the file.
    def entry_names
      names = []
      File.open(@filepath, "rb") do |file|
        loop do
          sig = file.read(4)
          break if sig.nil? || sig.bytesize < 4

          sig_val = sig.unpack1("V")
          break if [0x02014b50, 0x06054b50].include?(sig_val)
          break unless sig_val == 0x04034b50

          header = file.read(26)
          break if header.nil? || header.bytesize < 26

          _ver, flags, _cm, _mt, _md, _crc, comp_size, _unc, fname_len, extra_len = header.unpack("v v v v v V V V v v")
          break if flags.anybits?(0x0008)

          fname = file.read(fname_len)
          file.read(extra_len)
          file.read(comp_size)
          names << fname
        end
      end
      names
    end

    # Returns raw bytes for a ZIP entry by path.
    def raw_entry(name)
      extract_zip_entry(name)
    end

    # Returns true if the file contains VBA macros (vbaProject.bin).
    def has_macros?
      entry_names.any? { |n| n.include?("vbaProject.bin") }
    end

    # Returns images for the given sheet as an array of hashes.
    # Each hash: { name:, embed_rid:, target:, from_col:, from_row:, to_col:, to_row:, cx:, cy: }
    def images(sheet: nil)
      drawing_xml = load_drawing_xml(sheet)
      return [] if drawing_xml.nil? || drawing_xml.empty?

      sheet_index = resolve_sheet_index(sheet)
      drawing_rels = load_drawing_rels(sheet_index)

      parser = REXML::Parsers::SAX2Parser.new(drawing_xml)
      listener = DrawingImagesListener.new
      parser.listen(listener)
      parser.parse

      listener.images.each do |img|
        target = drawing_rels[img[:embed_rid]]
        img[:target] = target if target
      end
      listener.images
    end

    # Returns charts for the given sheet as an array of hashes.
    # Each hash: { name:, rid:, target:, chart_type:, title: }
    def charts(sheet: nil)
      drawing_xml = load_drawing_xml(sheet)
      return [] if drawing_xml.nil? || drawing_xml.empty?

      sheet_index = resolve_sheet_index(sheet)
      drawing_rels = load_drawing_rels(sheet_index)

      parser = REXML::Parsers::SAX2Parser.new(drawing_xml)
      listener = DrawingChartsListener.new
      parser.listen(listener)
      parser.parse

      listener.charts.each do |chart|
        target = drawing_rels[chart[:rid]]
        next unless target

        chart[:target] = target
        chart_path = resolve_drawing_relative_path(target, sheet_index)
        chart_xml = extract_zip_entry(chart_path)
        next if chart_xml.nil? || chart_xml.empty?

        cp = REXML::Parsers::SAX2Parser.new(chart_xml)
        cl = ChartTypeListener.new
        cp.listen(cl)
        cp.parse
        chart[:chart_type] = cl.chart_type
        chart[:title] = cl.title
        chart[:series] = cl.series unless cl.series.empty?
        chart[:legend] = cl.legend unless cl.legend.empty?
        chart[:data_labels] = cl.data_labels unless cl.data_labels.empty?
        chart[:cat_axis_title] = cl.cat_axis_title if cl.cat_axis_title
        chart[:val_axis_title] = cl.val_axis_title if cl.val_axis_title
      end
      listener.charts
    end

    # Returns shapes for the given sheet as an array of hashes.
    # Each hash: { name:, id:, preset:, text:, from_col:, from_row:, to_col:, to_row: }
    def shapes(sheet: nil)
      drawing_xml = load_drawing_xml(sheet)
      return [] if drawing_xml.nil? || drawing_xml.empty?

      parser = REXML::Parsers::SAX2Parser.new(drawing_xml)
      listener = DrawingShapesListener.new
      parser.listen(listener)
      parser.parse
      listener.shapes
    end

    # Returns comments for the given sheet as an array of hashes.
    # Each hash: { ref:, author:, text: }
    def comments(sheet: nil)
      sheet_index = resolve_sheet_index(sheet)
      comments_path = find_sheet_rel_target(sheet_index, "/comments")
      return [] unless comments_path

      xml = extract_zip_entry(comments_path)
      return [] if xml.nil? || xml.empty?

      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = CommentsListener.new
      parser.listen(listener)
      parser.parse
      listener.comments
    end

    # Returns pivot tables for the given sheet as an array of hashes.
    # Each hash: { name:, ref:, cache_id:, fields:, row_fields:, col_fields:, data_fields: }
    def pivot_tables(sheet: nil)
      sheet_index = resolve_sheet_index(sheet)
      pivot_paths = find_sheet_rel_targets(sheet_index, "/pivotTable")
      return [] if pivot_paths.empty?

      pivot_paths.filter_map do |path|
        xml = extract_zip_entry(path)
        next if xml.nil? || xml.empty?

        parser = REXML::Parsers::SAX2Parser.new(xml)
        listener = PivotTableListener.new
        parser.listen(listener)
        parser.parse
        listener.pivot_table
      end
    end

    private

    def parse_workbook_metadata
      workbook_xml = extract_zip_entry("xl/workbook.xml")
      return { workbook_properties: {}, workbook_views: {}, calc_properties: {}, workbook_protection: nil } if workbook_xml.nil? || workbook_xml.empty?

      parser = REXML::Parsers::SAX2Parser.new(workbook_xml)
      listener = WorkbookListener.new
      parser.listen(listener)
      parser.parse
      {
        workbook_properties: listener.workbook_properties,
        workbook_views: listener.workbook_views,
        calc_properties: listener.calc_properties,
        defined_names: listener.defined_names,
        workbook_protection: listener.workbook_protection
      }
    end

    def load_worksheet_xml(sheet)
      sheets = discover_sheets
      raise ArgumentError, "workbook has no sheets" if sheets.empty?

      target = resolve_sheet_target(sheets, sheet)
      raise ArgumentError, "sheet not found: #{sheet.inspect}" if target.nil?

      # Target may be absolute (/xl/worksheets/sheet1.xml) or relative (worksheets/sheet1.xml).
      entry_path = if target.start_with?("/")
                     target.delete_prefix("/")
                   else
                     "xl/#{target}"
                   end

      extract_zip_entry(entry_path)
    end

    def discover_sheets
      workbook_xml = extract_zip_entry("xl/workbook.xml")
      return [{ name: "Sheet1", rid: "rId1", target: "worksheets/sheet1.xml" }] if workbook_xml.nil? || workbook_xml.empty?

      rels_xml = extract_zip_entry("xl/_rels/workbook.xml.rels")
      rid_to_target = parse_rels(rels_xml)

      sheets = []
      parser = REXML::Parsers::SAX2Parser.new(workbook_xml)
      listener = WorkbookListener.new
      parser.listen(listener)
      parser.parse

      listener.sheets.each do |s|
        target = rid_to_target[s[:rid]]
        sheets << { name: s[:name], rid: s[:rid], target: target, state: s[:state] } if target
      end
      sheets
    end

    def parse_rels(rels_xml)
      return {} if rels_xml.nil? || rels_xml.empty?

      mapping = {}
      parser = REXML::Parsers::SAX2Parser.new(rels_xml)
      listener = RelsListener.new
      parser.listen(listener)
      parser.parse
      listener.relationships.each { |r| mapping[r[:id]] = r[:target] }
      mapping
    end

    def parse_rels_with_types(rels_xml)
      return [] if rels_xml.nil? || rels_xml.empty?

      parser = REXML::Parsers::SAX2Parser.new(rels_xml)
      listener = RelsListener.new
      parser.listen(listener)
      parser.parse
      listener.relationships
    end

    def resolve_sheet_target(sheets, sheet)
      case sheet
      when nil
        sheets.first&.dig(:target)
      when Integer
        sheets[sheet]&.dig(:target)
      when String
        sheets.find { |s| s[:name] == sheet }&.dig(:target)
      else
        raise ArgumentError, "sheet must be a String name or Integer index"
      end
    end

    def load_shared_strings
      sst_xml = extract_zip_entry("xl/sharedStrings.xml")
      return [] if sst_xml.nil? || sst_xml.empty?

      parser = REXML::Parsers::SAX2Parser.new(sst_xml)
      listener = SharedStringsListener.new
      parser.listen(listener)
      parser.parse
      listener.strings
    end

    def resolve_sheet_index(sheet)
      sheets = discover_sheets
      if sheet.nil?
        0
      elsif sheet.is_a?(Integer)
        sheet
      else
        idx = sheets.index { |s| s[:name] == sheet }
        idx || 0
      end
    end

    def load_tables(sheet_index)
      rels_path = "xl/worksheets/_rels/sheet#{sheet_index + 1}.xml.rels"
      rels_xml = extract_zip_entry(rels_path)
      return [] if rels_xml.nil? || rels_xml.empty?

      table_paths = []
      parser = REXML::Parsers::SAX2Parser.new(rels_xml)
      listener = RelsListener.new
      parser.listen(listener)
      parser.parse
      listener.relationships.each do |rel|
        table_paths << rel[:target] if rel[:type]&.end_with?("/table")
      end

      table_paths.map do |rel_target|
        path = if rel_target.start_with?("/")
                 rel_target[1..] # strip leading /
               elsif rel_target.start_with?("..")
                 "xl/#{rel_target.sub("../", "")}"
               else
                 "xl/worksheets/#{rel_target}"
               end
        tbl_xml = extract_zip_entry(path)
        next if tbl_xml.nil? || tbl_xml.empty?

        parse_table_xml(tbl_xml)
      end.compact
    end

    def parse_table_xml(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = TableListener.new
      parser.listen(listener)
      parser.parse
      listener.table
    end

    def load_styles
      styles_xml = extract_zip_entry("xl/styles.xml")
      return {} if styles_xml.nil? || styles_xml.empty?

      parser = REXML::Parsers::SAX2Parser.new(styles_xml)
      listener = StylesListener.new
      parser.listen(listener)
      parser.parse
      {
        num_fmts: listener.num_fmts, cell_xfs: listener.cell_xfs,
        cell_style_xfs: listener.cell_style_xfs, cell_styles: listener.cell_styles,
        fonts: listener.fonts, fills: listener.fills,
        borders: listener.borders, dxfs: listener.dxfs
      }
    end

    def parse_cell_style_indices(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = CellStyleListener.new
      parser.listen(listener)
      parser.parse
      listener.cell_style_indices
    end

    def resolve_date_cells(raw_cells, cell_style_map, styles)
      raw_cells.each do |cell_ref, value|
        next unless value.is_a?(Numeric)

        xf_index = cell_style_map[cell_ref]
        next unless xf_index

        xf = styles[:cell_xfs][xf_index]
        next unless xf

        fmt_id = xf[:num_fmt_id]
        next unless date_format?(fmt_id, styles[:num_fmts])

        raw_cells[cell_ref] = Xlsxrb.serial_to_date(value.to_i)
      end
      raw_cells
    end

    def date_format?(fmt_id, custom_num_fmts)
      return false unless fmt_id

      # Built-in date format IDs.
      return true if BUILTIN_DATE_FMT_IDS.include?(fmt_id)

      # Check custom format code for date-like patterns.
      code = custom_num_fmts[fmt_id]
      return false unless code

      date_pattern?(code)
    end

    def date_pattern?(code)
      # Strip quoted strings to avoid false matches.
      stripped = code.gsub(/"[^"]*"/, "").gsub(/\\[.]/, "")
      stripped.match?(/[ymdhsYMDHS]/)
    end

    def extract_zip_entry(entry_name)
      File.open(@filepath, "rb") do |file|
        loop do
          signature = file.read(4)
          break if signature.nil? || signature.bytesize < 4

          signature_value = signature.unpack1("V")
          break if [0x02014b50, 0x06054b50].include?(signature_value)

          raise Error, "invalid ZIP local header signature" unless signature_value == 0x04034b50

          header = file.read(26)
          raise Error, "truncated ZIP local header" if header.nil? || header.bytesize < 26

          _version, flags, compression_method, _mtime, _mdate, _crc32, compressed_size,
            _uncompressed_size, file_name_length, extra_field_length = header.unpack("v v v v v V V V v v")

          raise Error, "ZIP data descriptor is not supported" if flags.anybits?(0x0008)

          file_name = file.read(file_name_length)
          raise Error, "truncated ZIP file name" if file_name.nil? || file_name.bytesize < file_name_length

          file.read(extra_field_length)

          compressed_data = file.read(compressed_size)
          raise Error, "truncated ZIP entry data" if compressed_data.nil? || compressed_data.bytesize < compressed_size

          next unless file_name == entry_name

          case compression_method
          when 0
            return compressed_data
          when 8
            inflater = Zlib::Inflate.new(-Zlib::MAX_WBITS)
            begin
              return inflater.inflate(compressed_data)
            ensure
              inflater.close
            end
          else
            raise Error, "unsupported ZIP compression method: #{compression_method}"
          end
        end
      end

      nil
    end

    def parse_worksheet_cells(xml, shared_strings)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = WorksheetListener.new(shared_strings)
      parser.listen(listener)
      parser.parse
      listener.cells
    end

    def parse_worksheet_columns(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = ColumnsListener.new
      parser.listen(listener)
      parser.parse
      listener.raw_columns.transform_keys { |idx| column_index_to_letter(idx) }
    end

    def parse_worksheet_column_attributes(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = ColumnsListener.new
      parser.listen(listener)
      parser.parse
      listener.raw_column_attrs.transform_keys { |idx| column_index_to_letter(idx) }
    end

    def parse_worksheet_row_attributes(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = WorksheetListener.new([])
      parser.listen(listener)
      parser.parse
      listener.row_attributes
    end

    def parse_worksheet_merge_cells(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = MergeCellsListener.new
      parser.listen(listener)
      parser.parse
      listener.ranges
    end

    def parse_worksheet_auto_filter(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = AutoFilterListener.new
      parser.listen(listener)
      parser.parse
      listener.ref
    end

    def parse_worksheet_filter_columns(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = AutoFilterListener.new
      parser.listen(listener)
      parser.parse
      listener.filter_columns
    end

    def parse_worksheet_sort_state(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = SortStateListener.new
      parser.listen(listener)
      parser.parse
      listener.sort_state
    end

    def parse_worksheet_data_validations(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = DataValidationsListener.new
      parser.listen(listener)
      parser.parse
      listener.validations
    end

    def parse_worksheet_conditional_formats(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = ConditionalFormattingListener.new
      parser.listen(listener)
      parser.parse
      listener.rules
    end

    def parse_worksheet_properties(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = SheetPropertiesListener.new
      parser.listen(listener)
      parser.parse
      listener.properties
    end

    def parse_worksheet_sheet_protection(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = SheetProtectionListener.new
      parser.listen(listener)
      parser.parse
      listener.protection
    end

    def parse_worksheet_dimension(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = DimensionListener.new
      parser.listen(listener)
      parser.parse
      listener.ref
    end

    def parse_worksheet_sheet_format(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = SheetFormatListener.new
      parser.listen(listener)
      parser.parse
      listener.properties
    end

    def parse_worksheet_sheet_view(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = SheetViewListener.new
      parser.listen(listener)
      parser.parse
      { view: listener.view, pane: listener.pane, selection: listener.selection }
    end

    def parse_worksheet_print_page(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = PrintPageListener.new
      parser.listen(listener)
      parser.parse
      {
        print_options: listener.print_options,
        page_margins: listener.page_margins,
        page_setup: listener.page_setup,
        header_footer: listener.header_footer,
        row_breaks: listener.row_breaks,
        col_breaks: listener.col_breaks
      }
    end

    def column_index_to_letter(index)
      result = +""
      while index.positive?
        index -= 1
        result.prepend(("A".ord + (index % 26)).chr)
        index /= 26
      end
      result
    end

    def load_drawing_xml(sheet)
      sheet_index = resolve_sheet_index(sheet)
      drawing_path = find_sheet_rel_target(sheet_index, "/drawing")
      return nil unless drawing_path

      extract_zip_entry(drawing_path)
    end

    def load_drawing_rels(sheet_index)
      drawing_path = find_sheet_rel_target(sheet_index, "/drawing")
      return {} unless drawing_path

      dir = drawing_path.sub(%r{([^/]+)$}, '_rels/\1.rels')
      rels_xml = extract_zip_entry(dir)
      return {} if rels_xml.nil? || rels_xml.empty?

      parse_rels(rels_xml)
    end

    def find_sheet_rel_target(sheet_index, type_suffix)
      rels_path = "xl/worksheets/_rels/sheet#{sheet_index + 1}.xml.rels"
      rels_xml = extract_zip_entry(rels_path)
      return nil if rels_xml.nil? || rels_xml.empty?

      rels = parse_rels_with_types(rels_xml)
      rel = rels.find { |r| r[:type]&.end_with?(type_suffix) }
      return nil unless rel

      normalize_xl_path(rel[:target], "xl/worksheets")
    end

    def find_sheet_rel_targets(sheet_index, type_suffix)
      rels_path = "xl/worksheets/_rels/sheet#{sheet_index + 1}.xml.rels"
      rels_xml = extract_zip_entry(rels_path)
      return [] if rels_xml.nil? || rels_xml.empty?

      rels = parse_rels_with_types(rels_xml)
      rels.select { |r| r[:type]&.end_with?(type_suffix) }.map do |r|
        normalize_xl_path(r[:target], "xl/worksheets")
      end
    end

    def normalize_xl_path(target, base_dir)
      if target.start_with?("/")
        target[1..]
      elsif target.start_with?("..")
        # Resolve relative to base
        parts = base_dir.split("/") + target.split("/")
        resolved = []
        parts.each { |p| p == ".." ? resolved.pop : resolved << p }
        resolved.join("/")
      else
        "#{base_dir}/#{target}"
      end
    end

    def resolve_drawing_relative_path(target, sheet_index)
      drawing_path = find_sheet_rel_target(sheet_index, "/drawing")
      return target unless drawing_path

      base_dir = File.dirname(drawing_path)
      normalize_xl_path(target, base_dir)
    end

    # SAX2 listener for parsing shared string table (xl/sharedStrings.xml).
    class SharedStringsListener
      include REXML::SAX2Listener

      attr_reader :strings

      def initialize
        @strings = []
        @inside_si = false
        @inside_r = false
        @inside_rpr = false
        @inside_t = false
        @text_buffer = +""
        @runs = []
        @current_font = {}
        @has_runs = false
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)

        case name
        when "si"
          @inside_si = true
          @text_buffer = +""
          @runs = []
          @has_runs = false
        when "r"
          @inside_r = true
          @has_runs = true
          @current_font = {}
        when "rPr"
          @inside_rpr = true if @inside_r
        when "b"
          @current_font[:bold] = true if @inside_rpr
        when "i"
          @current_font[:italic] = true if @inside_rpr
        when "u"
          @current_font[:underline] = true if @inside_rpr
        when "sz"
          @current_font[:sz] = attributes["val"]&.to_f if @inside_rpr
        when "color"
          @current_font[:color] = attributes["rgb"] if @inside_rpr && attributes["rgb"]
        when "rFont"
          @current_font[:name] = attributes["val"] if @inside_rpr
        when "t"
          @inside_t = @inside_si
          @text_buffer = +""  if @inside_r
        end
      end

      def characters(text)
        @text_buffer << text if @inside_t
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)

        case name
        when "t"
          @inside_t = false
        when "rPr"
          @inside_rpr = false
        when "r"
          run = { text: @text_buffer.dup }
          run[:font] = @current_font.dup unless @current_font.empty?
          @runs << run
          @inside_r = false
        when "si"
          if @has_runs
            @strings << Xlsxrb::RichText.new(runs: @runs)
          else
            @strings << @text_buffer.dup
          end
          @inside_si = false
          @text_buffer = +""
          @runs = []
        end
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing worksheet cells.
    class WorksheetListener
      include REXML::SAX2Listener

      attr_reader :cells, :row_attributes

      def initialize(shared_strings = [])
        @shared_strings = shared_strings
        @cells = {}
        @row_attributes = {}
        @current_cell_ref = nil
        @current_cell_type = nil
        @inside_value = false
        @inside_inline_text = false
        @inside_formula = false
        @inside_is = false
        @inside_is_r = false
        @inside_is_rpr = false
        @value_buffer = +""
        @inline_text_buffer = +""
        @formula_buffer = +""
        @is_runs = []
        @is_has_runs = false
        @is_current_font = {}
        @is_run_text = +""
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)

        case name
        when "row"
          parse_row_attributes(attributes)
        when "c"
          @current_cell_ref = attributes["r"]
          @current_cell_type = attributes["t"]
          @value_buffer = +""
          @inline_text_buffer = +""
          @formula_buffer = +""
          @formula_type = nil
          @formula_ref = nil
          @formula_si = nil
          @is_runs = []
          @is_has_runs = false
        when "v"
          @inside_value = true
        when "f"
          @inside_formula = true
          @formula_type = attributes["t"]
          @formula_ref = attributes["ref"]
          si = attributes["si"]
          @formula_si = si&.to_i
        when "is"
          @inside_is = true if @current_cell_type == "inlineStr"
        when "r"
          if @inside_is
            @inside_is_r = true
            @is_has_runs = true
            @is_current_font = {}
            @is_run_text = +""
          end
        when "rPr"
          @inside_is_rpr = true if @inside_is_r
        when "b"
          @is_current_font[:bold] = true if @inside_is_rpr
        when "i"
          @is_current_font[:italic] = true if @inside_is_rpr
        when "u"
          @is_current_font[:underline] = true if @inside_is_rpr
        when "sz"
          @is_current_font[:sz] = attributes["val"]&.to_f if @inside_is_rpr
        when "color"
          @is_current_font[:color] = attributes["rgb"] if @inside_is_rpr && attributes["rgb"]
        when "rFont"
          @is_current_font[:name] = attributes["val"] if @inside_is_rpr
        when "t"
          if @inside_is_r
            @is_run_text = +""
            @inside_inline_text = true
          elsif @inside_is
            @inside_inline_text = true
          elsif @current_cell_type == "inlineStr" && !@current_cell_ref.nil?
            @inside_inline_text = true
          end
        end
      end

      def characters(text)
        @value_buffer << text if @inside_value
        if @inside_inline_text
          if @inside_is_r
            @is_run_text << text
          else
            @inline_text_buffer << text
          end
        end
        @formula_buffer << text if @inside_formula
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)

        case name
        when "v"
          @inside_value = false
        when "f"
          @inside_formula = false
        when "t"
          @inside_inline_text = false
        when "rPr"
          @inside_is_rpr = false
        when "r"
          if @inside_is_r
            run = { text: @is_run_text.dup }
            run[:font] = @is_current_font.dup unless @is_current_font.empty?
            @is_runs << run
            @inside_is_r = false
          end
        when "is"
          @inside_is = false
        when "c"
          store_cell_value
          @current_cell_ref = nil
          @current_cell_type = nil
          @value_buffer = +""
          @inline_text_buffer = +""
          @formula_buffer = +""
        end
      end

      private

      def parse_row_attributes(attributes)
        row_num = attributes["r"]&.to_i
        return unless row_num

        attrs = {}
        ht = attributes["ht"]
        attrs[:height] = ht.to_f if ht && attributes["customHeight"] == "1"
        attrs[:hidden] = true if attributes["hidden"] == "1"
        ol = attributes["outlineLevel"]
        attrs[:outline_level] = ol.to_i if ol && ol != "0"
        attrs[:collapsed] = true if attributes["collapsed"] == "1"
        @row_attributes[row_num] = attrs unless attrs.empty?
      end

      def store_cell_value
        return if @current_cell_ref.nil?

        unless @formula_buffer.empty?
          cached = @value_buffer.empty? ? nil : @value_buffer.dup
          f_type = case @formula_type
                   when "shared" then :shared
                   when "array" then :array
                   end
          @cells[@current_cell_ref] = Formula.new(
            expression: @formula_buffer.dup,
            cached_value: cached,
            type: f_type,
            ref: @formula_ref,
            shared_index: @formula_si
          )
          return
        end

        case @current_cell_type
        when "inlineStr"
          if @is_has_runs
            @cells[@current_cell_ref] = RichText.new(runs: @is_runs.map(&:dup))
          else
            @cells[@current_cell_ref] = @inline_text_buffer.dup
          end
        when "s"
          index = @value_buffer.to_i
          @cells[@current_cell_ref] = @shared_strings[index] || ""
        when "e"
          @cells[@current_cell_ref] = @value_buffer.dup
        when "b"
          @cells[@current_cell_ref] = @value_buffer.strip == "1"
        when nil, "", "n"
          return if @value_buffer.empty?

          raw = @value_buffer.dup
          @cells[@current_cell_ref] = numeric_value(raw)
        end
      end

      def numeric_value(raw)
        if raw.include?(".")
          raw.to_f
        else
          raw.to_i
        end
      end

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing workbook.xml to discover sheet names, rIds, and workbook-level properties.
    class WorkbookListener
      include REXML::SAX2Listener

      attr_reader :sheets, :workbook_properties, :workbook_views, :calc_properties, :defined_names, :workbook_protection

      def initialize
        @sheets = []
        @workbook_properties = {}
        @workbook_views = {}
        @calc_properties = {}
        @defined_names = []
        @workbook_protection = nil
        @inside_defined_name = false
        @current_dn_attrs = nil
        @dn_text_buffer = +""
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        case name
        when "sheet"
          @sheets << { name: attributes["name"], rid: attributes["r:id"], state: attributes["state"] }
        when "workbookPr"
          d1904 = attributes["date1904"]
          @workbook_properties[:date1904] = %w[1 true].include?(d1904) unless d1904.nil?
          dtv = attributes["defaultThemeVersion"]
          @workbook_properties[:default_theme_version] = dtv.to_i if dtv
        when "workbookView"
          at = attributes["activeTab"]
          @workbook_views[:active_tab] = at.to_i if at
          fs = attributes["firstSheet"]
          @workbook_views[:first_sheet] = fs.to_i if fs
        when "calcPr"
          ci = attributes["calcId"]
          @calc_properties[:calc_id] = ci.to_i if ci
          cm = attributes["calcMode"]
          @calc_properties[:calc_mode] = cm if cm
          fcol = attributes["fullCalcOnLoad"]
          @calc_properties[:full_calc_on_load] = %w[1 true].include?(fcol) unless fcol.nil?
          iter = attributes["iterate"]
          @calc_properties[:iterate] = %w[1 true].include?(iter) unless iter.nil?
          ic = attributes["iterateCount"]
          @calc_properties[:iterate_count] = ic.to_i if ic
          id = attributes["iterateDelta"]
          @calc_properties[:iterate_delta] = id.to_f if id
          rm = attributes["refMode"]
          @calc_properties[:ref_mode] = rm if rm
          cc = attributes["calcCompleted"]
          @calc_properties[:calc_completed] = %w[1 true].include?(cc) unless cc.nil?
          cos = attributes["calcOnSave"]
          @calc_properties[:calc_on_save] = %w[1 true].include?(cos) unless cos.nil?
        when "workbookProtection"
          prot = {}
          ls = attributes["lockStructure"]
          prot[:lock_structure] = %w[1 true].include?(ls) unless ls.nil?
          lw = attributes["lockWindows"]
          prot[:lock_windows] = %w[1 true].include?(lw) unless lw.nil?
          wp = attributes["workbookPassword"]
          prot[:password] = wp if wp
          an = attributes["workbookAlgorithmName"]
          if an
            prot[:algorithm_name] = an
            hv = attributes["workbookHashValue"]
            prot[:hash_value] = hv if hv
            sv = attributes["workbookSaltValue"]
            prot[:salt_value] = sv if sv
            sc = attributes["workbookSpinCount"]
            prot[:spin_count] = sc.to_i if sc
          end
          @workbook_protection = prot unless prot.empty?
        when "definedName"
          @inside_defined_name = true
          @current_dn_attrs = {
            name: attributes["name"],
            hidden: %w[1 true].include?(attributes["hidden"])
          }
          lsi = attributes["localSheetId"]
          @current_dn_attrs[:local_sheet_id] = lsi.to_i if lsi
          @dn_text_buffer = +""
        end
      end

      def characters(text)
        @dn_text_buffer << text if @inside_defined_name
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        return unless name == "definedName" && @inside_defined_name

        @current_dn_attrs[:value] = @dn_text_buffer.dup
        @defined_names << @current_dn_attrs
        @inside_defined_name = false
        @current_dn_attrs = nil
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing .rels files to map rId to Target.
    class RelsListener
      include REXML::SAX2Listener

      attr_reader :relationships

      def initialize
        @relationships = []
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        return unless name == "Relationship"

        @relationships << { id: attributes["Id"], target: attributes["Target"], type: attributes["Type"] }
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing <cols><col> elements from worksheet XML.
    class ColumnsListener
      include REXML::SAX2Listener

      # Returns { column_index => width } hash (1-based indices).
      attr_reader :raw_columns, :raw_column_attrs

      def initialize
        @raw_columns = {}
        @raw_column_attrs = {}
      end

      def columns
        @raw_columns
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        return unless name == "col"

        min_val = attributes["min"]&.to_i
        max_val = attributes["max"]&.to_i
        width = attributes["width"]&.to_f
        return unless min_val && max_val

        (min_val..max_val).each do |i|
          @raw_columns[i] = width if width
          attrs = {}
          attrs[:hidden] = true if attributes["hidden"] == "1"
          attrs[:best_fit] = true if attributes["bestFit"] == "1"
          ol = attributes["outlineLevel"]
          attrs[:outline_level] = ol.to_i if ol && ol != "0"
          attrs[:collapsed] = true if attributes["collapsed"] == "1"
          s = attributes["style"]
          attrs[:style] = s.to_i if s && s != "0"
          @raw_column_attrs[i] = attrs unless attrs.empty?
        end
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing <mergeCells><mergeCell> elements.
    class MergeCellsListener
      include REXML::SAX2Listener

      attr_reader :ranges

      def initialize
        @ranges = []
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        return unless name == "mergeCell"

        ref = attributes["ref"]
        @ranges << ref if ref
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing <hyperlinks><hyperlink> elements.
    class HyperlinksListener
      include REXML::SAX2Listener

      attr_reader :links

      def initialize
        @links = []
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        return unless name == "hyperlink"

        ref = attributes["ref"]
        rid = attributes["r:id"]
        @links << { ref: ref, rid: rid } if ref && rid
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing styles.xml (numFmts + cellXfs).
    class StylesListener
      include REXML::SAX2Listener

      attr_reader :num_fmts, :cell_xfs, :cell_style_xfs, :cell_styles, :fonts, :fills, :borders, :dxfs

      def initialize
        @num_fmts = {} # { numFmtId => formatCode }
        @cell_xfs = [] # Array of { num_fmt_id:, font_id:, fill_id:, border_id: }
        @cell_style_xfs = [] # Array of { num_fmt_id:, font_id:, fill_id:, border_id: }
        @cell_styles = [] # Array of { name:, xf_id:, builtin_id: }
        @fonts = []
        @fills = []
        @borders = []
        @dxfs = []
        @inside_cell_xfs = false
        @inside_cell_style_xfs = false
        @inside_cell_styles = false
        @inside_fonts = false
        @inside_fills = false
        @inside_borders = false
        @inside_dxfs = false
        @current_font = nil
        @current_fill = nil
        @current_border = nil
        @current_border_side = nil
        @current_dxf = nil
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)

        case name
        when "numFmt"
          id = attributes["numFmtId"]&.to_i
          code = attributes["formatCode"]
          @num_fmts[id] = code if id && code
        when "cellXfs"
          @inside_cell_xfs = true
        when "cellStyleXfs"
          @inside_cell_style_xfs = true
        when "cellStyles"
          @inside_cell_styles = true
        when "cellStyle"
          if @inside_cell_styles
            entry = { name: attributes["name"] }
            entry[:xf_id] = attributes["xfId"].to_i if attributes["xfId"]
            entry[:builtin_id] = attributes["builtinId"].to_i if attributes["builtinId"]
            @cell_styles << entry
          end
        when "xf"
          xf_entry = {
            num_fmt_id: attributes["numFmtId"]&.to_i,
            font_id: attributes["fontId"]&.to_i,
            fill_id: attributes["fillId"]&.to_i,
            border_id: attributes["borderId"]&.to_i
          }
          if @inside_cell_xfs
            xf_entry[:xf_id] = attributes["xfId"]&.to_i
            @cell_xfs << xf_entry
          elsif @inside_cell_style_xfs
            @cell_style_xfs << xf_entry
          end
        when "fonts"
          @inside_fonts = true
        when "font"
          @current_font = {} if @inside_fonts || @inside_dxfs
        when "b"
          @current_font[:bold] = true if @current_font
        when "i"
          @current_font[:italic] = true if @current_font
        when "u"
          @current_font[:underline] = true if @current_font
        when "sz"
          @current_font[:sz] = attributes["val"]&.to_f if @current_font
        when "color"
          parse_color(attributes)
        when "fills"
          @inside_fills = true
        when "fill"
          @current_fill = {} if @inside_fills || @inside_dxfs
        when "patternFill"
          @current_fill[:pattern] = attributes["patternType"] if @current_fill
        when "fgColor"
          parse_fill_color(:fg_color, attributes) if @current_fill
        when "bgColor"
          parse_fill_color(:bg_color, attributes) if @current_fill
        when "borders"
          @inside_borders = true
        when "border"
          @current_border = {} if @inside_borders || @inside_dxfs
        when "left", "right", "top", "bottom"
          if @current_border
            style = attributes["style"]
            @current_border_side = name.to_sym
            @current_border[@current_border_side] = { style: style } if style
          end
        when "dxfs"
          @inside_dxfs = true
        when "dxf"
          @current_dxf = {}
        end

        parse_font_name(name, attributes)
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        case name
        when "cellXfs"
          @inside_cell_xfs = false
        when "cellStyleXfs"
          @inside_cell_style_xfs = false
        when "cellStyles"
          @inside_cell_styles = false
        when "fonts"
          @inside_fonts = false
        when "font"
          if @inside_dxfs && @current_dxf
            @current_dxf[:font] = @current_font
          elsif @inside_fonts
            @fonts << @current_font
          end
          @current_font = nil
        when "fills"
          @inside_fills = false
        when "fill"
          if @inside_dxfs && @current_dxf
            @current_dxf[:fill] = @current_fill
          elsif @inside_fills
            @fills << @current_fill
          end
          @current_fill = nil
        when "borders"
          @inside_borders = false
        when "border"
          if @inside_dxfs && @current_dxf
            @current_dxf[:border] = @current_border
          elsif @inside_borders
            @borders << @current_border
          end
          @current_border = nil
        when "left", "right", "top", "bottom"
          @current_border_side = nil
        when "dxfs"
          @inside_dxfs = false
        when "dxf"
          @dxfs << @current_dxf if @current_dxf
          @current_dxf = nil
        end
      end

      private

      def parse_color(attributes)
        if @current_border_side && @current_border
          side_data = @current_border[@current_border_side]
          if side_data.is_a?(Hash)
            side_data[:color] = attributes["rgb"] if attributes["rgb"]
            side_data[:theme] = attributes["theme"].to_i if attributes["theme"]
            side_data[:tint] = attributes["tint"].to_f if attributes["tint"]
            side_data[:indexed] = attributes["indexed"].to_i if attributes["indexed"]
          end
        elsif @current_font
          @current_font[:color] = attributes["rgb"] if attributes["rgb"]
          @current_font[:theme] = attributes["theme"].to_i if attributes["theme"]
          @current_font[:tint] = attributes["tint"].to_f if attributes["tint"]
          @current_font[:indexed] = attributes["indexed"].to_i if attributes["indexed"]
        end
      end

      def parse_fill_color(key, attributes)
        if attributes["rgb"]
          @current_fill[key] = attributes["rgb"]
        elsif attributes["theme"]
          @current_fill[:"#{key}_theme"] = attributes["theme"].to_i
          @current_fill[:"#{key}_tint"] = attributes["tint"].to_f if attributes["tint"]
        elsif attributes["indexed"]
          @current_fill[:"#{key}_indexed"] = attributes["indexed"].to_i
        end
      end

      def parse_font_name(name, attributes)
        @current_font[:name] = attributes["val"] if name == "name" && @current_font && attributes["val"]
      end

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing <autoFilter> element.
    class AutoFilterListener
      include REXML::SAX2Listener

      attr_reader :ref, :filter_columns

      def initialize
        @ref = nil
        @filter_columns = {}
        @current_col_id = nil
        @current_filter = nil
        @inside_custom_filters = false
        @custom_filters_list = []
        @custom_filters_and = false
        @filter_values = []
        @filter_blank = false
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        case name
        when "autoFilter"
          @ref = attributes["ref"]
        when "filterColumn"
          @current_col_id = attributes["colId"]&.to_i
        when "filters"
          @filter_blank = attributes["blank"] == "1"
          @filter_values = []
        when "filter"
          val = attributes["val"]
          @filter_values << val if val
        when "customFilters"
          @inside_custom_filters = true
          @custom_filters_and = attributes["and"] == "1"
          @custom_filters_list = []
        when "customFilter"
          @custom_filters_list << { operator: attributes["operator"], val: attributes["val"] } if @inside_custom_filters
        when "dynamicFilter"
          @current_filter = { type: :dynamic, dynamic_type: attributes["type"] }
        when "top10"
          @current_filter = {
            type: :top10,
            top: attributes["top"] == "1",
            percent: attributes["percent"] == "1",
            val: attributes["val"]&.to_f&.to_i
          }
        end
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        case name
        when "filters"
          f = { type: :filters }
          f[:blank] = true if @filter_blank
          f[:values] = @filter_values unless @filter_values.empty?
          @current_filter = f
        when "customFilters"
          if @custom_filters_list.size == 1
            cf = @custom_filters_list.first
            @current_filter = { type: :custom, operator: cf[:operator], val: cf[:val] }
          else
            @current_filter = { type: :custom, filters: @custom_filters_list, and: @custom_filters_and }
          end
          @inside_custom_filters = false
        when "filterColumn"
          @filter_columns[@current_col_id] = @current_filter if @current_col_id && @current_filter
          @current_col_id = nil
          @current_filter = nil
        end
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing <sortState> element.
    class SortStateListener
      include REXML::SAX2Listener

      attr_reader :sort_state

      def initialize
        @sort_state = nil
        @inside_sort_state = false
        @sort_conditions = []
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        case name
        when "sortState"
          @inside_sort_state = true
          @sort_state = { ref: attributes["ref"], sort_conditions: [] }
        when "sortCondition"
          return unless @inside_sort_state

          sc = { ref: attributes["ref"] }
          sc[:descending] = true if attributes["descending"] == "1"
          @sort_conditions << sc
        end
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        return unless name == "sortState" && @inside_sort_state

        @sort_state[:sort_conditions] = @sort_conditions
        @inside_sort_state = false
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing docProps/app.xml.
    class AppPropertiesListener
      include REXML::SAX2Listener

      attr_reader :properties

      def initialize
        @properties = {}
        @current_field = nil
        @text_buffer = +""
        @inside_vector = false
        @vector_items = []
        @heading_pairs = []
        @titles_of_parts = []
        @inside_heading_pairs = false
        @inside_titles_of_parts = false
        @inside_variant = false
      end

      def start_element(_uri, local_name, qname, _attributes)
        name = element_name(local_name, qname)
        case name
        when "Application", "AppVersion"
          @current_field = name
          @text_buffer = +""
        when "HeadingPairs"
          @inside_heading_pairs = true
        when "TitlesOfParts"
          @inside_titles_of_parts = true
        when "variant"
          @inside_variant = true
          @text_buffer = +""
        when "lpstr", "i4"
          @text_buffer = +""
        end
      end

      def characters(text)
        @text_buffer << text
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        case name
        when "Application"
          @properties[:application] = @text_buffer.dup
          @current_field = nil
        when "AppVersion"
          @properties[:app_version] = @text_buffer.dup
          @current_field = nil
        when "lpstr"
          if @inside_titles_of_parts
            @titles_of_parts << @text_buffer.dup
          elsif @inside_heading_pairs && @inside_variant
            @vector_items << @text_buffer.dup
          end
        when "i4"
          @vector_items << @text_buffer.to_i if @inside_heading_pairs && @inside_variant
        when "variant"
          @inside_variant = false
        when "HeadingPairs"
          # Convert flat array to pairs: [label, count, label, count, ...]
          @heading_pairs = @vector_items.each_slice(2).to_a
          @vector_items = []
          @inside_heading_pairs = false
          @properties[:heading_pairs] = @heading_pairs
        when "TitlesOfParts"
          @inside_titles_of_parts = false
          @properties[:titles_of_parts] = @titles_of_parts
        end
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing docProps/core.xml.
    class CorePropertiesListener
      include REXML::SAX2Listener

      attr_reader :properties

      FIELD_MAP = {
        "title" => :title,
        "creator" => :creator,
        "created" => :created,
        "modified" => :modified
      }.freeze

      def initialize
        @properties = {}
        @current_field = nil
        @text_buffer = +""
      end

      def start_element(_uri, local_name, qname, _attributes)
        name = element_name(local_name, qname)
        return unless FIELD_MAP.key?(name)

        @current_field = FIELD_MAP[name]
        @text_buffer = +""
      end

      def characters(text)
        @text_buffer << text if @current_field
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        return unless @current_field && FIELD_MAP.key?(name)

        @properties[@current_field] = @text_buffer.dup unless @text_buffer.empty?
        @current_field = nil
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener that captures cell style index (s attribute) from worksheet.
    class CellStyleListener
      include REXML::SAX2Listener

      attr_reader :cell_style_indices

      def initialize
        @cell_style_indices = {}
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        return unless name == "c"

        ref = attributes["r"]
        s = attributes["s"]
        @cell_style_indices[ref] = s.to_i if ref && s
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing <sheetPr> element (tabColor, outlinePr).
    class SheetPropertiesListener
      include REXML::SAX2Listener

      attr_reader :properties

      def initialize
        @properties = {}
        @inside_sheet_pr = false
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        case name
        when "sheetPr"
          @inside_sheet_pr = true
        when "tabColor"
          @properties[:tab_color] = attributes["rgb"] if @inside_sheet_pr && attributes["rgb"]
        when "outlinePr"
          if @inside_sheet_pr
            sb = attributes["summaryBelow"]
            @properties[:summary_below] = %w[1 true].include?(sb) unless sb.nil?
            sr = attributes["summaryRight"]
            @properties[:summary_right] = %w[1 true].include?(sr) unless sr.nil?
          end
        end
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        @inside_sheet_pr = false if name == "sheetPr"
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing <sheetProtection> element.
    class SheetProtectionListener
      include REXML::SAX2Listener

      attr_reader :protection

      def initialize
        @protection = nil
      end

      BOOL_ATTRS = %i[sheet objects scenarios select_locked_cells select_unlocked_cells].freeze
      FALSE_ATTRS = %i[format_cells format_columns format_rows insert_columns insert_rows
                       insert_hyperlinks delete_columns delete_rows sort auto_filter pivot_tables].freeze
      ATTR_MAP = {
        "sheet" => :sheet, "objects" => :objects, "scenarios" => :scenarios,
        "formatCells" => :format_cells, "formatColumns" => :format_columns,
        "formatRows" => :format_rows, "insertColumns" => :insert_columns,
        "insertRows" => :insert_rows, "insertHyperlinks" => :insert_hyperlinks,
        "deleteColumns" => :delete_columns, "deleteRows" => :delete_rows,
        "selectLockedCells" => :select_locked_cells, "sort" => :sort,
        "autoFilter" => :auto_filter, "pivotTables" => :pivot_tables,
        "selectUnlockedCells" => :select_unlocked_cells,
        "password" => :password, "algorithmName" => :algorithm_name,
        "hashValue" => :hash_value, "saltValue" => :salt_value, "spinCount" => :spin_count
      }.freeze

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        return unless name == "sheetProtection"

        prot = {}
        ATTR_MAP.each do |xml_attr, sym|
          val = attributes[xml_attr]
          next if val.nil?

          prot[sym] = if sym == :spin_count
                        val.to_i
                      elsif %i[password algorithm_name hash_value salt_value].include?(sym)
                        val
                      else
                        %w[1 true].include?(val)
                      end
        end
        @protection = prot unless prot.empty?
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing <dimension> element.
    class DimensionListener
      include REXML::SAX2Listener

      attr_reader :ref

      def initialize
        @ref = nil
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        @ref = attributes["ref"] if name == "dimension"
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing <sheetFormatPr> element.
    class SheetFormatListener
      include REXML::SAX2Listener

      attr_reader :properties

      def initialize
        @properties = {}
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        return unless name == "sheetFormatPr"

        drh = attributes["defaultRowHeight"]
        @properties[:default_row_height] = drh.to_f if drh
        dcw = attributes["defaultColWidth"]
        @properties[:default_col_width] = dcw.to_f if dcw
        bcw = attributes["baseColWidth"]
        @properties[:base_col_width] = bcw.to_i if bcw
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing <sheetViews><sheetView>, <pane>, and <selection>.
    class SheetViewListener
      include REXML::SAX2Listener

      attr_reader :view, :pane, :selection

      def initialize
        @view = {}
        @pane = nil
        @selection = nil
        @inside_sheet_views = false
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        case name
        when "sheetViews"
          @inside_sheet_views = true
        when "sheetView"
          return unless @inside_sheet_views

          sgl = attributes["showGridLines"]
          @view[:show_grid_lines] = %w[1 true].include?(sgl) unless sgl.nil?
          srch = attributes["showRowColHeaders"]
          @view[:show_row_col_headers] = %w[1 true].include?(srch) unless srch.nil?
          rtl = attributes["rightToLeft"]
          @view[:right_to_left] = %w[1 true].include?(rtl) unless rtl.nil?
          zs = attributes["zoomScale"]
          @view[:zoom_scale] = zs.to_i if zs
          ts = attributes["tabSelected"]
          @view[:tab_selected] = true if ts == "1"
        when "pane"
          return unless @inside_sheet_views

          ys = attributes["ySplit"]
          xs = attributes["xSplit"]
          frozen = attributes["state"] == "frozen"
          @pane = {
            row: ys ? ys.to_i : 0,
            col: xs ? xs.to_i : 0,
            state: frozen ? :frozen : :split
          }
        when "selection"
          return unless @inside_sheet_views

          ac = attributes["activeCell"]
          sq = attributes["sqref"]
          @selection = { active_cell: ac, sqref: sq } if ac || sq
        end
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        @inside_sheet_views = false if name == "sheetViews"
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing <dataValidations> elements.
    class DataValidationsListener
      include REXML::SAX2Listener

      attr_reader :validations

      def initialize
        @validations = []
        @current_dv = nil
        @inside_formula1 = false
        @inside_formula2 = false
        @text_buffer = +""
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        case name
        when "dataValidation"
          @current_dv = { sqref: attributes["sqref"] }
          @current_dv[:type] = attributes["type"] if attributes["type"]
          @current_dv[:operator] = attributes["operator"] if attributes["operator"]
          @current_dv[:error_style] = attributes["errorStyle"] if attributes["errorStyle"]
          @current_dv[:allow_blank] = true if attributes["allowBlank"] == "1"
          @current_dv[:show_input_message] = true if attributes["showInputMessage"] == "1"
          @current_dv[:show_error_message] = true if attributes["showErrorMessage"] == "1"
          @current_dv[:error_title] = xml_unescape(attributes["errorTitle"]) if attributes["errorTitle"]
          @current_dv[:error] = xml_unescape(attributes["error"]) if attributes["error"]
          @current_dv[:prompt_title] = xml_unescape(attributes["promptTitle"]) if attributes["promptTitle"]
          @current_dv[:prompt] = xml_unescape(attributes["prompt"]) if attributes["prompt"]
        when "formula1"
          @inside_formula1 = true
          @text_buffer = +""
        when "formula2"
          @inside_formula2 = true
          @text_buffer = +""
        end
      end

      def characters(text)
        @text_buffer << text if @inside_formula1 || @inside_formula2
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        case name
        when "formula1"
          @current_dv[:formula1] = @text_buffer.dup if @current_dv
          @inside_formula1 = false
        when "formula2"
          @current_dv[:formula2] = @text_buffer.dup if @current_dv
          @inside_formula2 = false
        when "dataValidation"
          @validations << @current_dv if @current_dv
          @current_dv = nil
        end
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end

      def xml_unescape(str)
        str.gsub("&amp;", "&").gsub("&lt;", "<").gsub("&gt;", ">").gsub("&quot;", '"').gsub("&apos;", "'")
      end
    end

    # SAX2 listener for parsing conditionalFormatting elements.
    class ConditionalFormattingListener
      include REXML::SAX2Listener

      attr_reader :rules

      def initialize
        @rules = []
        @current_sqref = nil
        @current_rule = nil
        @inside_formula = false
        @text_buffer = +""
        @cfvo_target = nil
        @color_target = nil
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        case name
        when "conditionalFormatting"
          @current_sqref = attributes["sqref"]
        when "cfRule"
          @current_rule = { sqref: @current_sqref, type: attributes["type"] }
          @current_rule[:priority] = attributes["priority"].to_i if attributes["priority"]
          @current_rule[:operator] = attributes["operator"] if attributes["operator"]
          @current_rule[:format_id] = attributes["dxfId"].to_i if attributes["dxfId"]
          @current_rule[:stop_if_true] = true if attributes["stopIfTrue"] == "1"
        when "formula"
          @inside_formula = true
          @text_buffer = +""
        when "colorScale"
          @current_rule[:color_scale] = { cfvo: [], colors: [] } if @current_rule
          @cfvo_target = :color_scale
          @color_target = :color_scale
        when "dataBar"
          if @current_rule
            db = { cfvo: [] }
            db[:min_length] = attributes["minLength"].to_i if attributes["minLength"]
            db[:max_length] = attributes["maxLength"].to_i if attributes["maxLength"]
            sv = attributes["showValue"]
            db[:show_value] = %w[1 true].include?(sv) unless sv.nil?
            @current_rule[:data_bar] = db
          end
          @cfvo_target = :data_bar
          @color_target = :data_bar
        when "iconSet"
          if @current_rule
            is = { cfvo: [] }
            is[:icon_set] = attributes["iconSet"] if attributes["iconSet"]
            rv = attributes["reverse"]
            is[:reverse] = %w[1 true].include?(rv) unless rv.nil?
            sv = attributes["showValue"]
            is[:show_value] = %w[1 true].include?(sv) unless sv.nil?
            @current_rule[:icon_set] = is
          end
          @cfvo_target = :icon_set
        when "cfvo"
          cfvo = { type: attributes["type"] }
          cfvo[:val] = attributes["val"] if attributes["val"]
          append_cfvo(cfvo)
        when "color"
          append_color(attributes["rgb"]) if attributes["rgb"]
        end
      end

      def characters(text)
        @text_buffer << text if @inside_formula
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        case name
        when "formula"
          (@current_rule[:formulas] ||= []) << @text_buffer.dup if @current_rule
          @inside_formula = false
        when "cfRule"
          @rules << @current_rule if @current_rule
          @current_rule = nil
        when "conditionalFormatting"
          @current_sqref = nil
        when "colorScale", "dataBar", "iconSet"
          @cfvo_target = nil
          @color_target = nil
        end
      end

      private

      def append_cfvo(cfvo)
        return unless @current_rule && @cfvo_target

        container = @current_rule[@cfvo_target]
        container[:cfvo] << cfvo if container
      end

      def append_color(rgb)
        return unless @current_rule && @color_target

        container = @current_rule[@color_target]
        if container.is_a?(Hash) && container.key?(:colors)
          container[:colors] << rgb
        elsif container.is_a?(Hash)
          container[:color] = rgb
        end
      end

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing print/page elements from worksheet XML.
    class PrintPageListener
      include REXML::SAX2Listener

      attr_reader :print_options, :page_margins, :page_setup, :header_footer, :row_breaks, :col_breaks

      def initialize
        @print_options = {}
        @page_margins = nil
        @page_setup = {}
        @header_footer = {}
        @row_breaks = []
        @col_breaks = []
        @inside_header_footer = false
        @inside_row_breaks = false
        @inside_col_breaks = false
        @current_hf_field = nil
        @text_buffer = +""
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        case name
        when "printOptions"
          @print_options[:grid_lines] = true if attributes["gridLines"] == "1"
          @print_options[:headings] = true if attributes["headings"] == "1"
          @print_options[:horizontal_centered] = true if attributes["horizontalCentered"] == "1"
          @print_options[:vertical_centered] = true if attributes["verticalCentered"] == "1"
        when "pageMargins"
          m = {}
          %w[left right top bottom header footer].each do |k|
            v = attributes[k]
            m[k.to_sym] = v.to_f if v
          end
          @page_margins = m unless m.empty?
        when "pageSetup"
          o = attributes["orientation"]
          @page_setup[:orientation] = o if o
          ps = attributes["paperSize"]
          @page_setup[:paper_size] = ps.to_i if ps
          sc = attributes["scale"]
          @page_setup[:scale] = sc.to_i if sc
          ftw = attributes["fitToWidth"]
          @page_setup[:fit_to_width] = ftw.to_i if ftw
          fth = attributes["fitToHeight"]
          @page_setup[:fit_to_height] = fth.to_i if fth
        when "headerFooter"
          @inside_header_footer = true
        when "oddHeader", "oddFooter", "evenHeader", "evenFooter"
          if @inside_header_footer
            @current_hf_field = name
            @text_buffer = +""
          end
        when "rowBreaks"
          @inside_row_breaks = true
        when "colBreaks"
          @inside_col_breaks = true
        when "brk"
          id = attributes["id"]&.to_i
          if id
            @row_breaks << id if @inside_row_breaks
            @col_breaks << id if @inside_col_breaks
          end
        end
      end

      def characters(text)
        @text_buffer << text if @current_hf_field
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        case name
        when "headerFooter"
          @inside_header_footer = false
        when "oddHeader"
          @header_footer[:odd_header] = @text_buffer.dup if @current_hf_field == "oddHeader"
          @current_hf_field = nil
        when "oddFooter"
          @header_footer[:odd_footer] = @text_buffer.dup if @current_hf_field == "oddFooter"
          @current_hf_field = nil
        when "evenHeader"
          @header_footer[:even_header] = @text_buffer.dup if @current_hf_field == "evenHeader"
          @current_hf_field = nil
        when "evenFooter"
          @header_footer[:even_footer] = @text_buffer.dup if @current_hf_field == "evenFooter"
          @current_hf_field = nil
        when "rowBreaks"
          @inside_row_breaks = false
        when "colBreaks"
          @inside_col_breaks = false
        end
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing table XML.
    class TableListener
      include REXML::SAX2Listener

      attr_reader :table

      def initialize
        @table = nil
        @columns = []
        @current_column = nil
        @inside_calc_formula = false
        @text_buffer = +""
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        case name
        when "table"
          @table = {
            id: attributes["id"]&.to_i,
            name: attributes["name"],
            display_name: attributes["displayName"],
            ref: attributes["ref"]
          }
          trc = attributes["totalsRowCount"]
          @table[:totals_row_count] = trc.to_i if trc
        when "tableColumn"
          col = { name: attributes["name"] }
          trf = attributes["totalsRowFunction"]
          col[:totals_row_function] = trf if trf
          @current_column = col
        when "calculatedColumnFormula"
          @inside_calc_formula = true
          @text_buffer = +""
        when "tableStyleInfo"
          if @table
            si = {}
            si[:name] = attributes["name"] if attributes["name"]
            sfc = attributes["showFirstColumn"]
            si[:show_first_column] = sfc == "1" unless sfc.nil?
            slc = attributes["showLastColumn"]
            si[:show_last_column] = slc == "1" unless slc.nil?
            srs = attributes["showRowStripes"]
            si[:show_row_stripes] = srs == "1" unless srs.nil?
            scs = attributes["showColumnStripes"]
            si[:show_column_stripes] = scs == "1" unless scs.nil?
            @table[:style] = si
          end
        end
      end

      def characters(text)
        @text_buffer << text if @inside_calc_formula
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        case name
        when "calculatedColumnFormula"
          @current_column[:calculated_column_formula] = @text_buffer.dup if @current_column
          @inside_calc_formula = false
        when "tableColumn"
          @columns << @current_column if @current_column
          @current_column = nil
        when "table"
          @table[:columns] = @columns if @table
        end
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing calcChain.xml.
    class CalcChainListener
      include REXML::SAX2Listener

      attr_reader :entries

      def initialize
        @entries = []
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        return unless name == "c"

        entry = {}
        entry[:ref] = attributes["r"] if attributes["r"]
        i = attributes["i"]
        entry[:sheet_id] = i.to_i if i
        @entries << entry
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing drawing XML to extract image information.
    class DrawingImagesListener
      include REXML::SAX2Listener

      attr_reader :images

      def initialize
        @images = []
        @current_image = nil
        @inside_anchor = false
        @inside_pic = false
        @inside_from = false
        @inside_to = false
        @current_field = nil
        @text_buffer = +""
        @anchor_from = {}
        @anchor_to = {}
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        case name
        when "twoCellAnchor", "oneCellAnchor"
          @inside_anchor = true
          @anchor_from = {}
          @anchor_to = {}
        when "pic"
          @inside_pic = true
          @current_image = {}
        when "cNvPr"
          if @inside_pic && @current_image
            @current_image[:name] = attributes["name"] if attributes["name"]
            @current_image[:id] = attributes["id"]&.to_i
          end
        when "blip"
          rid = attributes["r:embed"] || attributes["embed"]
          @current_image[:embed_rid] = rid if @inside_pic && @current_image && rid
        when "from"
          @inside_from = true if @inside_anchor
        when "to"
          @inside_to = true if @inside_anchor
        when "ext"
          if @inside_pic && @current_image
            cx = attributes["cx"]
            cy = attributes["cy"]
            @current_image[:cx] = cx.to_i if cx
            @current_image[:cy] = cy.to_i if cy
          end
        when "col", "colOff", "row", "rowOff"
          @current_field = name
          @text_buffer = +""
        end
      end

      def characters(text)
        @text_buffer << text if @current_field
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        case name
        when "pic"
          if @current_image && !@current_image.empty?
            @anchor_from.each { |k, v| @current_image[:"from_#{k}"] = v }
            @anchor_to.each { |k, v| @current_image[:"to_#{k}"] = v }
            @images << @current_image
          end
          @current_image = nil
          @inside_pic = false
        when "twoCellAnchor", "oneCellAnchor"
          @inside_anchor = false
          @anchor_from = {}
          @anchor_to = {}
        when "from"
          @inside_from = false
        when "to"
          @inside_to = false
        when "col", "colOff", "row", "rowOff"
          if @current_field
            val = @text_buffer.to_i
            if @inside_from
              @anchor_from[@current_field] = val
            elsif @inside_to
              @anchor_to[@current_field] = val
            end
          end
          @current_field = nil
        end
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing drawing XML to extract chart references.
    class DrawingChartsListener
      include REXML::SAX2Listener

      attr_reader :charts

      def initialize
        @charts = []
        @inside_graphic_frame = false
        @current_chart = nil
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        case name
        when "graphicFrame"
          @inside_graphic_frame = true
          @current_chart = {}
        when "cNvPr"
          @current_chart[:name] = attributes["name"] if @inside_graphic_frame && @current_chart && attributes["name"]
        when "chart"
          rid = attributes["r:id"] || attributes["id"]
          @current_chart[:rid] = rid if @inside_graphic_frame && @current_chart && rid
        end
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        if name == "graphicFrame"
          @charts << @current_chart if @current_chart && @current_chart[:rid]
          @current_chart = nil
          @inside_graphic_frame = false
        end
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing drawing XML to extract shape elements.
    class DrawingShapesListener
      include REXML::SAX2Listener

      attr_reader :shapes

      def initialize
        @shapes = []
        @inside_anchor = false
        @inside_sp = false
        @current_shape = nil
        @inside_from = false
        @inside_to = false
        @inside_tx_body = false
        @inside_t = false
        @current_field = nil
        @text_buffer = +""
        @anchor_from = {}
        @anchor_to = {}
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        case name
        when "twoCellAnchor", "oneCellAnchor"
          @inside_anchor = true
          @anchor_from = {}
          @anchor_to = {}
        when "sp"
          @inside_sp = true
          @current_shape = {}
        when "cNvPr"
          if @inside_sp && @current_shape
            @current_shape[:name] = attributes["name"] if attributes["name"]
            @current_shape[:id] = attributes["id"]&.to_i
          end
        when "prstGeom"
          @current_shape[:preset] = attributes["prst"] if @inside_sp && @current_shape && attributes["prst"]
        when "from"
          @inside_from = true if @inside_anchor
        when "to"
          @inside_to = true if @inside_anchor
        when "txBody"
          @inside_tx_body = true if @inside_sp
        when "t"
          @inside_t = true if @inside_tx_body
          @text_buffer = +""
        when "col", "colOff", "row", "rowOff"
          @current_field = name
          @text_buffer = +""
        end
      end

      def characters(text)
        @text_buffer << text if @current_field || @inside_t
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        case name
        when "sp"
          if @current_shape && !@current_shape.empty?
            @anchor_from.each { |k, v| @current_shape[:"from_#{k}"] = v }
            @anchor_to.each { |k, v| @current_shape[:"to_#{k}"] = v }
            @shapes << @current_shape
          end
          @current_shape = nil
          @inside_sp = false
          @inside_tx_body = false
        when "twoCellAnchor", "oneCellAnchor"
          @inside_anchor = false
          @anchor_from = {}
          @anchor_to = {}
        when "from"
          @inside_from = false
        when "to"
          @inside_to = false
        when "txBody"
          @inside_tx_body = false
        when "t"
          if @inside_t && @inside_tx_body && @current_shape
            @current_shape[:text] = (@current_shape[:text] || +"") << @text_buffer
          end
          @inside_t = false
        when "col", "colOff", "row", "rowOff"
          if @current_field
            val = @text_buffer.to_i
            if @inside_from
              @anchor_from[@current_field] = val
            elsif @inside_to
              @anchor_to[@current_field] = val
            end
          end
          @current_field = nil
        end
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing chart XML to identify chart type and title.
    class ChartTypeListener
      include REXML::SAX2Listener

      attr_reader :chart_type, :title, :series, :legend, :data_labels, :cat_axis_title, :val_axis_title

      CHART_TYPES = %w[barChart lineChart pieChart areaChart scatterChart doughnutChart radarChart
                       bar3DChart line3DChart pie3DChart area3DChart surfaceChart stockChart bubbleChart].freeze

      def initialize
        @chart_type = nil
        @title = nil
        @series = []
        @legend = {}
        @data_labels = {}
        @cat_axis_title = nil
        @val_axis_title = nil
        @inside_title = false
        @inside_t = false
        @text_buffer = +""
        @inside_ser = false
        @current_ser = nil
        @inside_cat = false
        @inside_val = false
        @inside_f = false
        @inside_legend = false
        @inside_dlbls = false
        @inside_cat_ax = false
        @inside_val_ax = false
        @inside_ax_title = false
        @title_depth = 0
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        if CHART_TYPES.include?(name)
          @chart_type = name
        end

        case name
        when "ser"
          @inside_ser = true
          @current_ser = {}
        when "cat"
          @inside_cat = true if @inside_ser
        when "val"
          @inside_val = true if @inside_ser
        when "f"
          @inside_f = true
          @text_buffer = +""
        when "title"
          @title_depth += 1
          if @inside_cat_ax || @inside_val_ax
            @inside_ax_title = true
          elsif @title_depth == 1
            @inside_title = true
          end
        when "t"
          @inside_t = true
          @text_buffer = +""
        when "legend"
          @inside_legend = true
        when "legendPos"
          @legend[:position] = attributes["val"] if @inside_legend && attributes["val"]
        when "dLbls"
          @inside_dlbls = true if @inside_ser || @chart_type
        when "showVal"
          @data_labels[:show_val] = attributes["val"] == "1" if @inside_dlbls
        when "showCatName"
          @data_labels[:show_cat_name] = attributes["val"] == "1" if @inside_dlbls
        when "showSerName"
          @data_labels[:show_ser_name] = attributes["val"] == "1" if @inside_dlbls
        when "showPercent"
          @data_labels[:show_percent] = attributes["val"] == "1" if @inside_dlbls
        when "showLegendKey"
          @data_labels[:show_legend_key] = attributes["val"] == "1" if @inside_dlbls
        when "catAx"
          @inside_cat_ax = true
        when "valAx"
          @inside_val_ax = true
        end
      end

      def characters(text)
        @text_buffer << text if @inside_t || @inside_f
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        case name
        when "t"
          if @inside_ax_title
            if @inside_cat_ax
              @cat_axis_title = @text_buffer.dup
            elsif @inside_val_ax
              @val_axis_title = @text_buffer.dup
            end
          elsif @inside_title && @title_depth == 1
            @title = @text_buffer.dup
          end
          @inside_t = false
        when "f"
          if @inside_ser
            if @inside_cat
              @current_ser[:cat_ref] = @text_buffer.dup
            elsif @inside_val
              @current_ser[:val_ref] = @text_buffer.dup
            else
              @current_ser[:name] = @text_buffer.dup
            end
          end
          @inside_f = false
        when "cat"
          @inside_cat = false
        when "val"
          @inside_val = false
        when "ser"
          @series << @current_ser if @current_ser
          @current_ser = nil
          @inside_ser = false
        when "title"
          @title_depth -= 1
          @inside_title = false if @title_depth.zero?
          @inside_ax_title = false
        when "legend"
          @inside_legend = false
        when "dLbls"
          @inside_dlbls = false
        when "catAx"
          @inside_cat_ax = false
        when "valAx"
          @inside_val_ax = false
        end
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing comments XML.
    class CommentsListener
      include REXML::SAX2Listener

      attr_reader :comments

      def initialize
        @comments = []
        @authors = []
        @inside_authors = false
        @inside_author = false
        @inside_comment = false
        @inside_text = false
        @inside_t = false
        @current_comment = nil
        @text_buffer = +""
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        case name
        when "authors"
          @inside_authors = true
        when "author"
          @inside_author = true
          @text_buffer = +""
        when "comment"
          @inside_comment = true
          @current_comment = { ref: attributes["ref"], author_id: attributes["authorId"]&.to_i }
        when "text"
          @inside_text = true if @inside_comment
          @text_buffer = +""
        when "t"
          @inside_t = true if @inside_text
        end
      end

      def characters(text)
        @text_buffer << text if @inside_author || @inside_t
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        case name
        when "authors"
          @inside_authors = false
        when "author"
          @authors << @text_buffer.dup if @inside_authors
          @inside_author = false
        when "comment"
          if @current_comment
            aid = @current_comment[:author_id]
            @current_comment[:author] = @authors[aid] if aid && aid < @authors.size
            @current_comment.delete(:author_id)
            @comments << @current_comment
          end
          @inside_comment = false
          @current_comment = nil
        when "text"
          @current_comment[:text] = @text_buffer.dup if @current_comment && @inside_text
          @inside_text = false
        when "t"
          @inside_t = false
        end
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing pivotTable XML.
    class PivotTableListener
      include REXML::SAX2Listener

      attr_reader :pivot_table

      def initialize
        @pivot_table = nil
        @fields = []
        @row_fields = []
        @col_fields = []
        @data_fields = []
        @inside_row_fields = false
        @inside_col_fields = false
        @inside_data_fields = false
        @inside_pivot_field = false
        @inside_items = false
        @current_items = []
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        case name
        when "pivotTableDefinition"
          @pivot_table = {
            name: attributes["name"],
            cache_id: attributes["cacheId"]&.to_i
          }
        when "location"
          @pivot_table[:ref] = attributes["ref"] if @pivot_table
        when "pivotField"
          @inside_pivot_field = true
          @current_field = {}
          @current_field[:axis] = attributes["axis"] if attributes["axis"]
          @current_field[:data_field] = true if attributes["dataField"] == "1"
          @current_field[:name] = attributes["name"] if attributes["name"]
          @current_items = []
        when "items"
          @inside_items = true if @inside_pivot_field
        when "item"
          if @inside_items
            item_type = attributes["t"]
            item_x = attributes["x"]&.to_i
            @current_items << { x: item_x, t: item_type } if item_type || item_x
          end
        when "rowFields"
          @inside_row_fields = true
        when "colFields"
          @inside_col_fields = true
        when "dataFields"
          @inside_data_fields = true
        when "field"
          idx = attributes["x"]&.to_i
          @row_fields << idx if @inside_row_fields && idx
          @col_fields << idx if @inside_col_fields && idx
        when "dataField"
          if @inside_data_fields
            @data_fields << {
              name: attributes["name"],
              fld: attributes["fld"]&.to_i,
              subtotal: attributes["subtotal"] || "sum"
            }
          end
        end
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        case name
        when "pivotField"
          @current_field[:items] = @current_items unless @current_items.empty?
          @fields << @current_field
          @inside_pivot_field = false
        when "items"
          @inside_items = false
        when "pivotTableDefinition"
          if @pivot_table
            @pivot_table[:fields] = @fields
            @pivot_table[:row_fields] = @row_fields
            @pivot_table[:col_fields] = @col_fields
            @pivot_table[:data_fields] = @data_fields
          end
        when "rowFields"
          @inside_row_fields = false
        when "colFields"
          @inside_col_fields = false
        when "dataFields"
          @inside_data_fields = false
        end
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end
  end
end
