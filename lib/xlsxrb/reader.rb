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
        entry = {}
        if link[:rid]
          url = rid_to_url[link[:rid]]
          entry[:url] = url if url
        end
        entry[:display] = link[:display] if link[:display]
        entry[:tooltip] = link[:tooltip] if link[:tooltip]
        entry[:location] = link[:location] if link[:location]
        result[link[:ref]] = entry unless entry.empty?
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
        xf = resolve_effective_xf(styles[:cell_xfs][xf_index], styles[:cell_style_xfs])
        next unless xf

        fmt_id = xf[:num_fmt_id]
        next unless fmt_id && fmt_id != 0

        format_code = resolve_num_fmt_code(fmt_id, styles[:num_fmts])
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
        xf = resolve_effective_xf(styles[:cell_xfs][xf_index], styles[:cell_style_xfs])
        next unless xf

        entry = {}
        entry[:font] = styles[:fonts][xf[:font_id]] if xf[:font_id]&.positive? && styles[:fonts][xf[:font_id]]
        entry[:fill] = styles[:fills][xf[:fill_id]] if xf[:fill_id]&.positive? && styles[:fills][xf[:fill_id]]
        entry[:border] = styles[:borders][xf[:border_id]] if xf[:border_id]&.positive? && styles[:borders][xf[:border_id]]
        if xf[:num_fmt_id]&.positive?
          code = resolve_num_fmt_code(xf[:num_fmt_id], styles[:num_fmts])
          entry[:num_fmt] = code if code
        end
        entry[:alignment] = xf[:alignment] if xf[:alignment]
        entry[:protection] = xf[:protection] if xf[:protection]
        entry[:quote_prefix] = true if xf[:quote_prefix]
        entry[:pivot_button] = true if xf[:pivot_button]
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

    # Returns array of font entries from the styles.
    def fonts
      styles = load_styles
      return [] if styles.empty?

      styles[:fonts] || []
    end

    # Returns array of fill entries from the styles.
    def fills
      styles = load_styles
      return [] if styles.empty?

      styles[:fills] || []
    end

    # Returns array of border entries from the styles.
    def borders
      styles = load_styles
      return [] if styles.empty?

      styles[:borders] || []
    end

    # Returns custom number formats as { numFmtId => formatCode }.
    def num_fmts
      styles = load_styles
      return {} if styles.empty?

      styles[:num_fmts] || {}
    end

    # Returns indexed colors palette (array of ARGB hex strings).
    def indexed_colors
      styles = load_styles
      return [] if styles.empty?

      styles[:indexed_colors] || []
    end

    # Returns MRU (most recently used) colors (array of color hashes).
    def mru_colors
      styles = load_styles
      return [] if styles.empty?

      styles[:mru_colors] || []
    end

    # Returns table styles configuration hash.
    def table_styles
      styles = load_styles
      return {} if styles.empty?

      styles[:table_styles] || {}
    end

    # Returns array of cellStyleXfs entries (base style format definitions).
    def cell_style_xfs
      styles = load_styles
      return [] if styles.empty?

      styles[:cell_style_xfs] || []
    end

    # Returns array of cellXfs entries (cell format definitions).
    def cell_xfs
      styles = load_styles
      return [] if styles.empty?

      styles[:cell_xfs] || []
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

    # Returns data validations container options (disablePrompts, xWindow, yWindow).
    def data_validations_options(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return {} if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_data_validations_options(worksheet_xml)
    end

    # Returns conditional formatting rules for the given sheet.
    def conditional_formats(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return [] if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_conditional_formats(worksheet_xml)
    end

    # Returns sheet-level properties (tabColor, outlinePr) for the given sheet.
    # Returns sheet-level properties (tabColor, outlinePr) for the given sheet.
    def sheet_properties(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return {} if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_properties(worksheet_xml)
    end

    # Returns phonetic properties for the given sheet, or nil if not present.
    def phonetic_properties(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return nil if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_phonetic_pr(worksheet_xml)
    end

    # Returns sheet protection settings as a hash, or nil if unprotected.
    def sheet_protection(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return nil if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_sheet_protection(worksheet_xml)
    end

    # Returns protected ranges for the given sheet as an array of hashes.
    def protected_ranges(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return [] if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_protected_ranges(worksheet_xml)
    end

    # Returns cell watches for the given sheet as an array of cell references.
    def cell_watches(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return [] if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_cell_watches(worksheet_xml)
    end

    # Returns ignored errors for the given sheet.
    def ignored_errors(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return [] if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_ignored_errors(worksheet_xml)
    end

    # Returns data consolidation settings for the given sheet.
    def data_consolidate(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return nil if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_data_consolidate(worksheet_xml)
    end

    # Returns scenarios for the given sheet.
    def scenarios(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return nil if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_scenarios(worksheet_xml)
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

    # Returns custom document properties as an array of { name:, value:, type: }.
    def custom_properties
      xml = extract_zip_entry("docProps/custom.xml")
      if xml.nil? || xml.empty?
        rels_xml = extract_zip_entry("_rels/.rels")
        return [] if rels_xml.nil? || rels_xml.empty?

        rels = parse_rels_with_types(rels_xml)
        custom_rel = rels.find { |r| r[:type]&.end_with?("/custom-properties") }
        return [] unless custom_rel

        target = custom_rel[:target]
        entry_path = target.start_with?("/") ? target.delete_prefix("/") : target
        xml = extract_zip_entry(entry_path)
      end
      return [] if xml.nil? || xml.empty?

      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = CustomPropertiesListener.new
      parser.listen(listener)
      parser.parse
      listener.properties
    end

    # Returns workbook properties (e.g. { date1904: false, default_theme_version: 166925 }).
    def workbook_properties
      parse_workbook_metadata[:workbook_properties]
    end

    # Returns the workbook conformance class ("transitional" or "strict"), or nil if not set.
    def conformance
      parse_workbook_metadata[:conformance]
    end

    # Returns file version properties (e.g. { app_name: "xl", last_edited: "7" }).
    def file_version
      parse_workbook_metadata[:file_version]
    end

    # Returns file sharing properties (e.g. { read_only_recommended: true, user_name: "John" }).
    def file_sharing
      parse_workbook_metadata[:file_sharing]
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

    # Returns file recovery properties hash.
    def file_recovery_properties
      parse_workbook_metadata[:file_recovery_properties]
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

    # Returns the print area for the given sheet, or nil if not set.
    def print_area(sheet: nil)
      _sheet_name, idx = resolve_sheet_for_defined_name(sheet)
      dn = defined_names.find { |d| d[:name] == "_xlnm.Print_Area" && d[:local_sheet_id] == idx }
      return nil unless dn

      # Strip the sheet prefix (e.g. "'Sheet1'!$A$1:$D$20" → "$A$1:$D$20")
      dn[:value]&.sub(/\A'[^']*'!/, "")
    end

    # Returns the print titles for the given sheet, or nil if not set.
    def print_titles(sheet: nil)
      _sheet_name, idx = resolve_sheet_for_defined_name(sheet)
      dn = defined_names.find { |d| d[:name] == "_xlnm.Print_Titles" && d[:local_sheet_id] == idx }
      return nil unless dn

      dn[:value]
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
    def macros?
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
        chart[:grouping] = cl.grouping if cl.grouping
        chart[:bar_dir] = cl.bar_dir if cl.bar_dir
        chart[:vary_colors] = cl.vary_colors unless cl.vary_colors.nil?
        chart[:plot_vis_only] = cl.plot_vis_only unless cl.plot_vis_only.nil?
        chart[:disp_blanks_as] = cl.disp_blanks_as if cl.disp_blanks_as
        chart[:style] = cl.style if cl.style
        chart[:auto_title_deleted] = cl.auto_title_deleted unless cl.auto_title_deleted.nil?
        chart[:rounded_corners] = cl.rounded_corners unless cl.rounded_corners.nil?
        chart[:cat_axis_tick_lbl_pos] = cl.cat_axis_tick_lbl_pos if cl.cat_axis_tick_lbl_pos
        chart[:val_axis_tick_lbl_pos] = cl.val_axis_tick_lbl_pos if cl.val_axis_tick_lbl_pos
        chart[:cat_axis_major_gridlines] = cl.cat_axis_major_gridlines if cl.cat_axis_major_gridlines
        chart[:val_axis_major_gridlines] = cl.val_axis_major_gridlines if cl.val_axis_major_gridlines
        chart[:cat_axis_minor_gridlines] = cl.cat_axis_minor_gridlines if cl.cat_axis_minor_gridlines
        chart[:val_axis_minor_gridlines] = cl.val_axis_minor_gridlines if cl.val_axis_minor_gridlines
        chart[:show_d_lbls_over_max] = cl.show_d_lbls_over_max unless cl.show_d_lbls_over_max.nil?
        chart[:cat_axis_delete] = cl.cat_axis_delete unless cl.cat_axis_delete.nil?
        chart[:val_axis_delete] = cl.val_axis_delete unless cl.val_axis_delete.nil?
        chart[:cat_axis_orientation] = cl.cat_axis_orientation if cl.cat_axis_orientation
        chart[:val_axis_orientation] = cl.val_axis_orientation if cl.val_axis_orientation
        chart[:gap_width] = cl.gap_width if cl.gap_width
        chart[:overlap] = cl.overlap if cl.overlap
        chart[:gap_depth] = cl.gap_depth if cl.gap_depth
        chart[:bar_shape] = cl.bar_shape if cl.bar_shape
        chart[:bubble_3d] = cl.bubble_3d unless cl.bubble_3d.nil?
        chart[:bubble_scale] = cl.bubble_scale if cl.bubble_scale
        chart[:show_neg_bubbles] = cl.show_neg_bubbles unless cl.show_neg_bubbles.nil?
        chart[:size_represents] = cl.size_represents if cl.size_represents
        chart[:view_3d] = cl.view_3d if cl.view_3d
        chart[:cat_axis_num_fmt] = cl.cat_axis_num_fmt if cl.cat_axis_num_fmt
        chart[:val_axis_num_fmt] = cl.val_axis_num_fmt if cl.val_axis_num_fmt
        chart[:cat_axis_major_tick_mark] = cl.cat_axis_major_tick_mark if cl.cat_axis_major_tick_mark
        chart[:cat_axis_minor_tick_mark] = cl.cat_axis_minor_tick_mark if cl.cat_axis_minor_tick_mark
        chart[:val_axis_major_tick_mark] = cl.val_axis_major_tick_mark if cl.val_axis_major_tick_mark
        chart[:val_axis_minor_tick_mark] = cl.val_axis_minor_tick_mark if cl.val_axis_minor_tick_mark
        chart[:cat_axis_crosses] = cl.cat_axis_crosses if cl.cat_axis_crosses
        chart[:val_axis_crosses] = cl.val_axis_crosses if cl.val_axis_crosses
        chart[:cat_axis_crosses_at] = cl.cat_axis_crosses_at if cl.cat_axis_crosses_at
        chart[:val_axis_crosses_at] = cl.val_axis_crosses_at if cl.val_axis_crosses_at
        chart[:cat_axis_tick_lbl_skip] = cl.cat_axis_tick_lbl_skip if cl.cat_axis_tick_lbl_skip
        chart[:cat_axis_tick_mark_skip] = cl.cat_axis_tick_mark_skip if cl.cat_axis_tick_mark_skip
        chart[:cat_axis_lbl_offset] = cl.cat_axis_lbl_offset if cl.cat_axis_lbl_offset
        chart[:cat_axis_no_multi_lvl_lbl] = cl.cat_axis_no_multi_lvl_lbl unless cl.cat_axis_no_multi_lvl_lbl.nil?
        chart[:val_axis_cross_between] = cl.val_axis_cross_between if cl.val_axis_cross_between
        chart[:val_axis_major_unit] = cl.val_axis_major_unit if cl.val_axis_major_unit
        chart[:val_axis_minor_unit] = cl.val_axis_minor_unit if cl.val_axis_minor_unit
        chart[:val_axis_disp_units] = cl.val_axis_disp_units if cl.val_axis_disp_units
        chart[:cat_axis_scaling_max] = cl.cat_axis_scaling_max if cl.cat_axis_scaling_max
        chart[:cat_axis_scaling_min] = cl.cat_axis_scaling_min if cl.cat_axis_scaling_min
        chart[:val_axis_scaling_max] = cl.val_axis_scaling_max if cl.val_axis_scaling_max
        chart[:val_axis_scaling_min] = cl.val_axis_scaling_min if cl.val_axis_scaling_min
        chart[:cat_axis_log_base] = cl.cat_axis_log_base if cl.cat_axis_log_base
        chart[:val_axis_log_base] = cl.val_axis_log_base if cl.val_axis_log_base
        chart[:first_slice_ang] = cl.first_slice_ang if cl.first_slice_ang
        chart[:hole_size] = cl.hole_size if cl.hole_size
        chart[:smooth] = cl.smooth unless cl.smooth.nil?
        chart[:marker] = cl.marker unless cl.marker.nil?
        chart[:scatter_style] = cl.scatter_style if cl.scatter_style
        chart[:radar_style] = cl.radar_style if cl.radar_style
        chart[:cat_axis_pos] = cl.cat_axis_pos if cl.cat_axis_pos
        chart[:val_axis_pos] = cl.val_axis_pos if cl.val_axis_pos
        chart[:wireframe] = cl.wireframe unless cl.wireframe.nil?
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
    # Each hash: { name:, ref:, cache_id:, fields:, row_fields:, col_fields:, data_fields:, cache: }
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
        pt = listener.pivot_table
        next unless pt

        # Resolve pivotCacheDefinition via pivot table rels.
        cache_info = load_pivot_cache_definition(path)
        pt[:cache] = cache_info if cache_info
        pt
      end
    end

    # Returns external links from the workbook as an array of hashes.
    # Each hash: { target:, sheet_names: [] }
    def external_links
      wb_rels_xml = extract_zip_entry("xl/_rels/workbook.xml.rels")
      return [] if wb_rels_xml.nil? || wb_rels_xml.empty?

      # Find external link rels.
      el_targets = []
      wb_rels_xml.scan(/<Relationship\s[^>]*>/) do |rel_tag|
        next unless rel_tag.include?("/externalLink")

        target = rel_tag[/Target="([^"]+)"/, 1]
        el_targets << target if target
      end
      return [] if el_targets.empty?

      el_targets.filter_map do |target|
        path = target.start_with?("/") ? target[1..] : "xl/#{target}"
        xml = extract_zip_entry(path)
        next if xml.nil? || xml.empty?

        parser = REXML::Parsers::SAX2Parser.new(xml)
        listener = ExternalLinkListener.new
        parser.listen(listener)
        parser.parse

        # Resolve the external book target from rels.
        rels_path = path.sub(%r{([^/]+)\.xml$}, '_rels/\1.xml.rels')
        rels_xml = extract_zip_entry(rels_path)
        ext_target = nil
        rels_xml&.scan(/<Relationship[^>]+Target="([^"]+)"/) { |t,| ext_target = t }

        { target: ext_target, sheet_names: listener.sheet_names }
      end
    end

    STRICT_SSML_NS = "http://purl.oclc.org/ooxml/spreadsheetml/main/2006/main"
    TRANSITIONAL_SSML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

    # Returns :strict or :transitional based on the namespace of the workbook XML.
    def format_variant
      workbook_xml = extract_zip_entry("xl/workbook.xml")
      return :transitional if workbook_xml.nil? || workbook_xml.empty?

      if workbook_xml.include?(STRICT_SSML_NS)
        :strict
      else
        :transitional
      end
    end

    private

    def resolve_sheet_for_defined_name(sheet)
      sheets = discover_sheets
      if sheet
        idx = if sheet.is_a?(Integer)
                sheet
              else
                sheets.index { |s| s[:name] == sheet }
              end
        [sheets[idx][:name], idx]
      else
        [sheets.first[:name], 0]
      end
    end

    def parse_workbook_metadata
      workbook_xml = extract_zip_entry("xl/workbook.xml")
      if workbook_xml.nil? || workbook_xml.empty?
        return { workbook_properties: {}, workbook_views: {}, calc_properties: {}, file_recovery_properties: {},
                 workbook_protection: nil }
      end

      parser = REXML::Parsers::SAX2Parser.new(workbook_xml)
      listener = WorkbookListener.new
      parser.listen(listener)
      parser.parse
      {
        workbook_properties: listener.workbook_properties,
        workbook_views: listener.workbook_views,
        calc_properties: listener.calc_properties,
        defined_names: listener.defined_names,
        workbook_protection: listener.workbook_protection,
        file_version: listener.file_version,
        file_sharing: listener.file_sharing,
        conformance: listener.conformance,
        file_recovery_properties: listener.file_recovery_properties
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

    def load_pivot_cache_definition(pivot_table_path)
      rels_path = pivot_table_path.sub(%r{([^/]+)$}, '_rels/\1.rels')
      rels_xml = extract_zip_entry(rels_path)
      return nil if rels_xml.nil? || rels_xml.empty?

      rels = parse_rels_with_types(rels_xml)
      cache_rel = rels.find { |r| r[:type]&.end_with?("/pivotCacheDefinition") }
      return nil unless cache_rel

      cache_path = normalize_xl_path(cache_rel[:target], File.dirname(pivot_table_path))
      cache_xml = extract_zip_entry(cache_path)
      return nil if cache_xml.nil? || cache_xml.empty?

      parser = REXML::Parsers::SAX2Parser.new(cache_xml)
      listener = PivotCacheDefinitionListener.new
      parser.listen(listener)
      parser.parse
      cache_def = listener.cache_definition

      # Load pivotCacheRecords via cache definition rels.
      cache_rels_path = cache_path.sub(%r{([^/]+)$}, '_rels/\1.rels')
      cache_rels_xml = extract_zip_entry(cache_rels_path)
      if cache_rels_xml && !cache_rels_xml.empty?
        cache_rels = parse_rels_with_types(cache_rels_xml)
        rec_rel = cache_rels.find { |r| r[:type]&.end_with?("/pivotCacheRecords") }
        if rec_rel
          rec_path = normalize_xl_path(rec_rel[:target], File.dirname(cache_path))
          rec_xml = extract_zip_entry(rec_path)
          if rec_xml && !rec_xml.empty?
            rec_parser = REXML::Parsers::SAX2Parser.new(rec_xml)
            rec_listener = PivotCacheRecordsListener.new
            rec_parser.listen(rec_listener)
            rec_parser.parse
            cache_def[:records] = rec_listener.records unless rec_listener.records.empty?
          end
        end
      end

      cache_def
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
        borders: listener.borders, dxfs: listener.dxfs,
        indexed_colors: listener.indexed_colors, mru_colors: listener.mru_colors,
        table_styles: listener.table_styles
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

        xf = resolve_effective_xf(styles[:cell_xfs][xf_index], styles[:cell_style_xfs])
        next unless xf

        fmt_id = xf[:num_fmt_id]
        next unless date_format?(fmt_id, styles[:num_fmts])

        raw_cells[cell_ref] = if value.is_a?(Float) && (value % 1).positive?
                                Xlsxrb.serial_to_datetime(value)
                              else
                                Xlsxrb.serial_to_date(value.to_i)
                              end
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

    def resolve_num_fmt_code(fmt_id, custom_num_fmts)
      custom_num_fmts[fmt_id] || Xlsxrb::BUILTIN_NUM_FMT_CODES[fmt_id]
    end

    def resolve_effective_xf(xf_data, cell_style_xfs)
      return nil unless xf_data

      effective = xf_data.dup
      style_xf = nil
      style_xf = cell_style_xfs[effective[:xf_id]] if effective[:xf_id]
      return effective unless style_xf

      %i[num_fmt_id font_id fill_id border_id].each do |k|
        effective[k] = style_xf[k] if (effective[k].nil? || effective[k].zero?) && style_xf.key?(k)
      end
      effective[:alignment] ||= style_xf[:alignment]
      effective[:protection] ||= style_xf[:protection]
      effective
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

    def parse_worksheet_data_validations_options(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = DataValidationsListener.new
      parser.listen(listener)
      parser.parse
      listener.container_options
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

    def parse_worksheet_phonetic_pr(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = PhoneticPrListener.new
      parser.listen(listener)
      parser.parse
      listener.phonetic_pr
    end

    def parse_worksheet_sheet_protection(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = SheetProtectionListener.new
      parser.listen(listener)
      parser.parse
      listener.protection
    end

    def parse_worksheet_protected_ranges(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = ProtectedRangesListener.new
      parser.listen(listener)
      parser.parse
      listener.ranges
    end

    def parse_worksheet_cell_watches(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = CellWatchesListener.new
      parser.listen(listener)
      parser.parse
      listener.watches
    end

    def parse_worksheet_ignored_errors(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = IgnoredErrorsListener.new
      parser.listen(listener)
      parser.parse
      listener.errors
    end

    def parse_worksheet_data_consolidate(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = DataConsolidateListener.new
      parser.listen(listener)
      parser.parse
      listener.result
    end

    def parse_worksheet_scenarios(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = ScenariosListener.new
      parser.listen(listener)
      parser.parse
      listener.result
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
        when "strike"
          @current_font[:strike] = true if @inside_rpr
        when "u"
          if @inside_rpr
            val = attributes["val"]
            @current_font[:underline] = val || true
          end
        when "vertAlign"
          @current_font[:vert_align] = attributes["val"] if @inside_rpr && attributes["val"]
        when "sz"
          @current_font[:sz] = attributes["val"]&.to_f if @inside_rpr
        when "color"
          if @inside_rpr
            if attributes["rgb"]
              @current_font[:color] = attributes["rgb"]
            elsif attributes["theme"]
              @current_font[:theme] = attributes["theme"].to_i
              @current_font[:tint] = attributes["tint"].to_f if attributes["tint"]
            elsif attributes["indexed"]
              @current_font[:indexed] = attributes["indexed"].to_i
            end
          end
        when "rFont"
          @current_font[:name] = attributes["val"] if @inside_rpr
        when "family"
          @current_font[:family] = attributes["val"]&.to_i if @inside_rpr
        when "scheme"
          @current_font[:scheme] = attributes["val"] if @inside_rpr
        when "t"
          @inside_t = @inside_si
          @text_buffer = +"" if @inside_r
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
          @strings << if @has_runs
                        Xlsxrb::RichText.new(runs: @runs)
                      else
                        @text_buffer.dup
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
          @formula_ca = nil
          @formula_aca = nil
          @formula_bx = nil
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
          @formula_ca = true if %w[1 true].include?(attributes["ca"])
          @formula_aca = true if %w[1 true].include?(attributes["aca"])
          @formula_bx = true if %w[1 true].include?(attributes["bx"])
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
        when "strike"
          @is_current_font[:strike] = true if @inside_is_rpr
        when "u"
          if @inside_is_rpr
            val = attributes["val"]
            @is_current_font[:underline] = val || true
          end
        when "vertAlign"
          @is_current_font[:vert_align] = attributes["val"] if @inside_is_rpr && attributes["val"]
        when "sz"
          @is_current_font[:sz] = attributes["val"]&.to_f if @inside_is_rpr
        when "color"
          if @inside_is_rpr
            if attributes["rgb"]
              @is_current_font[:color] = attributes["rgb"]
            elsif attributes["theme"]
              @is_current_font[:theme] = attributes["theme"].to_i
              @is_current_font[:tint] = attributes["tint"].to_f if attributes["tint"]
            elsif attributes["indexed"]
              @is_current_font[:indexed] = attributes["indexed"].to_i
            end
          end
        when "rFont"
          @is_current_font[:name] = attributes["val"] if @inside_is_rpr
        when "family"
          @is_current_font[:family] = attributes["val"]&.to_i if @inside_is_rpr
        when "scheme"
          @is_current_font[:scheme] = attributes["val"] if @inside_is_rpr
        when "t"
          if @inside_is_r
            @is_run_text = +""
            @inside_inline_text = true
          elsif @inside_is || (@current_cell_type == "inlineStr" && !@current_cell_ref.nil?)
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
        s = attributes["s"]
        attrs[:style] = s.to_i if s
        attrs[:thick_top] = true if attributes["thickTop"] == "1"
        attrs[:thick_bot] = true if attributes["thickBot"] == "1"
        attrs[:ph] = true if attributes["ph"] == "1"
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
            shared_index: @formula_si,
            calculate_always: @formula_ca,
            aca: @formula_aca,
            bx: @formula_bx
          )
          return
        end

        case @current_cell_type
        when "inlineStr"
          @cells[@current_cell_ref] = if @is_has_runs
                                        RichText.new(runs: @is_runs.map(&:dup))
                                      else
                                        @inline_text_buffer.dup
                                      end
        when "s"
          index = @value_buffer.to_i
          @cells[@current_cell_ref] = @shared_strings[index] || ""
        when "e"
          code = @value_buffer.dup
          @cells[@current_cell_ref] = if VALID_ERROR_CODES.include?(code)
                                        CellError.new(code:)
                                      else
                                        code
                                      end
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

      attr_reader :sheets, :workbook_properties, :workbook_views, :calc_properties, :defined_names,
                  :workbook_protection, :file_version, :file_sharing, :conformance, :file_recovery_properties

      def initialize
        @sheets = []
        @workbook_properties = {}
        @workbook_views = {}
        @calc_properties = {}
        @file_recovery_properties = {}
        @defined_names = []
        @workbook_protection = nil
        @file_version = {}
        @file_sharing = {}
        @conformance = nil
        @inside_defined_name = false
        @current_dn_attrs = nil
        @dn_text_buffer = +""
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        case name
        when "workbook"
          @conformance = attributes["conformance"] if attributes["conformance"]
        when "sheet"
          @sheets << { name: attributes["name"], rid: attributes["r:id"], state: attributes["state"] }
        when "fileVersion"
          an = attributes["appName"]
          @file_version[:app_name] = an if an
          le = attributes["lastEdited"]
          @file_version[:last_edited] = le if le
          loe = attributes["lowestEdited"]
          @file_version[:lowest_edited] = loe if loe
          rb = attributes["rupBuild"]
          @file_version[:rup_build] = rb if rb
          cn = attributes["codeName"]
          @file_version[:code_name] = cn if cn
        when "fileSharing"
          @file_sharing[:read_only_recommended] = true if %w[1 true].include?(attributes["readOnlyRecommended"])
          un = attributes["userName"]
          @file_sharing[:user_name] = un if un
          an = attributes["algorithmName"]
          @file_sharing[:algorithm_name] = an if an
          hv = attributes["hashValue"]
          @file_sharing[:hash_value] = hv if hv
          sv = attributes["saltValue"]
          @file_sharing[:salt_value] = sv if sv
          sc = attributes["spinCount"]
          @file_sharing[:spin_count] = sc.to_i if sc
        when "workbookPr"
          d1904 = attributes["date1904"]
          @workbook_properties[:date1904] = %w[1 true].include?(d1904) unless d1904.nil?
          dtv = attributes["defaultThemeVersion"]
          @workbook_properties[:default_theme_version] = dtv.to_i if dtv
          cn = attributes["codeName"]
          @workbook_properties[:code_name] = cn if cn
          fp = attributes["filterPrivacy"]
          @workbook_properties[:filter_privacy] = %w[1 true].include?(fp) unless fp.nil?
          acp = attributes["autoCompressPictures"]
          @workbook_properties[:auto_compress_pictures] = %w[1 true].include?(acp) unless acp.nil?
          bf = attributes["backupFile"]
          @workbook_properties[:backup_file] = %w[1 true].include?(bf) unless bf.nil?
          so = attributes["showObjects"]
          @workbook_properties[:show_objects] = so if so
          ul = attributes["updateLinks"]
          @workbook_properties[:update_links] = ul if ul
          rac = attributes["refreshAllConnections"]
          @workbook_properties[:refresh_all_connections] = %w[1 true].include?(rac) unless rac.nil?
          cc = attributes["checkCompatibility"]
          @workbook_properties[:check_compatibility] = %w[1 true].include?(cc) unless cc.nil?
          hpfl = attributes["hidePivotFieldList"]
          @workbook_properties[:hide_pivot_field_list] = %w[1 true].include?(hpfl) unless hpfl.nil?
          sbut = attributes["showBorderUnselectedTables"]
          @workbook_properties[:show_border_unselected_tables] = %w[1 true].include?(sbut) unless sbut.nil?
          ps = attributes["promptedSolutions"]
          @workbook_properties[:prompted_solutions] = %w[1 true].include?(ps) unless ps.nil?
          sia = attributes["showInkAnnotation"]
          @workbook_properties[:show_ink_annotation] = %w[1 true].include?(sia) unless sia.nil?
          selv = attributes["saveExternalLinkValues"]
          @workbook_properties[:save_external_link_values] = %w[1 true].include?(selv) unless selv.nil?
          spcf = attributes["showPivotChartFilter"]
          @workbook_properties[:show_pivot_chart_filter] = %w[1 true].include?(spcf) unless spcf.nil?
          arq = attributes["allowRefreshQuery"]
          @workbook_properties[:allow_refresh_query] = %w[1 true].include?(arq) unless arq.nil?
          pi = attributes["publishItems"]
          @workbook_properties[:publish_items] = %w[1 true].include?(pi) unless pi.nil?
          dcompat = attributes["dateCompatibility"]
          @workbook_properties[:date_compatibility] = %w[1 true].include?(dcompat) unless dcompat.nil?
        when "workbookView"
          at = attributes["activeTab"]
          @workbook_views[:active_tab] = at.to_i if at
          fs = attributes["firstSheet"]
          @workbook_views[:first_sheet] = fs.to_i if fs
          vis = attributes["visibility"]
          @workbook_views[:visibility] = vis if vis
          min = attributes["minimized"]
          @workbook_views[:minimized] = %w[1 true].include?(min) unless min.nil?
          shs = attributes["showHorizontalScroll"]
          @workbook_views[:show_horizontal_scroll] = %w[1 true].include?(shs) unless shs.nil?
          svs = attributes["showVerticalScroll"]
          @workbook_views[:show_vertical_scroll] = %w[1 true].include?(svs) unless svs.nil?
          sst = attributes["showSheetTabs"]
          @workbook_views[:show_sheet_tabs] = %w[1 true].include?(sst) unless sst.nil?
          xw = attributes["xWindow"]
          @workbook_views[:x_window] = xw.to_i if xw
          yw = attributes["yWindow"]
          @workbook_views[:y_window] = yw.to_i if yw
          ww = attributes["windowWidth"]
          @workbook_views[:window_width] = ww.to_i if ww
          wh = attributes["windowHeight"]
          @workbook_views[:window_height] = wh.to_i if wh
          tr = attributes["tabRatio"]
          @workbook_views[:tab_ratio] = tr.to_i if tr
          afdg = attributes["autoFilterDateGrouping"]
          @workbook_views[:auto_filter_date_grouping] = %w[1 true].include?(afdg) unless afdg.nil?
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
          fprec = attributes["fullPrecision"]
          @calc_properties[:full_precision] = %w[1 true].include?(fprec) unless fprec.nil?
          conc = attributes["concurrentCalc"]
          @calc_properties[:concurrent_calc] = %w[1 true].include?(conc) unless conc.nil?
          cmc = attributes["concurrentManualCount"]
          @calc_properties[:concurrent_manual_count] = cmc.to_i if cmc
          ffc = attributes["forceFullCalc"]
          @calc_properties[:force_full_calc] = %w[1 true].include?(ffc) unless ffc.nil?
        when "fileRecoveryPr"
          ar = attributes["autoRecover"]
          @file_recovery_properties[:auto_recover] = %w[1 true].include?(ar) unless ar.nil?
          cs = attributes["crashSave"]
          @file_recovery_properties[:crash_save] = %w[1 true].include?(cs) unless cs.nil?
          del = attributes["dataExtractLoad"]
          @file_recovery_properties[:data_extract_load] = %w[1 true].include?(del) unless del.nil?
          rl = attributes["repairLoad"]
          @file_recovery_properties[:repair_load] = %w[1 true].include?(rl) unless rl.nil?
        when "workbookProtection"
          prot = {}
          ls = attributes["lockStructure"]
          prot[:lock_structure] = %w[1 true].include?(ls) unless ls.nil?
          lw = attributes["lockWindows"]
          prot[:lock_windows] = %w[1 true].include?(lw) unless lw.nil?
          lr = attributes["lockRevision"]
          prot[:lock_revision] = %w[1 true].include?(lr) unless lr.nil?
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
          ran = attributes["revisionsAlgorithmName"]
          if ran
            prot[:revisions_algorithm_name] = ran
            rhv = attributes["revisionsHashValue"]
            prot[:revisions_hash_value] = rhv if rhv
            rsv = attributes["revisionsSaltValue"]
            prot[:revisions_salt_value] = rsv if rsv
            rsc = attributes["revisionsSpinCount"]
            prot[:revisions_spin_count] = rsc.to_i if rsc
          end
          rp = attributes["revisionsPassword"]
          prot[:revisions_password] = rp if rp
          @workbook_protection = prot unless prot.empty?
        when "definedName"
          @inside_defined_name = true
          @current_dn_attrs = {
            name: attributes["name"],
            hidden: %w[1 true].include?(attributes["hidden"])
          }
          lsi = attributes["localSheetId"]
          @current_dn_attrs[:local_sheet_id] = lsi.to_i if lsi
          @current_dn_attrs[:comment] = attributes["comment"] if attributes["comment"]
          @current_dn_attrs[:description] = attributes["description"] if attributes["description"]
          @current_dn_attrs[:function] = true if %w[1 true].include?(attributes["function"])
          @current_dn_attrs[:vb_procedure] = true if %w[1 true].include?(attributes["vbProcedure"])
          @current_dn_attrs[:xlm] = true if %w[1 true].include?(attributes["xlm"])
          @current_dn_attrs[:shortcut_key] = attributes["shortcutKey"] if attributes["shortcutKey"]
          @current_dn_attrs[:publish_to_server] = true if %w[1 true].include?(attributes["publishToServer"])
          @current_dn_attrs[:workbook_parameter] = true if %w[1 true].include?(attributes["workbookParameter"])
          fgi = attributes["functionGroupId"]
          @current_dn_attrs[:function_group_id] = fgi.to_i if fgi
          @current_dn_attrs[:custom_menu] = attributes["customMenu"] if attributes["customMenu"]
          @current_dn_attrs[:help] = attributes["help"] if attributes["help"]
          @current_dn_attrs[:status_bar] = attributes["statusBar"] if attributes["statusBar"]
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
          attrs[:phonetic] = true if attributes["phonetic"] == "1"
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
        return unless ref

        link = { ref: ref }
        link[:rid] = attributes["r:id"] if attributes["r:id"]
        link[:display] = attributes["display"] if attributes["display"]
        link[:tooltip] = attributes["tooltip"] if attributes["tooltip"]
        link[:location] = attributes["location"] if attributes["location"]
        @links << link
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

      attr_reader :num_fmts, :cell_xfs, :cell_style_xfs, :cell_styles, :fonts, :fills, :borders, :dxfs,
                  :indexed_colors, :mru_colors, :table_styles

      def initialize
        @num_fmts = {} # { numFmtId => formatCode }
        @cell_xfs = [] # Array of { num_fmt_id:, font_id:, fill_id:, border_id: }
        @cell_style_xfs = [] # Array of { num_fmt_id:, font_id:, fill_id:, border_id: }
        @cell_styles = [] # Array of { name:, xf_id:, builtin_id: }
        @fonts = []
        @fills = []
        @borders = []
        @dxfs = []
        @indexed_colors = []
        @mru_colors = []
        @table_styles = {}
        @inside_cell_xfs = false
        @inside_cell_style_xfs = false
        @inside_cell_styles = false
        @inside_fonts = false
        @inside_fills = false
        @inside_borders = false
        @inside_dxfs = false
        @inside_indexed_colors = false
        @inside_mru_colors = false
        @inside_table_styles = false
        @current_table_style = nil
        @current_font = nil
        @current_fill = nil
        @current_border = nil
        @current_border_side = nil
        @current_dxf = nil
        @current_xf = nil
        @current_gradient_stop = nil
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)

        case name
        when "numFmt"
          id = attributes["numFmtId"]&.to_i
          code = attributes["formatCode"]
          if @inside_dxfs && @current_dxf
            @current_dxf[:num_fmt] = { num_fmt_id: id, format_code: code } if id && code
          elsif id && code
            @num_fmts[id] = code
          end
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
            il = attributes["iLevel"]
            entry[:i_level] = il.to_i if il
            entry[:hidden] = true if %w[1 true].include?(attributes["hidden"])
            entry[:custom_builtin] = true if %w[1 true].include?(attributes["customBuiltin"])
            @cell_styles << entry
          end
        when "xf"
          xf_entry = {
            num_fmt_id: attributes["numFmtId"]&.to_i,
            font_id: attributes["fontId"]&.to_i,
            fill_id: attributes["fillId"]&.to_i,
            border_id: attributes["borderId"]&.to_i
          }
          %w[applyNumberFormat applyFont applyFill applyBorder applyAlignment applyProtection].each do |attr|
            key = attr.gsub(/[A-Z]/) { |m| "_#{m.downcase}" }.to_sym
            xf_entry[key] = attributes[attr] == "1" unless attributes[attr].nil?
          end
          if @inside_cell_xfs
            xf_entry[:xf_id] = attributes["xfId"]&.to_i
            xf_entry[:quote_prefix] = true if attributes["quotePrefix"] == "1"
            xf_entry[:pivot_button] = true if attributes["pivotButton"] == "1"
            @cell_xfs << xf_entry
            @current_xf = xf_entry
          elsif @inside_cell_style_xfs
            @cell_style_xfs << xf_entry
            @current_xf = xf_entry
          else
            @current_xf = nil
          end
        when "alignment"
          alignment = {}
          alignment[:horizontal] = attributes["horizontal"] if attributes["horizontal"]
          alignment[:vertical] = attributes["vertical"] if attributes["vertical"]
          alignment[:wrap_text] = true if attributes["wrapText"] == "1"
          alignment[:text_rotation] = attributes["textRotation"].to_i if attributes["textRotation"]
          alignment[:indent] = attributes["indent"].to_i if attributes["indent"]
          ri = attributes["relativeIndent"]
          alignment[:relative_indent] = ri.to_i if ri
          alignment[:shrink_to_fit] = true if attributes["shrinkToFit"] == "1"
          ro = attributes["readingOrder"]
          alignment[:reading_order] = ro.to_i if ro
          alignment[:justify_last_line] = true if attributes["justifyLastLine"] == "1"
          unless alignment.empty?
            if @inside_dxfs && @current_dxf
              @current_dxf[:alignment] = alignment
            elsif @current_xf
              @current_xf[:alignment] = alignment
            end
          end
        when "protection"
          protection = {}
          protection[:locked] = attributes["locked"] != "0" if attributes.key?("locked")
          protection[:hidden] = attributes["hidden"] == "1" if attributes.key?("hidden")
          unless protection.empty?
            if @inside_dxfs && @current_dxf
              @current_dxf[:protection] = protection
            elsif @current_xf
              @current_xf[:protection] = protection
            end
          end
        when "fonts"
          @inside_fonts = true
        when "font"
          @current_font = {} if @inside_fonts || @inside_dxfs
        when "b"
          @current_font[:bold] = true if @current_font
        when "i"
          @current_font[:italic] = true if @current_font
        when "strike"
          @current_font[:strike] = true if @current_font
        when "shadow"
          @current_font[:shadow] = true if @current_font
        when "outline"
          @current_font[:outline] = true if @current_font
        when "condense"
          @current_font[:condense] = true if @current_font
        when "extend"
          @current_font[:extend] = true if @current_font
        when "u"
          if @current_font
            val = attributes["val"]
            @current_font[:underline] = val || true
          end
        when "vertAlign"
          @current_font[:vert_align] = attributes["val"] if @current_font && attributes["val"]
        when "scheme"
          @current_font[:scheme] = attributes["val"] if @current_font && attributes["val"]
        when "family"
          @current_font[:family] = attributes["val"].to_i if @current_font && attributes["val"]
        when "charset"
          @current_font[:charset] = attributes["val"].to_i if @current_font && attributes["val"]
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
        when "gradientFill"
          if @current_fill
            gradient = {}
            gradient[:type] = attributes["type"] if attributes["type"]
            gradient[:degree] = attributes["degree"].to_f if attributes["degree"]
            gradient[:left] = attributes["left"].to_f if attributes["left"]
            gradient[:right] = attributes["right"].to_f if attributes["right"]
            gradient[:top] = attributes["top"].to_f if attributes["top"]
            gradient[:bottom] = attributes["bottom"].to_f if attributes["bottom"]
            gradient[:stops] = []
            @current_fill[:gradient] = gradient
          end
        when "stop"
          @current_gradient_stop = { position: attributes["position"].to_f } if @current_fill&.dig(:gradient)
        when "fgColor"
          parse_fill_color(:fg_color, attributes) if @current_fill
        when "bgColor"
          parse_fill_color(:bg_color, attributes) if @current_fill
        when "borders"
          @inside_borders = true
        when "border"
          if @inside_borders || @inside_dxfs
            @current_border = {}
            @current_border[:diagonal_up] = true if attributes["diagonalUp"] == "1"
            @current_border[:diagonal_down] = true if attributes["diagonalDown"] == "1"
            ol = attributes["outline"]
            @current_border[:outline] = %w[1 true].include?(ol) unless ol.nil?
          end
        when "left", "right", "top", "bottom", "diagonal", "vertical", "horizontal", "start", "end"
          if @current_border
            style = attributes["style"]
            @current_border_side = name.to_sym
            @current_border[@current_border_side] = { style: style } if style
          end
        when "dxfs"
          @inside_dxfs = true
        when "dxf"
          @current_dxf = {}
        when "indexedColors"
          @inside_indexed_colors = true
        when "mruColors"
          @inside_mru_colors = true
        when "rgbColor"
          @indexed_colors << attributes["rgb"] if @inside_indexed_colors && attributes["rgb"]
        when "tableStyles"
          @inside_table_styles = true
          @table_styles[:default_table_style] = attributes["defaultTableStyle"] if attributes["defaultTableStyle"]
          @table_styles[:default_pivot_style] = attributes["defaultPivotStyle"] if attributes["defaultPivotStyle"]
          @table_styles[:styles] = []
        when "tableStyle"
          if @inside_table_styles
            ts = { name: attributes["name"], elements: [] }
            ts[:pivot] = %w[1 true].include?(attributes["pivot"]) if attributes.key?("pivot")
            ts[:table] = %w[1 true].include?(attributes["table"]) if attributes.key?("table")
            @current_table_style = ts
          end
        when "tableStyleElement"
          if @current_table_style
            el = { type: attributes["type"] }
            el[:size] = attributes["size"].to_i if attributes["size"]
            el[:dxf_id] = attributes["dxfId"].to_i if attributes["dxfId"]
            @current_table_style[:elements] << el
          end
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
        when "xf"
          @current_xf = nil
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
        when "stop"
          @current_fill[:gradient][:stops] << @current_gradient_stop if @current_gradient_stop && @current_fill&.dig(:gradient)
          @current_gradient_stop = nil
        when "borders"
          @inside_borders = false
        when "border"
          if @inside_dxfs && @current_dxf
            @current_dxf[:border] = @current_border
          elsif @inside_borders
            @borders << @current_border
          end
          @current_border = nil
        when "left", "right", "top", "bottom", "diagonal", "vertical", "horizontal", "start", "end"
          @current_border_side = nil
        when "dxfs"
          @inside_dxfs = false
        when "dxf"
          @dxfs << @current_dxf if @current_dxf
          @current_dxf = nil
        when "indexedColors"
          @inside_indexed_colors = false
        when "mruColors"
          @inside_mru_colors = false
        when "tableStyles"
          @inside_table_styles = false
        when "tableStyle"
          if @inside_table_styles && @current_table_style
            @table_styles[:styles] << @current_table_style
            @current_table_style = nil
          end
        end
      end

      private

      def parse_color(attributes)
        if @inside_mru_colors
          c = {}
          c[:auto] = true if %w[1 true].include?(attributes["auto"])
          c[:rgb] = attributes["rgb"] if attributes["rgb"]
          c[:theme] = attributes["theme"].to_i if attributes["theme"]
          c[:tint] = attributes["tint"].to_f if attributes["tint"]
          c[:indexed] = attributes["indexed"].to_i if attributes["indexed"]
          @mru_colors << c unless c.empty?
        elsif @current_gradient_stop
          @current_gradient_stop[:auto] = true if %w[1 true].include?(attributes["auto"])
          @current_gradient_stop[:color] = attributes["rgb"] if attributes["rgb"]
          @current_gradient_stop[:theme] = attributes["theme"].to_i if attributes["theme"]
          @current_gradient_stop[:tint] = attributes["tint"].to_f if attributes["tint"]
          @current_gradient_stop[:indexed] = attributes["indexed"].to_i if attributes["indexed"]
        elsif @current_border_side && @current_border
          side_data = @current_border[@current_border_side]
          if side_data.is_a?(Hash)
            side_data[:auto] = true if %w[1 true].include?(attributes["auto"])
            side_data[:color] = attributes["rgb"] if attributes["rgb"]
            side_data[:theme] = attributes["theme"].to_i if attributes["theme"]
            side_data[:tint] = attributes["tint"].to_f if attributes["tint"]
            side_data[:indexed] = attributes["indexed"].to_i if attributes["indexed"]
          end
        elsif @current_font
          @current_font[:auto] = true if %w[1 true].include?(attributes["auto"])
          @current_font[:color] = attributes["rgb"] if attributes["rgb"]
          @current_font[:theme] = attributes["theme"].to_i if attributes["theme"]
          @current_font[:tint] = attributes["tint"].to_f if attributes["tint"]
          @current_font[:indexed] = attributes["indexed"].to_i if attributes["indexed"]
        end
      end

      def parse_fill_color(key, attributes)
        if %w[1 true].include?(attributes["auto"])
          @current_fill[:"#{key}_auto"] = true
        elsif attributes["rgb"]
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
          @fc_hidden_button = attributes["hiddenButton"] == "1"
          @fc_show_button = attributes["showButton"] == "0" ? false : nil
        when "filters"
          @filter_blank = attributes["blank"] == "1"
          @filter_calendar_type = attributes["calendarType"]
          @filter_values = []
          @date_group_items = []
        when "filter"
          val = attributes["val"]
          @filter_values << val if val
        when "dateGroupItem"
          dg = { date_time_grouping: attributes["dateTimeGrouping"] }
          dg[:year] = attributes["year"].to_i if attributes["year"]
          dg[:month] = attributes["month"].to_i if attributes["month"]
          dg[:day] = attributes["day"].to_i if attributes["day"]
          dg[:hour] = attributes["hour"].to_i if attributes["hour"]
          dg[:minute] = attributes["minute"].to_i if attributes["minute"]
          dg[:second] = attributes["second"].to_i if attributes["second"]
          @date_group_items << dg
        when "customFilters"
          @inside_custom_filters = true
          @custom_filters_and = attributes["and"] == "1"
          @custom_filters_list = []
        when "customFilter"
          @custom_filters_list << { operator: attributes["operator"], val: attributes["val"] } if @inside_custom_filters
        when "dynamicFilter"
          df = { type: :dynamic, dynamic_type: attributes["type"] }
          df[:val] = attributes["val"].to_f if attributes["val"]
          df[:val_iso] = attributes["valIso"] if attributes["valIso"]
          df[:max_val] = attributes["maxVal"].to_f if attributes["maxVal"]
          df[:max_val_iso] = attributes["maxValIso"] if attributes["maxValIso"]
          @current_filter = df
        when "top10"
          t10 = {
            type: :top10,
            top: attributes["top"] == "1",
            percent: attributes["percent"] == "1",
            val: attributes["val"]&.to_f&.to_i
          }
          t10[:filter_val] = attributes["filterVal"].to_f if attributes["filterVal"]
          @current_filter = t10
        when "colorFilter"
          cf = { type: :color_filter, dxf_id: attributes["dxfId"]&.to_i }
          cf[:cell_color] = false if attributes["cellColor"] == "0"
          @current_filter = cf
        when "iconFilter"
          icf = { type: :icon_filter, icon_set: attributes["iconSet"] }
          icf[:icon_id] = attributes["iconId"].to_i if attributes["iconId"]
          @current_filter = icf
        end
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        case name
        when "filters"
          f = { type: :filters }
          f[:blank] = true if @filter_blank
          f[:calendar_type] = @filter_calendar_type if @filter_calendar_type
          f[:values] = @filter_values unless @filter_values.empty?
          f[:date_group_items] = @date_group_items unless @date_group_items.empty?
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
          if @current_col_id && @current_filter
            @current_filter[:hidden_button] = true if @fc_hidden_button
            @current_filter[:show_button] = false if @fc_show_button == false
            @filter_columns[@current_col_id] = @current_filter
          end
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
          ss = { ref: attributes["ref"], sort_conditions: [] }
          ss[:column_sort] = true if attributes["columnSort"] == "1"
          ss[:case_sensitive] = true if attributes["caseSensitive"] == "1"
          ss[:sort_method] = attributes["sortMethod"] if attributes["sortMethod"]
          @sort_state = ss
        when "sortCondition"
          return unless @inside_sort_state

          sc = { ref: attributes["ref"] }
          sc[:descending] = true if attributes["descending"] == "1"
          sc[:sort_by] = attributes["sortBy"] if attributes["sortBy"]
          sc[:custom_list] = attributes["customList"] if attributes["customList"]
          dxf = attributes["dxfId"]
          sc[:dxf_id] = dxf.to_i if dxf
          sc[:icon_set] = attributes["iconSet"] if attributes["iconSet"]
          iid = attributes["iconId"]
          sc[:icon_id] = iid.to_i if iid
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
        "subject" => :subject,
        "creator" => :creator,
        "keywords" => :keywords,
        "description" => :description,
        "lastModifiedBy" => :last_modified_by,
        "revision" => :revision,
        "created" => :created,
        "modified" => :modified,
        "category" => :category,
        "contentStatus" => :content_status,
        "language" => :language
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

    # SAX2 listener for parsing custom properties (docProps/custom.xml).
    class CustomPropertiesListener
      include REXML::SAX2Listener

      attr_reader :properties

      def initialize
        @properties = []
        @current_name = nil
        @current_type = nil
        @text_buffer = +""
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        case name
        when "property"
          @current_name = attributes["name"]
        when "lpwstr"
          @current_type = :string
          @text_buffer = +""
        when "i4"
          @current_type = :number
          @text_buffer = +""
        when "r8"
          @current_type = :float
          @text_buffer = +""
        when "bool"
          @current_type = :bool
          @text_buffer = +""
        when "filetime"
          @current_type = :date
          @text_buffer = +""
        end
      end

      def characters(text)
        @text_buffer << text if @current_type
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        case name
        when "property"
          @current_name = nil
        when "lpwstr", "i4", "r8", "bool", "filetime"
          if @current_name
            value = case @current_type
                    when :number then @text_buffer.to_i
                    when :float then @text_buffer.to_f
                    when :bool then @text_buffer == "true"
                    else @text_buffer.dup
                    end
            @properties << { name: @current_name, value: value, type: @current_type }
          end
          @current_type = nil
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
          sh = attributes["syncHorizontal"]
          @properties[:sync_horizontal] = %w[1 true].include?(sh) unless sh.nil?
          sv = attributes["syncVertical"]
          @properties[:sync_vertical] = %w[1 true].include?(sv) unless sv.nil?
          @properties[:sync_ref] = attributes["syncRef"] if attributes["syncRef"]
          te = attributes["transitionEvaluation"]
          @properties[:transition_evaluation] = %w[1 true].include?(te) unless te.nil?
          tent = attributes["transitionEntry"]
          @properties[:transition_entry] = %w[1 true].include?(tent) unless tent.nil?
          @properties[:code_name] = attributes["codeName"] if attributes["codeName"]
          fm = attributes["filterMode"]
          @properties[:filter_mode] = %w[1 true].include?(fm) unless fm.nil?
          pub = attributes["published"]
          @properties[:published] = %w[1 true].include?(pub) unless pub.nil?
          efcc = attributes["enableFormatConditionsCalculation"]
          @properties[:enable_format_conditions_calculation] = %w[1 true].include?(efcc) unless efcc.nil?
        when "tabColor"
          if @inside_sheet_pr
            @properties[:tab_color] = attributes["rgb"] if attributes["rgb"]
            @properties[:tab_color_theme] = attributes["theme"].to_i if attributes["theme"]
            @properties[:tab_color_tint] = attributes["tint"].to_f if attributes["tint"]
            @properties[:tab_color_indexed] = attributes["indexed"].to_i if attributes["indexed"]
            @properties[:tab_color_auto] = %w[1 true].include?(attributes["auto"]) if attributes["auto"]
          end
        when "outlinePr"
          if @inside_sheet_pr
            apply_s = attributes["applyStyles"]
            @properties[:apply_styles] = %w[1 true].include?(apply_s) unless apply_s.nil?
            sb = attributes["summaryBelow"]
            @properties[:summary_below] = %w[1 true].include?(sb) unless sb.nil?
            sr = attributes["summaryRight"]
            @properties[:summary_right] = %w[1 true].include?(sr) unless sr.nil?
            sos = attributes["showOutlineSymbols"]
            @properties[:show_outline_symbols] = %w[1 true].include?(sos) unless sos.nil?
          end
        when "pageSetUpPr"
          if @inside_sheet_pr
            ftp = attributes["fitToPage"]
            @properties[:fit_to_page] = %w[1 true].include?(ftp) unless ftp.nil?
            apb = attributes["autoPageBreaks"]
            @properties[:auto_page_breaks] = %w[1 true].include?(apb) unless apb.nil?
          end
        when "sheetCalcPr"
          fcol = attributes["fullCalcOnLoad"]
          @properties[:full_calc_on_load] = %w[1 true].include?(fcol) unless fcol.nil?
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

    # SAX2 listener for parsing <phoneticPr> element.
    class PhoneticPrListener
      include REXML::SAX2Listener

      attr_reader :phonetic_pr

      def initialize
        @phonetic_pr = nil
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        return unless name == "phoneticPr"

        pr = {}
        pr[:font_id] = attributes["fontId"].to_i if attributes["fontId"]
        pr[:type] = attributes["type"] if attributes["type"]
        pr[:alignment] = attributes["alignment"] if attributes["alignment"]
        @phonetic_pr = pr
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

    # SAX2 listener for parsing <protectedRanges> element.
    class ProtectedRangesListener
      include REXML::SAX2Listener

      attr_reader :ranges

      def initialize
        @ranges = []
        @current = nil
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        case name
        when "protectedRange"
          pr = {}
          pr[:sqref] = attributes["sqref"] if attributes["sqref"]
          pr[:name] = attributes["name"] if attributes["name"]
          pr[:algorithm_name] = attributes["algorithmName"] if attributes["algorithmName"]
          pr[:hash_value] = attributes["hashValue"] if attributes["hashValue"]
          pr[:salt_value] = attributes["saltValue"] if attributes["saltValue"]
          pr[:spin_count] = attributes["spinCount"].to_i if attributes["spinCount"]
          @current = pr
        when "securityDescriptor"
          @in_sd = true
          @sd_text = +""
        end
      end

      def characters(text)
        @sd_text << text if @in_sd
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        case name
        when "securityDescriptor"
          (@current[:security_descriptors] ||= []) << @sd_text if @current
          @in_sd = false
        when "protectedRange"
          @ranges << @current if @current
          @current = nil
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

    # SAX2 listener for parsing <cellWatches> element.
    class CellWatchesListener
      include REXML::SAX2Listener

      attr_reader :watches

      def initialize
        @watches = []
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        @watches << attributes["r"] if name == "cellWatch" && attributes["r"]
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

    # SAX2 listener for parsing <ignoredErrors> element.
    class IgnoredErrorsListener
      include REXML::SAX2Listener

      attr_reader :errors

      def initialize
        @errors = []
      end

      IGNORED_ERROR_BOOL_ATTRS = {
        "evalError" => :eval_error, "twoDigitTextYear" => :two_digit_text_year,
        "numberStoredAsText" => :number_stored_as_text, "formula" => :formula,
        "formulaRange" => :formula_range, "unlockedFormula" => :unlocked_formula,
        "emptyCellReference" => :empty_cell_reference, "listDataValidation" => :list_data_validation,
        "calculatedColumn" => :calculated_column
      }.freeze

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        return unless name == "ignoredError" && attributes["sqref"]

        ie = { sqref: attributes["sqref"] }
        IGNORED_ERROR_BOOL_ATTRS.each do |xml_attr, sym|
          ie[sym] = true if attributes[xml_attr] == "1"
        end
        @errors << ie
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

    # SAX2 listener for parsing <dataConsolidate> element.
    class DataConsolidateListener
      include REXML::SAX2Listener

      attr_reader :result

      def initialize
        @result = nil
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        case name
        when "dataConsolidate"
          @result = {}
          @result[:function] = attributes["function"] if attributes["function"]
          @result[:start_labels] = true if %w[1 true].include?(attributes["startLabels"])
          @result[:left_labels] = true if %w[1 true].include?(attributes["leftLabels"])
          @result[:top_labels] = true if %w[1 true].include?(attributes["topLabels"])
          @result[:link] = true if %w[1 true].include?(attributes["link"])
          @result[:data_refs] = []
        when "dataRef"
          if @result
            ref = {}
            ref[:ref] = attributes["ref"] if attributes["ref"]
            ref[:name] = attributes["name"] if attributes["name"]
            ref[:sheet] = attributes["sheet"] if attributes["sheet"]
            @result[:data_refs] << ref
          end
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

    # SAX2 listener for parsing <scenarios> element.
    class ScenariosListener
      include REXML::SAX2Listener

      attr_reader :result

      def initialize
        @result = nil
        @current_scenario = nil
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        case name
        when "scenarios"
          @result = { scenarios: [] }
          @result[:current] = attributes["current"].to_i if attributes["current"]
          @result[:show] = attributes["show"].to_i if attributes["show"]
          @result[:sqref] = attributes["sqref"] if attributes["sqref"]
        when "scenario"
          if @result
            sc = { name: attributes["name"], input_cells: [] }
            sc[:locked] = true if %w[1 true].include?(attributes["locked"])
            sc[:hidden] = true if %w[1 true].include?(attributes["hidden"])
            sc[:user] = attributes["user"] if attributes["user"]
            sc[:comment] = attributes["comment"] if attributes["comment"]
            @current_scenario = sc
          end
        when "inputCells"
          if @current_scenario
            ic = { r: attributes["r"], val: attributes["val"] }
            ic[:deleted] = true if %w[1 true].include?(attributes["deleted"])
            ic[:undone] = true if %w[1 true].include?(attributes["undone"])
            ic[:num_fmt_id] = attributes["numFmtId"].to_i if attributes["numFmtId"]
            @current_scenario[:input_cells] << ic
          end
        end
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        return unless name == "scenario" && @current_scenario && @result

        @result[:scenarios] << @current_scenario
        @current_scenario = nil
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
        olr = attributes["outlineLevelRow"]
        @properties[:outline_level_row] = olr.to_i if olr
        olc = attributes["outlineLevelCol"]
        @properties[:outline_level_col] = olc.to_i if olc
        ch = attributes["customHeight"]
        @properties[:custom_height] = true if %w[1 true].include?(ch)
        zh = attributes["zeroHeight"]
        @properties[:zero_height] = true if %w[1 true].include?(zh)
        tt = attributes["thickTop"]
        @properties[:thick_top] = true if %w[1 true].include?(tt)
        tb = attributes["thickBottom"]
        @properties[:thick_bottom] = true if %w[1 true].include?(tb)
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

          wp = attributes["windowProtection"]
          @view[:window_protection] = %w[1 true].include?(wp) unless wp.nil?
          sf = attributes["showFormulas"]
          @view[:show_formulas] = %w[1 true].include?(sf) unless sf.nil?
          sgl = attributes["showGridLines"]
          @view[:show_grid_lines] = %w[1 true].include?(sgl) unless sgl.nil?
          srch = attributes["showRowColHeaders"]
          @view[:show_row_col_headers] = %w[1 true].include?(srch) unless srch.nil?
          szv = attributes["showZeros"]
          @view[:show_zeros] = %w[1 true].include?(szv) unless szv.nil?
          rtl = attributes["rightToLeft"]
          @view[:right_to_left] = %w[1 true].include?(rtl) unless rtl.nil?
          ts = attributes["tabSelected"]
          @view[:tab_selected] = true if ts == "1"
          srr = attributes["showRuler"]
          @view[:show_ruler] = %w[1 true].include?(srr) unless srr.nil?
          soss = attributes["showOutlineSymbols"]
          @view[:show_outline_symbols] = %w[1 true].include?(soss) unless soss.nil?
          dgc = attributes["defaultGridColor"]
          @view[:default_grid_color] = %w[1 true].include?(dgc) unless dgc.nil?
          sws = attributes["showWhiteSpace"]
          @view[:show_white_space] = %w[1 true].include?(sws) unless sws.nil?
          vm = attributes["view"]
          @view[:view] = vm if vm
          tlc = attributes["topLeftCell"]
          @view[:top_left_cell] = tlc if tlc
          cid = attributes["colorId"]
          @view[:color_id] = cid.to_i if cid
          zs = attributes["zoomScale"]
          @view[:zoom_scale] = zs.to_i if zs
          zsn = attributes["zoomScaleNormal"]
          @view[:zoom_scale_normal] = zsn.to_i if zsn
          zssl = attributes["zoomScaleSheetLayoutView"]
          @view[:zoom_scale_sheet_layout_view] = zssl.to_i if zssl
          zspl = attributes["zoomScalePageLayoutView"]
          @view[:zoom_scale_page_layout_view] = zspl.to_i if zspl
        when "pane"
          return unless @inside_sheet_views

          ys = attributes["ySplit"]
          xs = attributes["xSplit"]
          frozen = attributes["state"] == "frozen"
          tlc = attributes["topLeftCell"]
          ap = attributes["activePane"]
          p = if frozen
                {
                  row: ys ? ys.to_i : 0,
                  col: xs ? xs.to_i : 0,
                  state: :frozen
                }
              else
                {
                  row: ys ? ys.to_i : 0,
                  col: xs ? xs.to_i : 0,
                  x_split: xs ? xs.to_i : 0,
                  y_split: ys ? ys.to_i : 0,
                  top_left_cell: tlc,
                  state: :split
                }
              end
          p[:active_pane] = ap if ap
          @pane = p
        when "selection"
          return unless @inside_sheet_views

          ac = attributes["activeCell"]
          sq = attributes["sqref"]
          sel = { active_cell: ac, sqref: sq }
          pn = attributes["pane"]
          sel[:pane] = pn if pn
          acid = attributes["activeCellId"]
          sel[:active_cell_id] = acid.to_i if acid
          @selection = sel if ac || sq || pn || acid
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

      attr_reader :validations, :container_options

      def initialize
        @validations = []
        @container_options = {}
        @current_dv = nil
        @inside_formula1 = false
        @inside_formula2 = false
        @text_buffer = +""
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        case name
        when "dataValidations"
          dp = attributes["disablePrompts"]
          @container_options[:disable_prompts] = true if %w[1 true].include?(dp)
          xw = attributes["xWindow"]
          @container_options[:x_window] = xw.to_i if xw
          yw = attributes["yWindow"]
          @container_options[:y_window] = yw.to_i if yw
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
          @current_dv[:show_drop_down] = true if attributes["showDropDown"] == "1"
          @current_dv[:ime_mode] = attributes["imeMode"] if attributes["imeMode"]
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
        @current_pivot = false
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
          @current_pivot = attributes["pivot"] == "1"
        when "cfRule"
          @current_rule = { sqref: @current_sqref, type: attributes["type"] }
          @current_rule[:pivot] = true if @current_pivot
          @current_rule[:priority] = attributes["priority"].to_i if attributes["priority"]
          @current_rule[:operator] = attributes["operator"] if attributes["operator"]
          @current_rule[:format_id] = attributes["dxfId"].to_i if attributes["dxfId"]
          @current_rule[:stop_if_true] = true if attributes["stopIfTrue"] == "1"
          @current_rule[:above_average] = false if attributes["aboveAverage"] == "0"
          @current_rule[:equal_average] = true if attributes["equalAverage"] == "1"
          @current_rule[:rank] = attributes["rank"].to_i if attributes["rank"]
          @current_rule[:percent] = true if attributes["percent"] == "1"
          @current_rule[:bottom] = true if attributes["bottom"] == "1"
          @current_rule[:text] = attributes["text"] if attributes["text"]
          @current_rule[:time_period] = attributes["timePeriod"] if attributes["timePeriod"]
          sd = attributes["stdDev"]
          @current_rule[:std_dev] = sd.to_i if sd
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
            pct = attributes["percent"]
            is[:percent] = %w[1 true].include?(pct) unless pct.nil?
            sv = attributes["showValue"]
            is[:show_value] = %w[1 true].include?(sv) unless sv.nil?
            @current_rule[:icon_set] = is
          end
          @cfvo_target = :icon_set
        when "cfvo"
          cfvo = { type: attributes["type"] }
          cfvo[:val] = attributes["val"] if attributes["val"]
          gte = attributes["gte"]
          cfvo[:gte] = %w[1 true].include?(gte) unless gte.nil?
          append_cfvo(cfvo)
        when "color"
          if attributes["rgb"]
            append_cf_color({ rgb: attributes["rgb"] })
          elsif attributes["theme"]
            c = { theme: attributes["theme"].to_i }
            c[:tint] = attributes["tint"].to_f if attributes["tint"]
            append_cf_color(c)
          elsif attributes["indexed"]
            append_cf_color({ indexed: attributes["indexed"].to_i })
          end
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

      def append_cf_color(color)
        return unless @current_rule && @color_target

        container = @current_rule[@color_target]
        if container.is_a?(Hash) && container.key?(:colors)
          container[:colors] << color
        elsif container.is_a?(Hash)
          container[:color] = color
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
          gls = attributes["gridLinesSet"]
          @print_options[:grid_lines_set] = gls != "0" if gls
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
          po = attributes["pageOrder"]
          @page_setup[:page_order] = po if po
          baw = attributes["blackAndWhite"]
          @page_setup[:black_and_white] = true if %w[1 true].include?(baw)
          dr = attributes["draft"]
          @page_setup[:draft] = true if %w[1 true].include?(dr)
          cc = attributes["cellComments"]
          @page_setup[:cell_comments] = cc if cc
          fpn = attributes["firstPageNumber"]
          @page_setup[:first_page_number] = fpn.to_i if fpn
          ufpn = attributes["useFirstPageNumber"]
          @page_setup[:use_first_page_number] = true if %w[1 true].include?(ufpn)
          hdpi = attributes["horizontalDpi"]
          @page_setup[:horizontal_dpi] = hdpi.to_i if hdpi
          vdpi = attributes["verticalDpi"]
          @page_setup[:vertical_dpi] = vdpi.to_i if vdpi
          cp = attributes["copies"]
          @page_setup[:copies] = cp.to_i if cp
          ph = attributes["paperHeight"]
          @page_setup[:paper_height] = ph if ph
          pw = attributes["paperWidth"]
          @page_setup[:paper_width] = pw if pw
          err = attributes["errors"]
          @page_setup[:errors] = err if err
          upd = attributes["usePrinterDefaults"]
          @page_setup[:use_printer_defaults] = %w[1 true].include?(upd) unless upd.nil?
        when "headerFooter"
          @inside_header_footer = true
          df = attributes["differentFirst"]
          @header_footer[:different_first] = true if %w[1 true].include?(df)
          doe = attributes["differentOddEven"]
          @header_footer[:different_odd_even] = true if %w[1 true].include?(doe)
          swd = attributes["scaleWithDoc"]
          @header_footer[:scale_with_doc] = swd != "0" if swd
          awm = attributes["alignWithMargins"]
          @header_footer[:align_with_margins] = awm != "0" if awm
        when "oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter"
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
            brk = { id: id }
            mn = attributes["min"]
            brk[:min] = mn.to_i if mn
            mx = attributes["max"]
            brk[:max] = mx.to_i if mx
            brk[:man] = true if %w[1 true].include?(attributes["man"])
            brk[:pt] = true if %w[1 true].include?(attributes["pt"])
            @row_breaks << brk if @inside_row_breaks
            @col_breaks << brk if @inside_col_breaks
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
        when "firstHeader"
          @header_footer[:first_header] = @text_buffer.dup if @current_hf_field == "firstHeader"
          @current_hf_field = nil
        when "firstFooter"
          @header_footer[:first_footer] = @text_buffer.dup if @current_hf_field == "firstFooter"
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
        @inside_totals_formula = false
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
          hrc = attributes["headerRowCount"]
          @table[:header_row_count] = hrc.to_i if hrc
          @table[:published] = true if attributes["published"] == "1"
          @table[:comment] = attributes["comment"] if attributes["comment"]
          @table[:insert_row] = true if attributes["insertRow"] == "1"
          @table[:insert_row_shift] = true if attributes["insertRowShift"] == "1"
          hrd = attributes["headerRowDxfId"]
          @table[:header_row_dxf_id] = hrd.to_i if hrd
          dd = attributes["dataDxfId"]
          @table[:data_dxf_id] = dd.to_i if dd
          trd = attributes["totalsRowDxfId"]
          @table[:totals_row_dxf_id] = trd.to_i if trd
          hrbd = attributes["headerRowBorderDxfId"]
          @table[:header_row_border_dxf_id] = hrbd.to_i if hrbd
          tbd = attributes["tableBorderDxfId"]
          @table[:table_border_dxf_id] = tbd.to_i if tbd
          trbd = attributes["totalsRowBorderDxfId"]
          @table[:totals_row_border_dxf_id] = trbd.to_i if trbd
          @table[:header_row_cell_style] = attributes["headerRowCellStyle"] if attributes["headerRowCellStyle"]
          @table[:totals_row_cell_style] = attributes["totalsRowCellStyle"] if attributes["totalsRowCellStyle"]
          cid = attributes["connectionId"]
          @table[:connection_id] = cid.to_i if cid
          @table[:table_type] = attributes["tableType"] if attributes["tableType"]
        when "tableColumn"
          col = { name: attributes["name"] }
          trf = attributes["totalsRowFunction"]
          col[:totals_row_function] = trf if trf
          trl = attributes["totalsRowLabel"]
          col[:totals_row_label] = trl if trl
          dd = attributes["dataDxfId"]
          col[:data_dxf_id] = dd.to_i if dd
          td = attributes["totalsRowDxfId"]
          col[:totals_row_dxf_id] = td.to_i if td
          hd = attributes["headerRowDxfId"]
          col[:header_row_dxf_id] = hd.to_i if hd
          dcs = attributes["dataCellStyle"]
          col[:data_cell_style] = dcs if dcs
          @current_column = col
        when "calculatedColumnFormula"
          @inside_calc_formula = true
          @text_buffer = +""
        when "totalsRowFormula"
          @inside_totals_formula = true
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
        @text_buffer << text if @inside_calc_formula || @inside_totals_formula
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        case name
        when "calculatedColumnFormula"
          @current_column[:calculated_column_formula] = @text_buffer.dup if @current_column
          @inside_calc_formula = false
        when "totalsRowFormula"
          @current_column[:totals_row_formula] = @text_buffer.dup if @current_column
          @inside_totals_formula = false
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
          @anchor_edit_as = attributes["editAs"]
        when "pic"
          @inside_pic = true
          @current_image = {}
        when "cNvPr"
          if @inside_pic && @current_image
            @current_image[:name] = attributes["name"] if attributes["name"]
            @current_image[:id] = attributes["id"]&.to_i
            @current_image[:description] = attributes["descr"] if attributes["descr"]
            @current_image[:title] = attributes["title"] if attributes["title"]
            @current_image[:hidden] = %w[1 true].include?(attributes["hidden"]) if attributes["hidden"]
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
        when "clientData"
          if @inside_anchor
            @anchor_locks_with_sheet = attributes["fLocksWithSheet"]
            @anchor_prints_with_sheet = attributes["fPrintsWithSheet"]
          end
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
            @current_image[:edit_as] = @anchor_edit_as if @anchor_edit_as
            @images << @current_image
          end
          @current_image = nil
          @inside_pic = false
        when "twoCellAnchor", "oneCellAnchor"
          @images.last[:locks_with_sheet] = @anchor_locks_with_sheet == "1" if @anchor_locks_with_sheet && !@images.empty?
          @images.last[:prints_with_sheet] = @anchor_prints_with_sheet == "1" if @anchor_prints_with_sheet && !@images.empty?
          @inside_anchor = false
          @anchor_from = {}
          @anchor_to = {}
          @anchor_locks_with_sheet = nil
          @anchor_prints_with_sheet = nil
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
        @anchor_edit_as = nil
        @inside_anchor = false
        @inside_from = false
        @inside_to = false
        @anchor_from = {}
        @anchor_to = {}
        @current_field = nil
        @text_buffer = +""
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        case name
        when "twoCellAnchor", "oneCellAnchor"
          @anchor_edit_as = attributes["editAs"]
          @inside_anchor = true
          @anchor_from = {}
          @anchor_to = {}
        when "graphicFrame"
          @inside_graphic_frame = true
          @current_chart = {}
        when "cNvPr"
          if @inside_graphic_frame && @current_chart
            @current_chart[:name] = attributes["name"] if attributes["name"]
            @current_chart[:description] = attributes["descr"] if attributes["descr"]
            @current_chart[:frame_title] = attributes["title"] if attributes["title"]
            @current_chart[:frame_hidden] = %w[1 true].include?(attributes["hidden"]) if attributes["hidden"]
          end
        when "chart"
          rid = attributes["r:id"] || attributes["id"]
          @current_chart[:rid] = rid if @inside_graphic_frame && @current_chart && rid
        when "from"
          @inside_from = true if @inside_anchor
        when "to"
          @inside_to = true if @inside_anchor
        when "col", "colOff", "row", "rowOff"
          @current_field = name
          @text_buffer = +""
        when "clientData"
          @anchor_locks_with_sheet = attributes["fLocksWithSheet"]
          @anchor_prints_with_sheet = attributes["fPrintsWithSheet"]
        end
      end

      def characters(text)
        @text_buffer << text if @current_field
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        case name
        when "col", "colOff", "row", "rowOff"
          val = @text_buffer.strip.to_i
          @anchor_from[name] = val if @inside_from
          @anchor_to[name] = val if @inside_to
          @current_field = nil
        when "from"
          @inside_from = false
        when "to"
          @inside_to = false
        when "graphicFrame"
          if @current_chart && @current_chart[:rid]
            @current_chart[:edit_as] = @anchor_edit_as if @anchor_edit_as
            @current_chart[:from_col] = @anchor_from["col"] if @anchor_from["col"]
            @current_chart[:from_row] = @anchor_from["row"] if @anchor_from["row"]
            @current_chart[:from_col_off] = @anchor_from["colOff"] if @anchor_from["colOff"]
            @current_chart[:from_row_off] = @anchor_from["rowOff"] if @anchor_from["rowOff"]
            @current_chart[:to_col] = @anchor_to["col"] if @anchor_to["col"]
            @current_chart[:to_row] = @anchor_to["row"] if @anchor_to["row"]
            @current_chart[:to_col_off] = @anchor_to["colOff"] if @anchor_to["colOff"]
            @current_chart[:to_row_off] = @anchor_to["rowOff"] if @anchor_to["rowOff"]
            @charts << @current_chart
          end
          @current_chart = nil
          @inside_graphic_frame = false
        when "twoCellAnchor", "oneCellAnchor"
          @charts.last[:locks_with_sheet] = @anchor_locks_with_sheet == "1" if @anchor_locks_with_sheet && !@charts.empty?
          @charts.last[:prints_with_sheet] = @anchor_prints_with_sheet == "1" if @anchor_prints_with_sheet && !@charts.empty?
          @inside_anchor = false
          @anchor_edit_as = nil
          @anchor_locks_with_sheet = nil
          @anchor_prints_with_sheet = nil
          @anchor_from = {}
          @anchor_to = {}
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
          @anchor_edit_as = attributes["editAs"]
        when "sp"
          @inside_sp = true
          @current_shape = {}
        when "cNvPr"
          if @inside_sp && @current_shape
            @current_shape[:name] = attributes["name"] if attributes["name"]
            @current_shape[:id] = attributes["id"]&.to_i
            @current_shape[:description] = attributes["descr"] if attributes["descr"]
            @current_shape[:title] = attributes["title"] if attributes["title"]
            @current_shape[:hidden] = %w[1 true].include?(attributes["hidden"]) if attributes["hidden"]
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
        when "clientData"
          if @inside_anchor
            @anchor_locks_with_sheet = attributes["fLocksWithSheet"]
            @anchor_prints_with_sheet = attributes["fPrintsWithSheet"]
          end
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
            @current_shape[:edit_as] = @anchor_edit_as if @anchor_edit_as
            @shapes << @current_shape
          end
          @current_shape = nil
          @inside_sp = false
          @inside_tx_body = false
        when "twoCellAnchor", "oneCellAnchor"
          @shapes.last[:locks_with_sheet] = @anchor_locks_with_sheet == "1" if @anchor_locks_with_sheet && !@shapes.empty?
          @shapes.last[:prints_with_sheet] = @anchor_prints_with_sheet == "1" if @anchor_prints_with_sheet && !@shapes.empty?
          @inside_anchor = false
          @anchor_from = {}
          @anchor_to = {}
          @anchor_locks_with_sheet = nil
          @anchor_prints_with_sheet = nil
        when "from"
          @inside_from = false
        when "to"
          @inside_to = false
        when "txBody"
          @inside_tx_body = false
        when "t"
          @current_shape[:text] = (@current_shape[:text] || +"") << @text_buffer if @inside_t && @inside_tx_body && @current_shape
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

      attr_reader :chart_type, :title, :series, :legend, :data_labels, :cat_axis_title, :val_axis_title,
                  :grouping, :bar_dir, :vary_colors, :plot_vis_only, :disp_blanks_as, :style, :auto_title_deleted,
                  :rounded_corners, :cat_axis_tick_lbl_pos, :val_axis_tick_lbl_pos,
                  :cat_axis_major_gridlines, :val_axis_major_gridlines,
                  :cat_axis_minor_gridlines, :val_axis_minor_gridlines,
                  :show_d_lbls_over_max, :cat_axis_delete, :val_axis_delete,
                  :cat_axis_orientation, :val_axis_orientation,
                  :gap_width, :overlap, :view_3d,
                  :gap_depth, :bar_shape,
                  :bubble_3d, :bubble_scale, :show_neg_bubbles, :size_represents,
                  :cat_axis_num_fmt, :val_axis_num_fmt,
                  :cat_axis_major_tick_mark, :cat_axis_minor_tick_mark,
                  :val_axis_major_tick_mark, :val_axis_minor_tick_mark,
                  :cat_axis_crosses, :val_axis_crosses,
                  :cat_axis_crosses_at, :val_axis_crosses_at,
                  :cat_axis_tick_lbl_skip, :cat_axis_tick_mark_skip,
                  :cat_axis_lbl_offset, :cat_axis_no_multi_lvl_lbl,
                  :val_axis_cross_between, :val_axis_major_unit, :val_axis_minor_unit,
                  :val_axis_disp_units,
                  :cat_axis_scaling_max, :cat_axis_scaling_min,
                  :val_axis_scaling_max, :val_axis_scaling_min,
                  :cat_axis_log_base, :val_axis_log_base,
                  :first_slice_ang, :hole_size,
                  :smooth, :marker,
                  :scatter_style, :radar_style,
                  :cat_axis_pos, :val_axis_pos,
                  :wireframe

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
        @grouping = nil
        @bar_dir = nil
        @vary_colors = nil
        @plot_vis_only = nil
        @disp_blanks_as = nil
        @style = nil
        @auto_title_deleted = nil
        @rounded_corners = nil
        @cat_axis_tick_lbl_pos = nil
        @val_axis_tick_lbl_pos = nil
        @cat_axis_major_gridlines = false
        @val_axis_major_gridlines = false
        @cat_axis_minor_gridlines = false
        @val_axis_minor_gridlines = false
        @show_d_lbls_over_max = nil
        @cat_axis_delete = nil
        @val_axis_delete = nil
        @cat_axis_orientation = nil
        @val_axis_orientation = nil
        @gap_width = nil
        @overlap = nil
        @gap_depth = nil
        @bar_shape = nil
        @bubble_3d = nil
        @bubble_scale = nil
        @show_neg_bubbles = nil
        @size_represents = nil
        @view_3d = nil
        @cat_axis_num_fmt = nil
        @val_axis_num_fmt = nil
        @cat_axis_major_tick_mark = nil
        @cat_axis_minor_tick_mark = nil
        @val_axis_major_tick_mark = nil
        @val_axis_minor_tick_mark = nil
        @cat_axis_crosses = nil
        @val_axis_crosses = nil
        @cat_axis_crosses_at = nil
        @val_axis_crosses_at = nil
        @cat_axis_tick_lbl_skip = nil
        @cat_axis_tick_mark_skip = nil
        @cat_axis_lbl_offset = nil
        @cat_axis_no_multi_lvl_lbl = nil
        @val_axis_cross_between = nil
        @val_axis_major_unit = nil
        @val_axis_minor_unit = nil
        @val_axis_disp_units = nil
        @cat_axis_scaling_max = nil
        @cat_axis_scaling_min = nil
        @val_axis_scaling_max = nil
        @val_axis_scaling_min = nil
        @cat_axis_log_base = nil
        @val_axis_log_base = nil
        @first_slice_ang = nil
        @hole_size = nil
        @smooth = nil
        @marker = nil
        @scatter_style = nil
        @radar_style = nil
        @cat_axis_pos = nil
        @val_axis_pos = nil
        @wireframe = nil
        @inside_view_3d = false
        @inside_scaling = false
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
        @inside_separator = false
        @inside_cat_ax = false
        @inside_val_ax = false
        @inside_ax_title = false
        @title_depth = 0
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        @chart_type = name if CHART_TYPES.include?(name)

        case name
        when "grouping"
          @grouping = attributes["val"] if attributes["val"]
        when "barDir"
          @bar_dir = attributes["val"] if attributes["val"]
        when "varyColors"
          @vary_colors = attributes["val"] == "1" if attributes["val"]
        when "autoTitleDeleted"
          @auto_title_deleted = attributes["val"] == "1" if attributes["val"]
        when "view3D"
          @inside_view_3d = true
          @view_3d = {}
        when "rotX"
          @view_3d[:rot_x] = attributes["val"].to_i if @inside_view_3d && attributes["val"]
        when "hPercent"
          @view_3d[:h_percent] = attributes["val"].to_i if @inside_view_3d && attributes["val"]
        when "rotY"
          @view_3d[:rot_y] = attributes["val"].to_i if @inside_view_3d && attributes["val"]
        when "depthPercent"
          @view_3d[:depth_percent] = attributes["val"].to_i if @inside_view_3d && attributes["val"]
        when "rAngAx"
          @view_3d[:r_ang_ax] = attributes["val"] == "1" if @inside_view_3d && attributes["val"]
        when "perspective"
          @view_3d[:perspective] = attributes["val"].to_i if @inside_view_3d && attributes["val"]
        when "gapWidth"
          @gap_width = attributes["val"]&.to_i if attributes["val"]
        when "overlap"
          @overlap = attributes["val"]&.to_i if attributes["val"]
        when "gapDepth"
          @gap_depth = attributes["val"]&.to_i if attributes["val"]
        when "shape"
          @bar_shape = attributes["val"] if attributes["val"]
        when "bubble3D"
          @bubble_3d = attributes["val"] == "1" if attributes["val"] && !@inside_ser
        when "bubbleScale"
          @bubble_scale = attributes["val"]&.to_i if attributes["val"]
        when "showNegBubbles"
          @show_neg_bubbles = attributes["val"] == "1" if attributes["val"]
        when "sizeRepresents"
          @size_represents = attributes["val"] if attributes["val"]
        when "firstSliceAng"
          @first_slice_ang = attributes["val"]&.to_i if attributes["val"]
        when "holeSize"
          @hole_size = attributes["val"]&.to_i if attributes["val"]
        when "smooth"
          @smooth = attributes["val"] == "1" if attributes["val"] && !@inside_ser
        when "marker"
          @marker = attributes["val"] == "1" if attributes["val"] && !@inside_ser
        when "scatterStyle"
          @scatter_style = attributes["val"] if attributes["val"]
        when "radarStyle"
          @radar_style = attributes["val"] if attributes["val"]
        when "wireframe"
          @wireframe = attributes["val"] == "1" if attributes["val"]
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
        when "overlay"
          @legend[:overlay] = attributes["val"] == "1" if @inside_legend && attributes["val"]
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
        when "dLblPos"
          @data_labels[:position] = attributes["val"] if @inside_dlbls && attributes["val"]
        when "showBubbleSize"
          @data_labels[:show_bubble_size] = attributes["val"] == "1" if @inside_dlbls
        when "separator"
          if @inside_dlbls
            @inside_separator = true
            @text_buffer = +""
          end
        when "catAx"
          @inside_cat_ax = true
        when "valAx"
          @inside_val_ax = true
        when "scaling"
          @inside_scaling = true if @inside_cat_ax || @inside_val_ax
        when "logBase"
          if @inside_scaling && attributes["val"]
            if @inside_cat_ax
              @cat_axis_log_base = attributes["val"].to_f
            elsif @inside_val_ax
              @val_axis_log_base = attributes["val"].to_f
            end
          end
        when "max"
          if @inside_scaling && attributes["val"]
            if @inside_cat_ax
              @cat_axis_scaling_max = attributes["val"].to_f
            elsif @inside_val_ax
              @val_axis_scaling_max = attributes["val"].to_f
            end
          end
        when "min"
          if @inside_scaling && attributes["val"]
            if @inside_cat_ax
              @cat_axis_scaling_min = attributes["val"].to_f
            elsif @inside_val_ax
              @val_axis_scaling_min = attributes["val"].to_f
            end
          end
        when "delete"
          if attributes["val"]
            if @inside_cat_ax
              @cat_axis_delete = attributes["val"] == "1"
            elsif @inside_val_ax
              @val_axis_delete = attributes["val"] == "1"
            end
          end
        when "orientation"
          if attributes["val"]
            if @inside_cat_ax
              @cat_axis_orientation = attributes["val"]
            elsif @inside_val_ax
              @val_axis_orientation = attributes["val"]
            end
          end
        when "numFmt"
          if (@inside_cat_ax || @inside_val_ax) && attributes["formatCode"]
            nf = { format_code: attributes["formatCode"] }
            nf[:source_linked] = attributes["sourceLinked"] == "1" if attributes["sourceLinked"]
            if @inside_cat_ax
              @cat_axis_num_fmt = nf
            elsif @inside_val_ax
              @val_axis_num_fmt = nf
            end
          end
        when "majorTickMark"
          if attributes["val"]
            if @inside_cat_ax
              @cat_axis_major_tick_mark = attributes["val"]
            elsif @inside_val_ax
              @val_axis_major_tick_mark = attributes["val"]
            end
          end
        when "minorTickMark"
          if attributes["val"]
            if @inside_cat_ax
              @cat_axis_minor_tick_mark = attributes["val"]
            elsif @inside_val_ax
              @val_axis_minor_tick_mark = attributes["val"]
            end
          end
        when "crosses"
          if attributes["val"]
            if @inside_cat_ax
              @cat_axis_crosses = attributes["val"]
            elsif @inside_val_ax
              @val_axis_crosses = attributes["val"]
            end
          end
        when "crossesAt"
          if attributes["val"]
            if @inside_cat_ax
              @cat_axis_crosses_at = attributes["val"]&.to_f
            elsif @inside_val_ax
              @val_axis_crosses_at = attributes["val"]&.to_f
            end
          end
        when "tickLblSkip"
          @cat_axis_tick_lbl_skip = attributes["val"]&.to_i if attributes["val"] && @inside_cat_ax
        when "tickMarkSkip"
          @cat_axis_tick_mark_skip = attributes["val"]&.to_i if attributes["val"] && @inside_cat_ax
        when "lblOffset"
          @cat_axis_lbl_offset = attributes["val"]&.to_i if attributes["val"] && @inside_cat_ax
        when "noMultiLvlLbl"
          @cat_axis_no_multi_lvl_lbl = attributes["val"] == "1" if attributes["val"] && @inside_cat_ax
        when "axPos"
          if attributes["val"]
            if @inside_cat_ax
              @cat_axis_pos = attributes["val"]
            elsif @inside_val_ax
              @val_axis_pos = attributes["val"]
            end
          end
        when "crossBetween"
          @val_axis_cross_between = attributes["val"] if @inside_val_ax && attributes["val"]
        when "majorUnit"
          @val_axis_major_unit = attributes["val"].to_f if @inside_val_ax && attributes["val"]
        when "minorUnit"
          @val_axis_minor_unit = attributes["val"].to_f if @inside_val_ax && attributes["val"]
        when "builtInUnit"
          @val_axis_disp_units = attributes["val"] if @inside_val_ax && attributes["val"]
        when "tickLblPos"
          if attributes["val"]
            if @inside_cat_ax
              @cat_axis_tick_lbl_pos = attributes["val"]
            elsif @inside_val_ax
              @val_axis_tick_lbl_pos = attributes["val"]
            end
          end
        when "majorGridlines"
          if @inside_cat_ax
            @cat_axis_major_gridlines = true
          elsif @inside_val_ax
            @val_axis_major_gridlines = true
          end
        when "minorGridlines"
          if @inside_cat_ax
            @cat_axis_minor_gridlines = true
          elsif @inside_val_ax
            @val_axis_minor_gridlines = true
          end
        when "plotVisOnly"
          @plot_vis_only = attributes["val"] == "1" if attributes["val"]
        when "dispBlanksAs"
          @disp_blanks_as = attributes["val"] if attributes["val"]
        when "style"
          @style = attributes["val"]&.to_i if attributes["val"]
        when "roundedCorners"
          @rounded_corners = attributes["val"] == "1" if attributes["val"]
        when "showDLblsOverMax"
          @show_d_lbls_over_max = attributes["val"] == "1" if attributes["val"]
        end
      end

      def characters(text)
        @text_buffer << text if @inside_t || @inside_f || @inside_separator
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
        when "separator"
          if @inside_separator
            @data_labels[:separator] = @text_buffer.dup
            @inside_separator = false
          end
        when "catAx"
          @inside_cat_ax = false
        when "valAx"
          @inside_val_ax = false
        when "scaling"
          @inside_scaling = false
        when "view3D"
          @inside_view_3d = false
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
        @inside_r = false
        @inside_rpr = false
        @inside_t = false
        @current_comment = nil
        @text_buffer = +""
        @runs = []
        @current_font = {}
        @has_runs = false
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
          @current_comment[:guid] = attributes["guid"] if attributes["guid"]
          sid = attributes["shapeId"]
          @current_comment[:shape_id] = sid.to_i if sid
        when "text"
          if @inside_comment
            @inside_text = true
            @text_buffer = +""
            @runs = []
            @has_runs = false
          end
        when "r"
          if @inside_text
            @inside_r = true
            @has_runs = true
            @current_font = {}
          end
        when "rPr"
          @inside_rpr = true if @inside_r
        when "b"
          @current_font[:bold] = true if @inside_rpr
        when "i"
          @current_font[:italic] = true if @inside_rpr
        when "strike"
          @current_font[:strike] = true if @inside_rpr
        when "u"
          if @inside_rpr
            val = attributes["val"]
            @current_font[:underline] = val || true
          end
        when "vertAlign"
          @current_font[:vert_align] = attributes["val"] if @inside_rpr && attributes["val"]
        when "sz"
          @current_font[:sz] = attributes["val"]&.to_f if @inside_rpr
        when "color"
          if @inside_rpr
            if attributes["rgb"]
              @current_font[:color] = attributes["rgb"]
            elsif attributes["theme"]
              @current_font[:theme] = attributes["theme"].to_i
              @current_font[:tint] = attributes["tint"].to_f if attributes["tint"]
            elsif attributes["indexed"]
              @current_font[:indexed] = attributes["indexed"].to_i
            end
          end
        when "rFont"
          @current_font[:name] = attributes["val"] if @inside_rpr
        when "family"
          @current_font[:family] = attributes["val"]&.to_i if @inside_rpr
        when "scheme"
          @current_font[:scheme] = attributes["val"] if @inside_rpr
        when "t"
          if @inside_text
            @inside_t = true
            @text_buffer = +"" if @inside_r
          end
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
          if @current_comment && @inside_text
            if @has_runs && @runs.any? { |r| r[:font] }
              @current_comment[:text] = Xlsxrb::RichText.new(runs: @runs)
            else
              plain = @has_runs ? @runs.map { |r| r[:text] }.join : @text_buffer.dup
              @current_comment[:text] = plain
            end
          end
          @inside_text = false
        when "t"
          @inside_t = false
        when "rPr"
          @inside_rpr = false
        when "r"
          if @inside_r
            run = { text: @text_buffer.dup }
            run[:font] = @current_font.dup unless @current_font.empty?
            @runs << run
            @inside_r = false
          end
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
          @pivot_table[:data_caption] = attributes["dataCaption"] if attributes["dataCaption"]
          @pivot_table[:data_on_rows] = attributes["dataOnRows"] == "1" if attributes["dataOnRows"]
          @pivot_table[:row_grand_totals] = attributes["rowGrandTotals"] != "0" if attributes["rowGrandTotals"]
          @pivot_table[:col_grand_totals] = attributes["colGrandTotals"] != "0" if attributes["colGrandTotals"]
          @pivot_table[:compact] = attributes["compact"] != "0" if attributes["compact"]
          @pivot_table[:outline] = attributes["outline"] != "0" if attributes["outline"]
          @pivot_table[:outline_data] = attributes["outlineData"] == "1" if attributes["outlineData"]
          @pivot_table[:compact_data] = attributes["compactData"] != "0" if attributes["compactData"]
          @pivot_table[:show_headers] = attributes["showHeaders"] != "0" if attributes["showHeaders"]
          @pivot_table[:show_multiple_label] = attributes["showMultipleLabel"] != "0" if attributes["showMultipleLabel"]
          @pivot_table[:show_data_drop_down] = attributes["showDataDropDown"] != "0" if attributes["showDataDropDown"]
          @pivot_table[:grand_total_caption] = attributes["grandTotalCaption"] if attributes["grandTotalCaption"]
          @pivot_table[:error_caption] = attributes["errorCaption"] if attributes["errorCaption"]
          @pivot_table[:show_error] = attributes["showError"] == "1" if attributes["showError"]
          @pivot_table[:missing_caption] = attributes["missingCaption"] if attributes["missingCaption"]
          @pivot_table[:show_missing] = attributes["showMissing"] != "0" if attributes["showMissing"]
          @pivot_table[:tag] = attributes["tag"] if attributes["tag"]
          @pivot_table[:indent] = attributes["indent"]&.to_i if attributes["indent"]
          @pivot_table[:published] = attributes["published"] == "1" if attributes["published"]
          @pivot_table[:edit_data] = attributes["editData"] == "1" if attributes["editData"]
          @pivot_table[:disable_field_list] = attributes["disableFieldList"] == "1" if attributes["disableFieldList"]
          @pivot_table[:visual_totals] = attributes["visualTotals"] != "0" if attributes["visualTotals"]
          @pivot_table[:print_drill] = attributes["printDrill"] == "1" if attributes["printDrill"]
          @pivot_table[:created_version] = attributes["createdVersion"]&.to_i if attributes["createdVersion"]
          @pivot_table[:updated_version] = attributes["updatedVersion"]&.to_i if attributes["updatedVersion"]
          @pivot_table[:min_refreshable_version] = attributes["minRefreshableVersion"]&.to_i if attributes["minRefreshableVersion"]
          %w[applyNumberFormats applyBorderFormats applyFontFormats
             applyPatternFormats applyAlignmentFormats applyWidthHeightFormats].each do |attr|
            next if attributes[attr].nil?

            key = attr.gsub(/[A-Z]/) { |m| "_#{m.downcase}" }.to_sym
            @pivot_table[key] = %w[1 true].include?(attributes[attr])
          end
          mff = attributes["multipleFieldFilters"]
          @pivot_table[:multiple_field_filters] = mff != "0" unless mff.nil?
          sdr = attributes["showDrill"]
          @pivot_table[:show_drill] = sdr != "0" unless sdr.nil?
          sdt = attributes["showDataTips"]
          @pivot_table[:show_data_tips] = sdt != "0" unless sdt.nil?
          edr = attributes["enableDrill"]
          @pivot_table[:enable_drill] = edr != "0" unless edr.nil?
          smpt = attributes["showMemberPropertyTips"]
          @pivot_table[:show_member_property_tips] = smpt != "0" unless smpt.nil?
          ipt = attributes["itemPrintTitles"]
          @pivot_table[:item_print_titles] = ipt == "1" unless ipt.nil?
          fpt = attributes["fieldPrintTitles"]
          @pivot_table[:field_print_titles] = fpt == "1" unless fpt.nil?
          pf = attributes["preserveFormatting"]
          @pivot_table[:preserve_formatting] = pf != "0" unless pf.nil?
          potd = attributes["pageOverThenDown"]
          @pivot_table[:page_over_then_down] = potd == "1" unless potd.nil?
          pw = attributes["pageWrap"]
          @pivot_table[:page_wrap] = pw.to_i if pw
        when "location"
          @pivot_table[:ref] = attributes["ref"] if @pivot_table
          @pivot_table[:row_page_count] = attributes["rowPageCount"]&.to_i if attributes["rowPageCount"]
          @pivot_table[:col_page_count] = attributes["colPageCount"]&.to_i if attributes["colPageCount"]
        when "pivotField"
          @inside_pivot_field = true
          @current_field = {}
          @current_field[:axis] = attributes["axis"] if attributes["axis"]
          @current_field[:data_field] = true if attributes["dataField"] == "1"
          @current_field[:name] = attributes["name"] if attributes["name"]
          @current_field[:show_all] = attributes["showAll"] != "0" if attributes["showAll"]
          @current_field[:compact] = attributes["compact"] != "0" if attributes["compact"]
          @current_field[:outline] = attributes["outline"] != "0" if attributes["outline"]
          @current_field[:subtotal_top] = attributes["subtotalTop"] != "0" if attributes["subtotalTop"]
          @current_field[:num_fmt_id] = attributes["numFmtId"]&.to_i if attributes["numFmtId"]
          @current_field[:sort_type] = attributes["sortType"] if attributes["sortType"]
          ds = attributes["defaultSubtotal"]
          @current_field[:default_subtotal] = ds != "0" unless ds.nil?
          @current_field[:insert_blank_row] = true if attributes["insertBlankRow"] == "1"
          @current_field[:insert_page_break] = true if attributes["insertPageBreak"] == "1"
          @current_field[:include_new_items_in_filter] = true if attributes["includeNewItemsInFilter"] == "1"
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
            df = {
              name: attributes["name"],
              fld: attributes["fld"]&.to_i,
              subtotal: attributes["subtotal"] || "sum"
            }
            df[:show_data_as] = attributes["showDataAs"] if attributes["showDataAs"]
            df[:base_field] = attributes["baseField"]&.to_i if attributes["baseField"]
            df[:base_item] = attributes["baseItem"]&.to_i if attributes["baseItem"]
            df[:num_fmt_id] = attributes["numFmtId"]&.to_i if attributes["numFmtId"]
            @data_fields << df
          end
        when "pivotTableStyleInfo"
          if @pivot_table
            psi = {}
            psi[:name] = attributes["name"] if attributes["name"]
            srh = attributes["showRowHeaders"]
            psi[:show_row_headers] = srh == "1" unless srh.nil?
            sch = attributes["showColHeaders"]
            psi[:show_col_headers] = sch == "1" unless sch.nil?
            srs = attributes["showRowStripes"]
            psi[:show_row_stripes] = srs == "1" unless srs.nil?
            scs = attributes["showColStripes"]
            psi[:show_col_stripes] = scs == "1" unless scs.nil?
            slc = attributes["showLastColumn"]
            psi[:show_last_column] = slc == "1" unless slc.nil?
            @pivot_table[:pivot_table_style] = psi
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

    # SAX2 listener for parsing pivotCacheDefinition XML.
    class PivotCacheDefinitionListener
      include REXML::SAX2Listener

      attr_reader :cache_definition

      def initialize
        @cache_definition = {}
        @fields = []
        @current_field = nil
        @current_shared_items = nil
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        case name
        when "pivotCacheDefinition"
          sd = attributes["saveData"]
          @cache_definition[:save_data] = sd != "0" unless sd.nil?
          er = attributes["enableRefresh"]
          @cache_definition[:enable_refresh] = er != "0" unless er.nil?
          @cache_definition[:refreshed_by] = attributes["refreshedBy"] if attributes["refreshedBy"]
          @cache_definition[:refreshed_version] = attributes["refreshedVersion"]&.to_i if attributes["refreshedVersion"]
          @cache_definition[:created_version] = attributes["createdVersion"]&.to_i if attributes["createdVersion"]
          @cache_definition[:record_count] = attributes["recordCount"]&.to_i if attributes["recordCount"]
          om = attributes["optimizeMemory"]
          @cache_definition[:optimize_memory] = om == "1" unless om.nil?
        when "cacheSource"
          @cache_definition[:source_type] = attributes["type"] if attributes["type"]
        when "worksheetSource"
          @cache_definition[:source_ref] = attributes["ref"] if attributes["ref"]
          @cache_definition[:source_sheet] = attributes["sheet"] if attributes["sheet"]
          @cache_definition[:source_name] = attributes["name"] if attributes["name"]
        when "cacheField"
          @current_field = {}
          @current_field[:name] = attributes["name"] if attributes["name"]
          @current_field[:num_fmt_id] = attributes["numFmtId"]&.to_i if attributes["numFmtId"]
          @current_field[:caption] = attributes["caption"] if attributes["caption"]
          @current_field[:formula] = xml_unescape(attributes["formula"]) if attributes["formula"]
        when "sharedItems"
          @current_shared_items = [] if @current_field
        when "s", "d", "e"
          @current_shared_items << attributes["v"] if @current_shared_items && attributes["v"]
        when "n"
          @current_shared_items << attributes["v"]&.to_f if @current_shared_items && attributes["v"]
        when "b"
          @current_shared_items << (attributes["v"] == "1") if @current_shared_items
        when "m"
          @current_shared_items << nil if @current_shared_items
        end
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        case name
        when "sharedItems"
          @current_field[:shared_items] = @current_shared_items if @current_field && @current_shared_items && !@current_shared_items.empty?
          @current_shared_items = nil
        when "cacheField"
          if @current_field
            @fields << @current_field
            @current_field = nil
          end
        end
      end

      def characters(_text); end

      def end_document
        @cache_definition[:fields] = @fields unless @fields.empty?
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

    # SAX2 listener for parsing pivotCacheRecords XML.
    class PivotCacheRecordsListener
      include REXML::SAX2Listener

      attr_reader :records

      def initialize
        @records = []
        @current_record = nil
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        case name
        when "r"
          @current_record = []
        when "x"
          @current_record << { x: attributes["v"]&.to_i } if @current_record
        when "s", "d", "e"
          @current_record << attributes["v"] if @current_record && attributes["v"]
        when "n"
          @current_record << attributes["v"]&.to_f if @current_record && attributes["v"]
        when "b"
          @current_record << (attributes["v"] == "1") if @current_record
        when "m"
          @current_record << nil if @current_record
        end
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        return unless name == "r" && @current_record

        @records << @current_record
        @current_record = nil
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

    # SAX2 listener for parsing externalLink XML.
    class ExternalLinkListener
      include REXML::SAX2Listener

      attr_reader :sheet_names

      def initialize
        @sheet_names = []
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        return unless name == "sheetName"

        @sheet_names << attributes["val"] if attributes["val"]
      end

      def end_element(_uri, _local_name, _qname); end

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
