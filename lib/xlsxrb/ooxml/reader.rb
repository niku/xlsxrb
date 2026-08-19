# frozen_string_literal: true

# rbs_inline: enabled

require "zlib"
require "rexml/parsers/sax2parser"
require "rexml/sax2listener"

module Xlsxrb
  module Ooxml
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

      # Returns cell addresses marked as phonetic: { "A1" => true }.
      def cell_phonetic(sheet: nil)
        worksheet_xml = load_worksheet_xml(sheet)
        return {} if worksheet_xml.nil? || worksheet_xml.empty?

        parse_worksheet_cell_phonetic(worksheet_xml)
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
        links = []
        WorksheetParser.each_event(worksheet_xml, part_name: entry_path) do |event|
          next unless event.type == :hyperlink

          ref, rid, display, tooltip, location = event.args
          link = { ref: ref }
          link[:rid] = rid if rid
          link[:display] = display if display
          link[:tooltip] = tooltip if tooltip
          link[:location] = location if location
          links << link
        end

        # Parse rels to resolve rId -> URL.
        rels_path = entry_path.sub(%r{([^/]+)$}, '_rels/\1.rels')
        rels_xml = extract_zip_entry(rels_path)
        rid_to_url = {}
        rid_to_url = parse_rels(rels_xml).transform_values { |v| v } if rels_xml && !rels_xml.empty?

        result = {}
        links.each do |link|
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
          chart[:title_overlay] = cl.title_overlay unless cl.title_overlay.nil?
          chart[:title_font] = cl.title_font if cl.title_font
          chart[:title_fill_color] = cl.title_fill_color if cl.title_fill_color
          chart[:title_no_fill] = cl.title_no_fill if cl.title_no_fill
          chart[:title_line_color] = cl.title_line_color if cl.title_line_color
          chart[:title_line_width] = cl.title_line_width if cl.title_line_width
          chart[:title_line_dash] = cl.title_line_dash if cl.title_line_dash
          chart[:series] = cl.series unless cl.series.empty?
          chart[:legend] = cl.legend unless cl.legend.empty?
          if cl.legend_font
            chart[:legend] ||= {}
            chart[:legend][:font] = cl.legend_font
          end
          chart[:data_labels] = cl.data_labels unless cl.data_labels.empty?
          chart[:cat_axis_title] = cl.cat_axis_title if cl.cat_axis_title
          chart[:val_axis_title] = cl.val_axis_title if cl.val_axis_title
          chart[:cat_axis_title_font] = cl.cat_axis_title_font if cl.cat_axis_title_font
          chart[:val_axis_title_font] = cl.val_axis_title_font if cl.val_axis_title_font
          chart[:cat_axis_title_fill] = cl.cat_axis_title_fill if cl.cat_axis_title_fill
          chart[:cat_axis_title_no_fill] = cl.cat_axis_title_no_fill if cl.cat_axis_title_no_fill
          chart[:cat_axis_title_line_color] = cl.cat_axis_title_line_color if cl.cat_axis_title_line_color
          chart[:cat_axis_title_line_width] = cl.cat_axis_title_line_width if cl.cat_axis_title_line_width
          chart[:cat_axis_title_line_dash] = cl.cat_axis_title_line_dash if cl.cat_axis_title_line_dash
          chart[:val_axis_title_fill] = cl.val_axis_title_fill if cl.val_axis_title_fill
          chart[:val_axis_title_no_fill] = cl.val_axis_title_no_fill if cl.val_axis_title_no_fill
          chart[:val_axis_title_line_color] = cl.val_axis_title_line_color if cl.val_axis_title_line_color
          chart[:val_axis_title_line_width] = cl.val_axis_title_line_width if cl.val_axis_title_line_width
          chart[:val_axis_title_line_dash] = cl.val_axis_title_line_dash if cl.val_axis_title_line_dash
          chart[:title_layout] = cl.title_layout if cl.title_layout
          chart[:cat_axis_title_layout] = cl.cat_axis_title_layout if cl.cat_axis_title_layout
          chart[:val_axis_title_layout] = cl.val_axis_title_layout if cl.val_axis_title_layout
          chart[:title_rotation] = cl.title_rotation if cl.title_rotation
          chart[:cat_axis_title_rotation] = cl.cat_axis_title_rotation if cl.cat_axis_title_rotation
          chart[:val_axis_title_rotation] = cl.val_axis_title_rotation if cl.val_axis_title_rotation
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
          chart[:cat_axis_auto] = cl.cat_axis_auto unless cl.cat_axis_auto.nil?
          chart[:cat_axis_lbl_algn] = cl.cat_axis_lbl_algn if cl.cat_axis_lbl_algn
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
          chart[:drop_lines] = cl.drop_lines unless cl.drop_lines.nil?
          chart[:hi_low_lines] = cl.hi_low_lines unless cl.hi_low_lines.nil?
          chart[:ser_lines] = cl.ser_lines unless cl.ser_lines.nil?
          chart[:up_down_bars] = cl.up_down_bars if cl.up_down_bars
          chart[:scatter_style] = cl.scatter_style if cl.scatter_style
          chart[:radar_style] = cl.radar_style if cl.radar_style
          chart[:cat_axis_pos] = cl.cat_axis_pos if cl.cat_axis_pos
          chart[:val_axis_pos] = cl.val_axis_pos if cl.val_axis_pos
          chart[:wireframe] = cl.wireframe unless cl.wireframe.nil?
          chart[:band_fmts] = cl.band_fmts if cl.band_fmts
          chart[:of_pie_type] = cl.of_pie_type if cl.of_pie_type
          chart[:split_type] = cl.split_type if cl.split_type
          chart[:split_pos] = cl.split_pos if cl.split_pos
          chart[:cust_split] = cl.cust_split if cl.cust_split&.any?
          chart[:second_pie_size] = cl.second_pie_size if cl.second_pie_size
          chart[:data_table] = cl.data_table if cl.data_table
          chart[:plot_area_fill] = cl.plot_area_fill if cl.plot_area_fill
          chart[:plot_area_line_color] = cl.plot_area_line_color if cl.plot_area_line_color
          chart[:plot_area_line_width] = cl.plot_area_line_width if cl.plot_area_line_width
          chart[:plot_area_line_dash] = cl.plot_area_line_dash if cl.plot_area_line_dash
          chart[:plot_area_no_fill] = cl.plot_area_no_fill if cl.plot_area_no_fill
          chart[:plot_area_layout] = cl.plot_area_layout if cl.plot_area_layout
          chart[:cat_axis_label_rotation] = cl.cat_axis_label_rotation if cl.cat_axis_label_rotation
          chart[:val_axis_label_rotation] = cl.val_axis_label_rotation if cl.val_axis_label_rotation
          chart[:cat_axis_font] = cl.cat_axis_font if cl.cat_axis_font
          chart[:val_axis_font] = cl.val_axis_font if cl.val_axis_font
          chart[:cat_axis_fill] = cl.cat_axis_fill if cl.cat_axis_fill
          chart[:cat_axis_no_fill] = cl.cat_axis_no_fill if cl.cat_axis_no_fill
          chart[:val_axis_fill] = cl.val_axis_fill if cl.val_axis_fill
          chart[:val_axis_no_fill] = cl.val_axis_no_fill if cl.val_axis_no_fill
          chart[:cat_axis_line_color] = cl.cat_axis_line_color if cl.cat_axis_line_color
          chart[:cat_axis_line_width] = cl.cat_axis_line_width if cl.cat_axis_line_width
          chart[:cat_axis_line_dash] = cl.cat_axis_line_dash if cl.cat_axis_line_dash
          chart[:val_axis_line_color] = cl.val_axis_line_color if cl.val_axis_line_color
          chart[:val_axis_line_width] = cl.val_axis_line_width if cl.val_axis_line_width
          chart[:val_axis_line_dash] = cl.val_axis_line_dash if cl.val_axis_line_dash
          chart[:floor] = cl.floor if cl.floor
          chart[:side_wall] = cl.side_wall if cl.side_wall
          chart[:back_wall] = cl.back_wall if cl.back_wall
          chart[:cat_axis_type] = cl.cat_axis_type if cl.cat_axis_type
          chart[:cat_axis_base_time_unit] = cl.cat_axis_base_time_unit if cl.cat_axis_base_time_unit
          chart[:cat_axis_major_time_unit] = cl.cat_axis_major_time_unit if cl.cat_axis_major_time_unit
          chart[:cat_axis_minor_time_unit] = cl.cat_axis_minor_time_unit if cl.cat_axis_minor_time_unit
          chart[:cat_axis_major_unit] = cl.cat_axis_major_unit if cl.cat_axis_major_unit
          chart[:cat_axis_minor_unit] = cl.cat_axis_minor_unit if cl.cat_axis_minor_unit
          chart[:chart_fill] = cl.chart_fill if cl.chart_fill
          chart[:chart_no_fill] = cl.chart_no_fill if cl.chart_no_fill
          chart[:chart_line_color] = cl.chart_line_color if cl.chart_line_color
          chart[:chart_line_width] = cl.chart_line_width if cl.chart_line_width
          chart[:chart_line_dash] = cl.chart_line_dash if cl.chart_line_dash
          chart[:protection] = cl.protection if cl.protection
          chart[:print_settings] = cl.print_settings if cl.print_settings
          chart[:chart_font] = cl.chart_font if cl.chart_font
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
                                  Xlsxrb::Ooxml::Utils.serial_to_datetime(value)
                                else
                                  Xlsxrb::Ooxml::Utils.serial_to_date(value.to_i)
                                end
        end
        raw_cells
      end

      def date_format?(fmt_id, custom_num_fmts)
        return false unless fmt_id

        # Built-in date format IDs.
        return true if Xlsxrb::Ooxml::Utils::BUILTIN_DATE_FMT_IDS.include?(fmt_id)

        # Check custom format code for date-like patterns.
        code = custom_num_fmts[fmt_id]
        return false unless code

        date_pattern?(code)
      end

      def resolve_num_fmt_code(fmt_id, custom_num_fmts)
        custom_num_fmts[fmt_id] || Xlsxrb::Ooxml::Utils::BUILTIN_NUM_FMT_CODES[fmt_id]
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

            raise ZipError, "invalid ZIP local header signature" unless signature_value == 0x04034b50

            header = file.read(26)
            raise ZipError, "truncated ZIP local header" if header.nil? || header.bytesize < 26

            _version, flags, compression_method, _mtime, _mdate, _crc32, compressed_size,
              _uncompressed_size, file_name_length, extra_field_length = header.unpack("v v v v v V V V v v")

            raise ZipError, "ZIP data descriptor is not supported" if flags.anybits?(0x0008)

            file_name = file.read(file_name_length)
            raise ZipError, "truncated ZIP file name" if file_name.nil? || file_name.bytesize < file_name_length

            file.read(extra_field_length)

            compressed_data = file.read(compressed_size)
            raise ZipError, "truncated ZIP entry data" if compressed_data.nil? || compressed_data.bytesize < compressed_size

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
              raise ZipError, "unsupported ZIP compression method: #{compression_method}"
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

      def parse_worksheet_cell_phonetic(xml)
        parser = REXML::Parsers::SAX2Parser.new(xml)
        listener = WorksheetListener.new([])
        parser.listen(listener)
        parser.parse
        listener.cell_phonetic
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
    end
  end
end

require_relative "reader/listeners"
