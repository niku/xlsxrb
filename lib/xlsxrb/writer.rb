# frozen_string_literal: true

require_relative "zip_generator"

module Xlsxrb
  # Writes spreadsheet data into an XLSX file.
  class Writer
    XML_HEADER = %(<?xml version="1.0" encoding="UTF-8"?>)
    CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
    REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
    SSML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    DC_NS = "http://purl.org/dc/elements/1.1/"
    DCTERMS_NS = "http://purl.org/dc/terms/"
    XSI_NS = "http://www.w3.org/2001/XMLSchema-instance"
    APP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
    VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    CELL_ADDRESS_PATTERN = /\A([A-Z]{1,3})(\d+)\z/
    MAX_ROW = 1_048_576
    MAX_COLUMN_INDEX = 16_384 # XFD

    def initialize
      @sheets = { "Sheet1" => {} }
      @column_widths = { "Sheet1" => {} }
      @column_attrs = { "Sheet1" => {} }
      @row_attrs = { "Sheet1" => {} }
      @merge_cells = { "Sheet1" => [] }
      @hyperlinks = { "Sheet1" => {} }
      @cell_styles = { "Sheet1" => {} }
      @cell_phonetic = { "Sheet1" => {} }
      @auto_filters = { "Sheet1" => nil }
      @filter_columns = { "Sheet1" => {} }
      @sort_state = { "Sheet1" => nil }
      @num_fmts = []
      @fonts = [{ sz: 11, name: "Calibri" }]
      @fills = [{ pattern: "none" }, { pattern: "gray125" }]
      @borders = [{ left: nil, right: nil, top: nil, bottom: nil }]
      @xf_entries = [{ num_fmt_id: 0, font_id: 0, fill_id: 0, border_id: 0 }]
      @cell_style_xfs = [{ num_fmt_id: 0, font_id: 0, fill_id: 0, border_id: 0 }]
      @cell_style_names = [{ name: "Normal", xf_id: 0, builtin_id: 0 }]
      @dxfs = []
      @indexed_colors = []
      @mru_colors = []
      @table_styles = {}
      @sheet_order = ["Sheet1"]
      @core_properties = {}
      @app_properties = {}
      @custom_properties = []
      @file_version = {}
      @file_sharing = {}
      @workbook_properties = {}
      @workbook_views = {}
      @calc_properties = {}
      @file_recovery_properties = {}
      @sheet_states = {}
      @defined_names = []
      @sheet_properties = { "Sheet1" => {} }
      @phonetic_properties = { "Sheet1" => nil }
      @sheet_formats = { "Sheet1" => {} }
      @sheet_views = { "Sheet1" => {} }
      @freeze_panes = { "Sheet1" => nil }
      @selections = { "Sheet1" => nil }
      @print_options = { "Sheet1" => {} }
      @page_margins = { "Sheet1" => nil }
      @page_setup = { "Sheet1" => {} }
      @header_footer = { "Sheet1" => {} }
      @row_breaks = { "Sheet1" => [] }
      @col_breaks = { "Sheet1" => [] }
      @data_validations = { "Sheet1" => [] }
      @data_validations_options = { "Sheet1" => {} }
      @conditional_formats = { "Sheet1" => [] }
      @tables = { "Sheet1" => [] }
      @cell_watches = { "Sheet1" => [] }
      @ignored_errors = { "Sheet1" => [] }
      @data_consolidate = { "Sheet1" => nil }
      @scenarios = { "Sheet1" => nil }
      @use_shared_strings = true
      @images = { "Sheet1" => [] }
      @charts_data = { "Sheet1" => [] }
      @shapes_data = { "Sheet1" => [] }
      @comments_data = { "Sheet1" => [] }
      @pivot_tables_data = { "Sheet1" => [] }
      @extra_entries = {}
      @extra_ct_defaults = {}
      @extra_ct_overrides = {}
      @preserve_macros = false
      @sheet_protection = { "Sheet1" => nil }
      @protected_ranges = { "Sheet1" => [] }
      @workbook_protection = nil
      @external_links = []
    end

    # Adds a new sheet. Raises if name is already taken.
    def add_sheet(name)
      raise ArgumentError, "sheet already exists: #{name}" if @sheets.key?(name)

      @sheets[name] = {}
      @column_widths[name] = {}
      @column_attrs[name] = {}
      @row_attrs[name] = {}
      @merge_cells[name] = []
      @hyperlinks[name] = {}
      @cell_styles[name] = {}
      @cell_phonetic[name] = {}
      @auto_filters[name] = nil
      @filter_columns[name] = {}
      @sort_state[name] = nil
      @sheet_properties[name] = {}
      @sheet_formats[name] = {}
      @sheet_views[name] = {}
      @freeze_panes[name] = nil
      @selections[name] = nil
      @print_options[name] = {}
      @page_margins[name] = nil
      @page_setup[name] = {}
      @header_footer[name] = {}
      @row_breaks[name] = []
      @col_breaks[name] = []
      @data_validations[name] = []
      @data_validations_options[name] = {}
      @conditional_formats[name] = []
      @tables[name] = []
      @cell_watches[name] = []
      @ignored_errors[name] = []
      @data_consolidate[name] = nil
      @scenarios[name] = nil
      @images[name] = []
      @charts_data[name] = []
      @shapes_data[name] = []
      @comments_data[name] = []
      @pivot_tables_data[name] = []
      @sheet_protection[name] = nil
      @protected_ranges[name] = []
      @phonetic_properties[name] = nil
      @sheet_order << name
    end

    # Registers a cell value at the given address (e.g. "A1").
    def set_cell(cell_address, value, sheet: nil)
      validate_cell_address!(cell_address)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @sheets.key?(sheet_name)

      @sheets[sheet_name][cell_address] = value
    end

    # Returns the registered cells for the first (or given) sheet.
    def cells(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @sheets[sheet_name] || {}
    end

    # Sets the width for a column (e.g. "A", "BC").
    def set_column_width(col_letter, width, sheet: nil)
      raise ArgumentError, "column must be a String of uppercase letters" unless col_letter.is_a?(String) && col_letter.match?(/\A[A-Z]+\z/)

      col_index = column_letter_to_index(col_letter)
      raise ArgumentError, "column out of range: #{col_letter}" unless col_index.between?(1, MAX_COLUMN_INDEX)

      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @column_widths.key?(sheet_name)

      @column_widths[sheet_name][col_letter] = width
    end

    # Returns column widths for the first (or given) sheet.
    def column_widths(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @column_widths[sheet_name] || {}
    end

    # Sets a column attribute (e.g. :hidden, :best_fit, :outline_level, :collapsed, :style).
    def set_column_attribute(col_letter, name, value, sheet: nil)
      raise ArgumentError, "column must be a String of uppercase letters" unless col_letter.is_a?(String) && col_letter.match?(/\A[A-Z]+\z/)

      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @column_attrs.key?(sheet_name)

      @column_attrs[sheet_name][col_letter] ||= {}
      @column_attrs[sheet_name][col_letter][name] = value
    end

    # Returns column attributes for the first (or given) sheet.
    def column_attributes(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @column_attrs[sheet_name] || {}
    end

    # Sets a row height.
    def set_row_height(row_num, height, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @row_attrs.key?(sheet_name)
      raise ArgumentError, "row must be a positive Integer" unless row_num.is_a?(Integer) && row_num >= 1

      @row_attrs[sheet_name][row_num] ||= {}
      @row_attrs[sheet_name][row_num][:height] = height
    end

    # Hides a row.
    def set_row_hidden(row_num, hidden: true, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @row_attrs.key?(sheet_name)
      raise ArgumentError, "row must be a positive Integer" unless row_num.is_a?(Integer) && row_num >= 1

      @row_attrs[sheet_name][row_num] ||= {}
      @row_attrs[sheet_name][row_num][:hidden] = hidden
    end

    # Sets a row outline level.
    def set_row_outline_level(row_num, level, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @row_attrs.key?(sheet_name)
      raise ArgumentError, "row must be a positive Integer" unless row_num.is_a?(Integer) && row_num >= 1

      @row_attrs[sheet_name][row_num] ||= {}
      @row_attrs[sheet_name][row_num][:outline_level] = level
    end

    # Sets a row collapsed state.
    def set_row_collapsed(row_num, collapsed: true, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @row_attrs.key?(sheet_name)
      raise ArgumentError, "row must be a positive Integer" unless row_num.is_a?(Integer) && row_num >= 1

      @row_attrs[sheet_name][row_num] ||= {}
      @row_attrs[sheet_name][row_num][:collapsed] = collapsed
    end

    # Sets a row default style index (cellXfs index).
    def set_row_style(row_num, style_id, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @row_attrs.key?(sheet_name)
      raise ArgumentError, "row must be a positive Integer" unless row_num.is_a?(Integer) && row_num >= 1
      raise ArgumentError, "style_id must be a non-negative Integer" unless style_id.is_a?(Integer) && style_id >= 0

      @row_attrs[sheet_name][row_num] ||= {}
      @row_attrs[sheet_name][row_num][:style] = style_id
    end

    # Sets a row thick top border flag.
    def set_row_thick_top(row_num, thick: true, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @row_attrs.key?(sheet_name)
      raise ArgumentError, "row must be a positive Integer" unless row_num.is_a?(Integer) && row_num >= 1

      @row_attrs[sheet_name][row_num] ||= {}
      @row_attrs[sheet_name][row_num][:thick_top] = thick
    end

    # Sets a row thick bottom border flag.
    def set_row_thick_bot(row_num, thick: true, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @row_attrs.key?(sheet_name)
      raise ArgumentError, "row must be a positive Integer" unless row_num.is_a?(Integer) && row_num >= 1

      @row_attrs[sheet_name][row_num] ||= {}
      @row_attrs[sheet_name][row_num][:thick_bot] = thick
    end

    # Sets a row phonetic flag.
    def set_row_phonetic(row_num, phonetic: true, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @row_attrs.key?(sheet_name)
      raise ArgumentError, "row must be a positive Integer" unless row_num.is_a?(Integer) && row_num >= 1

      @row_attrs[sheet_name][row_num] ||= {}
      @row_attrs[sheet_name][row_num][:ph] = phonetic
    end

    # Returns row attributes for the first (or given) sheet.
    def row_attributes(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @row_attrs[sheet_name] || {}
    end

    # Merges a range of cells (e.g. "A1:B2").
    def merge_cells(range, sheet: nil)
      raise ArgumentError, "range must be a String like 'A1:B2'" unless range.is_a?(String) && range.match?(/\A[A-Z]+\d+:[A-Z]+\d+\z/)

      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @merge_cells.key?(sheet_name)

      @merge_cells[sheet_name] << range
    end

    # Returns merged cell ranges for the first (or given) sheet.
    def merged_cells(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @merge_cells[sheet_name] || []
    end

    # Adds a hyperlink on a cell.
    def add_hyperlink(cell_address, url = nil, sheet: nil, display: nil, tooltip: nil, location: nil)
      validate_cell_address!(cell_address)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @hyperlinks.key?(sheet_name)
      raise ArgumentError, "url or location required" if url.nil? && location.nil?

      link = {}
      link[:url] = url if url
      link[:display] = display if display
      link[:tooltip] = tooltip if tooltip
      link[:location] = location if location
      @hyperlinks[sheet_name][cell_address] = link
    end

    # Returns hyperlinks for the first (or given) sheet.
    def hyperlinks(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @hyperlinks[sheet_name] || {}
    end

    # Sets an autoFilter range (e.g. "A1:B10") for the given sheet.
    def set_auto_filter(range, sheet: nil)
      raise ArgumentError, "range must be a String like 'A1:B10'" unless range.is_a?(String) && range.match?(/\A[A-Z]+\d+:[A-Z]+\d+\z/)

      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @auto_filters.key?(sheet_name)

      @auto_filters[sheet_name] = range
    end

    # Returns the autoFilter range for the first (or given) sheet.
    def auto_filter(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @auto_filters[sheet_name]
    end

    # Adds a filter column to the autoFilter.
    # col_id: 0-based column index within the autoFilter range.
    # filter: hash describing the filter, e.g.
    #   { type: :filters, values: ["A", "B"] }
    #   { type: :filters, blank: true }
    #   { type: :custom, operator: "greaterThan", val: "100" }
    #   { type: :custom, filters: [{ operator: "greaterThan", val: "10" }, { operator: "lessThan", val: "100" }], and: true }
    #   { type: :dynamic, dynamic_type: "today" }
    #   { type: :top10, top: true, percent: false, val: 10 }
    def add_filter_column(col_id, filter, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @filter_columns.key?(sheet_name)

      @filter_columns[sheet_name][col_id] = filter
    end

    # Returns filter columns for the first (or given) sheet.
    def filter_columns(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @filter_columns[sheet_name] || {}
    end

    # Sets a sort state for the sheet. ref: sort range, sort_conditions: array of { ref:, descending: }.
    # Options: column_sort:, case_sensitive:, sort_method:
    def set_sort_state(ref, sort_conditions, sheet: nil, **opts)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @sort_state.key?(sheet_name)

      ss = { ref: ref, sort_conditions: sort_conditions }
      %i[column_sort case_sensitive sort_method].each do |key|
        ss[key] = opts[key] if opts.key?(key)
      end
      @sort_state[sheet_name] = ss
    end

    # Returns sort state for the first (or given) sheet.
    def sort_state(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @sort_state[sheet_name]
    end

    # Registers a custom number format and returns its numFmtId (starting at 164).
    def add_number_format(format_code)
      existing = @num_fmts.find { |nf| nf[:format_code] == format_code }
      return existing[:num_fmt_id] if existing

      num_fmt_id = 164 + @num_fmts.size
      @num_fmts << { num_fmt_id: num_fmt_id, format_code: format_code }
      num_fmt_id
    end

    # Sets a number format on a cell. num_fmt_id is from add_number_format or a built-in id.
    def set_cell_format(cell_address, num_fmt_id, sheet: nil)
      validate_cell_address!(cell_address)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @cell_styles.key?(sheet_name)

      @cell_styles[sheet_name][cell_address] = num_fmt_id
    end

    # Registers a font and returns its font_id. Opts: bold, italic, sz, color, name.
    def add_font(**opts)
      existing = @fonts.index(opts)
      return existing if existing

      @fonts << opts
      @fonts.size - 1
    end

    # Registers a fill and returns its fill_id. Opts: pattern, fg_color, bg_color.
    def add_fill(**opts)
      existing = @fills.index(opts)
      return existing if existing

      @fills << opts
      @fills.size - 1
    end

    # Registers a border and returns its border_id.
    # Opts: left, right, top, bottom (each a hash with :style and optional :color).
    def add_border(**opts)
      existing = @borders.index(opts)
      return existing if existing

      @borders << opts
      @borders.size - 1
    end

    # Registers a cell style (xf entry) and returns its index for use with set_cell_style.
    # Opts: font_id, fill_id, border_id, num_fmt_id, alignment (hash with horizontal, vertical, wrap_text, text_rotation, indent, shrink_to_fit).
    def add_cell_style(**opts)
      entry = {
        num_fmt_id: opts[:num_fmt_id] || 0,
        font_id: opts[:font_id] || 0,
        fill_id: opts[:fill_id] || 0,
        border_id: opts[:border_id] || 0,
        xf_id: opts[:xf_id] || 0
      }
      entry[:alignment] = opts[:alignment] if opts[:alignment]
      entry[:protection] = opts[:protection] if opts[:protection]
      entry[:quote_prefix] = true if opts[:quote_prefix]
      entry[:pivot_button] = true if opts[:pivot_button]
      existing = @xf_entries.index(entry)
      return existing if existing

      @xf_entries << entry
      @xf_entries.size - 1
    end

    # Registers a base style definition (cellStyleXf) and a named cellStyle.
    # Returns the xfId for the new base style.
    def add_named_cell_style(name:, num_fmt_id: 0, font_id: 0, fill_id: 0, border_id: 0, builtin_id: nil,
                             i_level: nil, hidden: nil, custom_builtin: nil)
      entry = { num_fmt_id: num_fmt_id, font_id: font_id, fill_id: fill_id, border_id: border_id }
      @cell_style_xfs << entry
      xf_id = @cell_style_xfs.size - 1
      cs = { name: name, xf_id: xf_id }
      cs[:builtin_id] = builtin_id if builtin_id
      cs[:i_level] = i_level if i_level
      cs[:hidden] = true if hidden
      cs[:custom_builtin] = true if custom_builtin
      @cell_style_names << cs
      xf_id
    end

    # Sets a cell style by xf index (from add_cell_style).
    def set_cell_style(cell_address, style_id, sheet: nil)
      validate_cell_address!(cell_address)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @cell_styles.key?(sheet_name)

      @cell_styles[sheet_name][cell_address] = { xf_index: style_id }
    end

    # Marks a cell as containing phonetic text (ph="1" on the <c> element).
    def set_cell_phonetic(cell_address, sheet: nil)
      validate_cell_address!(cell_address)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @cell_phonetic.key?(sheet_name)

      @cell_phonetic[sheet_name][cell_address] = true
    end

    # Registers a differential format (dxf) for conditional formatting. Returns dxf_id.
    # Opts: font (hash), fill (hash), border (hash), num_fmt (hash).
    def add_dxf(**opts)
      @dxfs << opts
      @dxfs.size - 1
    end

    # Sets the indexed colors palette (array of ARGB hex strings, e.g. ["FF000000", "FFFFFFFF"]).
    # rubocop:disable Naming/AccessorMethodName
    def set_indexed_colors(colors)
      @indexed_colors = colors
    end
    # rubocop:enable Naming/AccessorMethodName

    # Returns the indexed colors palette.
    def indexed_colors
      @indexed_colors.dup
    end

    # Sets the MRU (most recently used) colors (array of color hashes, e.g. [{rgb: "FFFF0000"}]).
    # rubocop:disable Naming/AccessorMethodName
    def set_mru_colors(colors)
      @mru_colors = colors
    end
    # rubocop:enable Naming/AccessorMethodName

    # Returns the MRU colors.
    def mru_colors
      @mru_colors.map(&:dup)
    end

    # Sets table styles options (defaultTableStyle, defaultPivotStyle).
    def set_table_styles_option(name, value)
      raise ArgumentError, "name must be a Symbol" unless name.is_a?(Symbol)

      @table_styles[name] = value
    end

    # Adds a table style definition. Returns the style name.
    # elements: array of { type:, dxf_id:, size: }
    def add_table_style(name:, elements: [], pivot: nil, table: nil)
      @table_styles[:styles] ||= []
      style = { name: name, elements: elements }
      style[:pivot] = pivot unless pivot.nil?
      style[:table] = table unless table.nil?
      @table_styles[:styles] << style
      name
    end

    # Returns table styles configuration.
    def table_styles
      deep_copy = {}
      @table_styles.each do |k, v|
        deep_copy[k] = v.is_a?(Array) ? v.map(&:dup) : v
      end
      deep_copy
    end

    # Sets a core property.
    def set_core_property(name, value)
      raise ArgumentError, "name must be a Symbol" unless name.is_a?(Symbol)

      @core_properties[name] = value
    end

    # Returns core properties hash.
    def core_properties
      @core_properties.dup
    end

    # Sets an app property.
    def set_app_property(name, value)
      raise ArgumentError, "name must be a Symbol" unless name.is_a?(Symbol)

      @app_properties[name] = value
    end

    # Returns app properties hash.
    def app_properties
      @app_properties.dup
    end

    # Adds a custom document property. type: :string (default), :number, :bool, :date.
    def add_custom_property(name, value, type: :string)
      @custom_properties << { name: name, value: value, type: type }
    end

    # Returns custom properties array.
    def custom_properties
      @custom_properties.map(&:dup)
    end

    # Sets a sheet-level property (e.g. :tab_color, :summary_below, :summary_right).
    def set_sheet_property(name, value, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @sheet_properties.key?(sheet_name)

      @sheet_properties[sheet_name][name] = value
    end

    # Returns sheet properties for the first (or given) sheet.
    def sheet_properties(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      (@sheet_properties[sheet_name] || {}).dup
    end

    # Sets phonetic properties for a sheet (e.g. :font_id, :type, :alignment).
    def set_phonetic_properties(props, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @phonetic_properties.key?(sheet_name)

      @phonetic_properties[sheet_name] = props
    end

    # Returns phonetic properties for the first (or given) sheet.
    def phonetic_properties(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @phonetic_properties[sheet_name]&.dup
    end

    # Sets a sheet format property (e.g. :default_row_height, :default_col_width, :base_col_width).
    def set_sheet_format(name, value, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @sheet_formats.key?(sheet_name)

      @sheet_formats[sheet_name][name] = value
    end

    # Returns sheet format properties for the first (or given) sheet.
    def sheet_format(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      (@sheet_formats[sheet_name] || {}).dup
    end

    # Sets a sheet view property (e.g. :show_grid_lines, :show_row_col_headers, :right_to_left, :zoom_scale).
    def set_sheet_view(name, value, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @sheet_views.key?(sheet_name)

      @sheet_views[sheet_name][name] = value
    end

    # Returns sheet view properties for the first (or given) sheet.
    def sheet_view(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      (@sheet_views[sheet_name] || {}).dup
    end

    # Sets a freeze pane. row: rows to freeze from top, col: columns to freeze from left.
    def set_freeze_pane(row: 0, col: 0, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @freeze_panes.key?(sheet_name)

      @freeze_panes[sheet_name] = { row: row, col: col, state: :frozen }
    end

    # Sets a split pane (non-frozen). x_split/y_split are in 1/20th of a point (twips).
    # top_left_cell: the cell at top-left of the bottom-right pane.
    def set_split_pane(x_split: 0, y_split: 0, top_left_cell: nil, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @freeze_panes.key?(sheet_name)

      @freeze_panes[sheet_name] = { x_split: x_split, y_split: y_split, top_left_cell: top_left_cell, state: :split }
    end

    # Returns freeze pane settings for the first (or given) sheet.
    def freeze_pane(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @freeze_panes[sheet_name]
    end

    # Sets the active cell selection.
    def set_selection(active_cell, sqref: nil, pane: nil, active_cell_id: nil, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @selections.key?(sheet_name)

      sel = { active_cell: active_cell, sqref: sqref || active_cell }
      sel[:pane] = pane if pane
      sel[:active_cell_id] = active_cell_id if active_cell_id
      @selections[sheet_name] = sel
    end

    # Returns selection for the first (or given) sheet.
    def selection(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @selections[sheet_name]
    end

    # Sets a print option (e.g. :grid_lines, :headings, :horizontal_centered, :vertical_centered).
    def set_print_option(name, value, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @print_options.key?(sheet_name)

      @print_options[sheet_name][name] = value
    end

    # Returns print options for the first (or given) sheet.
    def print_options(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      (@print_options[sheet_name] || {}).dup
    end

    # Sets page margins (all values in inches).
    def set_page_margins(left: nil, right: nil, top: nil, bottom: nil, header: nil, footer: nil, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @page_margins.key?(sheet_name)

      m = {}
      m[:left] = left if left
      m[:right] = right if right
      m[:top] = top if top
      m[:bottom] = bottom if bottom
      m[:header] = header if header
      m[:footer] = footer if footer
      @page_margins[sheet_name] = m
    end

    # Returns page margins for the first (or given) sheet.
    def page_margins(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @page_margins[sheet_name]
    end

    # Sets a page setup property (e.g. :orientation, :paper_size, :scale, :fit_to_width, :fit_to_height).
    def set_page_setup(name, value, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @page_setup.key?(sheet_name)

      @page_setup[sheet_name][name] = value
    end

    # Returns page setup for the first (or given) sheet.
    def page_setup(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      (@page_setup[sheet_name] || {}).dup
    end

    # Sets header/footer text (:odd_header, :odd_footer, :even_header, :even_footer).
    def set_header_footer(name, value, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @header_footer.key?(sheet_name)

      @header_footer[sheet_name][name] = value
    end

    # Returns header/footer for the first (or given) sheet.
    def header_footer(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      (@header_footer[sheet_name] || {}).dup
    end

    # Adds a row break (page break before a given row number).
    def add_row_break(row_num, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @row_breaks.key?(sheet_name)

      @row_breaks[sheet_name] << row_num
    end

    # Returns row breaks for the first (or given) sheet.
    def row_breaks(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @row_breaks[sheet_name] || []
    end

    # Adds a column break (page break before a given column index, 1-based).
    def add_col_break(col_index, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @col_breaks.key?(sheet_name)

      @col_breaks[sheet_name] << col_index
    end

    # Returns column breaks for the first (or given) sheet.
    def col_breaks(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @col_breaks[sheet_name] || []
    end

    # Adds a data validation rule.
    # sqref: cell range (e.g. "A1:A100")
    # Options: type:, operator:, formula1:, formula2:, allow_blank:, show_input_message:,
    #          show_error_message:, error_style:, error_title:, error:, prompt_title:, prompt:
    def add_data_validation(sqref, sheet: nil, **opts)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @data_validations.key?(sheet_name)

      @data_validations[sheet_name] << opts.merge(sqref: sqref)
    end

    # Returns data validations for the first (or given) sheet.
    def data_validations(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @data_validations[sheet_name] || []
    end

    # Sets a data validations container option (e.g. disable_prompts, x_window, y_window).
    def set_data_validations_option(name, value, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @data_validations_options.key?(sheet_name)

      @data_validations_options[sheet_name][name] = value
    end

    # Adds a conditional formatting rule to the specified range.
    # Options: type (:cell_is, :expression, :color_scale, :data_bar, :icon_set),
    # operator, priority, formula/formulas, format_id, color_scale, data_bar, icon_set.
    def add_conditional_format(sqref, sheet: nil, **opts)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @conditional_formats.key?(sheet_name)

      @conditional_formats[sheet_name] << opts.merge(sqref: sqref)
    end

    # Returns conditional formatting rules for the first (or given) sheet.
    def conditional_formats(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @conditional_formats[sheet_name] || []
    end

    # Adds a table definition to a sheet.
    # columns: array of column name strings.
    def add_table(ref, columns:, name: nil, display_name: nil, sheet: nil, totals_row_count: 0, style: nil, **opts)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @tables.key?(sheet_name)

      table_id = @tables.values.flatten.size + 1
      tbl_name = name || "Table#{table_id}"
      tbl = {
        id: table_id, ref: ref, name: tbl_name,
        display_name: display_name || tbl_name, columns: columns,
        totals_row_count: totals_row_count
      }
      tbl[:style] = style if style
      %i[header_row_count published comment insert_row insert_row_shift
         header_row_dxf_id data_dxf_id totals_row_dxf_id
         header_row_border_dxf_id table_border_dxf_id totals_row_border_dxf_id
         header_row_cell_style totals_row_cell_style connection_id table_type].each do |key|
        tbl[key] = opts[key] if opts.key?(key)
      end
      @tables[sheet_name] << tbl
    end

    # Returns table definitions for the first (or given) sheet.
    def tables(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @tables[sheet_name] || []
    end

    # Enables shared string table mode (strings stored in sharedStrings.xml).
    def use_shared_strings!
      @use_shared_strings = true
    end

    # Sets a sheet's visibility state (:visible, :hidden, :very_hidden).
    def set_sheet_state(sheet_name, state)
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @sheets.key?(sheet_name)
      raise ArgumentError, "state must be :visible, :hidden, or :very_hidden" unless %i[visible hidden very_hidden].include?(state)

      @sheet_states[sheet_name] = state
    end

    # Returns the sheet state for a given sheet.
    def sheet_state(sheet_name)
      @sheet_states[sheet_name] || :visible
    end

    # Adds a defined name. Options: sheet: (local scope), hidden: true, value: (formula or constant).
    def add_defined_name(name, value, sheet: nil, hidden: false, **opts)
      entry = { name: name, value: value, hidden: hidden }
      %i[comment description function vb_procedure xlm shortcut_key publish_to_server workbook_parameter
         function_group_id custom_menu help status_bar].each do |key|
        entry[key] = opts[key] if opts.key?(key)
      end
      if sheet
        idx = @sheet_order.index(sheet)
        raise ArgumentError, "unknown sheet: #{sheet}" unless idx

        entry[:local_sheet_id] = idx
      end
      @defined_names << entry
    end

    # Returns defined names array.
    def defined_names
      @defined_names.map(&:dup)
    end

    # Sets the print area for a sheet. range should be like "A1:D20".
    # Generates the _xlnm.Print_Area defined name automatically.
    def set_print_area(range, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @sheets.key?(sheet_name)

      value = "'#{sheet_name}'!#{absolute_range(range)}"
      # Remove any existing print area for this sheet
      idx = @sheet_order.index(sheet_name)
      @defined_names.reject! { |dn| dn[:name] == "_xlnm.Print_Area" && dn[:local_sheet_id] == idx }
      add_defined_name("_xlnm.Print_Area", value, sheet: sheet_name)
    end

    # Sets print titles (rows and/or columns to repeat on each page).
    # rows: "1:3" repeats rows 1-3, cols: "A:B" repeats columns A-B.
    def set_print_titles(rows: nil, cols: nil, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @sheets.key?(sheet_name)
      raise ArgumentError, "at least one of rows: or cols: must be specified" unless rows || cols

      parts = []
      parts << "'#{sheet_name}'!$#{cols.sub(":", ":$")}" if cols
      parts << "'#{sheet_name}'!$#{rows.sub(":", ":$")}" if rows
      value = parts.join(",")

      idx = @sheet_order.index(sheet_name)
      @defined_names.reject! { |dn| dn[:name] == "_xlnm.Print_Titles" && dn[:local_sheet_id] == idx }
      add_defined_name("_xlnm.Print_Titles", value, sheet: sheet_name)
    end

    # Sets a workbook property (e.g. :date1904, :default_theme_version).
    def set_workbook_property(name, value)
      @workbook_properties[name] = value
    end

    # Returns workbook properties hash.
    def workbook_properties
      @workbook_properties.dup
    end

    # Sets a file version property (e.g. :app_name, :last_edited, :lowest_edited, :rup_build, :code_name).
    def set_file_version(name, value)
      @file_version[name] = value
    end

    # Returns file version hash.
    def file_version
      @file_version.dup
    end

    # Sets a file sharing property (e.g. :read_only_recommended, :user_name).
    def set_file_sharing(name, value)
      @file_sharing[name] = value
    end

    # Returns file sharing hash.
    def file_sharing
      @file_sharing.dup
    end

    # Sets a workbook view property (e.g. :active_tab, :first_sheet).
    def set_workbook_view(name, value)
      @workbook_views[name] = value
    end

    # Returns workbook view properties hash.
    def workbook_views
      @workbook_views.dup
    end

    # Sets a calc property (e.g. :calc_id, :full_calc_on_load).
    def set_calc_property(name, value)
      @calc_properties[name] = value
    end

    # Returns calc properties hash.
    def calc_properties
      @calc_properties.dup
    end

    # Sets a file recovery property (e.g. :auto_recover, :crash_save).
    def set_file_recovery_property(name, value)
      @file_recovery_properties[name] = value
    end

    # Returns file recovery properties hash.
    def file_recovery_properties
      @file_recovery_properties.dup
    end

    # Sets sheet protection options.
    # Options: :password, :sheet, :objects, :scenarios, :format_cells, :format_columns,
    #   :format_rows, :insert_columns, :insert_rows, :insert_hyperlinks,
    #   :delete_columns, :delete_rows, :select_locked_cells, :sort, :auto_filter,
    #   :pivot_tables, :select_unlocked_cells, :algorithm_name, :hash_value, :salt_value, :spin_count
    def set_sheet_protection(sheet: nil, **opts)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @sheets.key?(sheet_name)

      @sheet_protection[sheet_name] = opts
    end

    # Returns sheet protection settings for the given sheet.
    def sheet_protection(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @sheet_protection[sheet_name]&.dup
    end

    # Adds a protected range to the given sheet.
    # Required: name:, sqref:
    # Optional: algorithm_name:, hash_value:, salt_value:, spin_count:, security_descriptors: []
    def add_protected_range(sheet: nil, **opts)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @sheets.key?(sheet_name)
      raise ArgumentError, "name is required" unless opts[:name]
      raise ArgumentError, "sqref is required" unless opts[:sqref]

      @protected_ranges[sheet_name] << opts
    end

    # Returns protected ranges for the given sheet.
    def protected_ranges(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      (@protected_ranges[sheet_name] || []).map(&:dup)
    end

    # Adds a cell watch to the given sheet.
    def add_cell_watch(cell_ref, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @sheets.key?(sheet_name)

      @cell_watches[sheet_name] << cell_ref
    end

    # Returns cell watches for the given sheet.
    def cell_watches(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      (@cell_watches[sheet_name] || []).dup
    end

    # Adds an ignored error entry for the given sheet.
    # Options: sqref:, eval_error:, two_digit_text_year:, number_stored_as_text:, formula:,
    #   formula_range:, unlocked_formula:, empty_cell_reference:, list_data_validation:, calculated_column:
    def add_ignored_error(sheet: nil, **opts)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @sheets.key?(sheet_name)
      raise ArgumentError, "sqref is required" unless opts[:sqref]

      @ignored_errors[sheet_name] << opts
    end

    # Returns ignored errors for the given sheet.
    def ignored_errors(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      (@ignored_errors[sheet_name] || []).dup
    end

    # Sets data consolidation options for the given sheet.
    # Options: function:, start_labels:, top_labels:, link:, data_refs: [{ref:, name:, sheet:}]
    def set_data_consolidate(sheet: nil, **opts)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @sheets.key?(sheet_name)

      @data_consolidate[sheet_name] = opts
    end

    # Returns data consolidation settings for the given sheet.
    def data_consolidate(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @data_consolidate[sheet_name]&.dup
    end

    # Sets scenarios for the given sheet.
    # Options: current:, show:, sqref:, scenarios: [{name:, input_cells: [{r:, val:, ...}], ...}]
    def set_scenarios(sheet: nil, **opts)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @sheets.key?(sheet_name)

      @scenarios[sheet_name] = opts
    end

    # Returns scenarios for the given sheet.
    def scenarios(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @scenarios[sheet_name]&.dup
    end

    # Sets workbook protection options.
    # Options: :lock_structure, :lock_windows, :password, :algorithm_name, :hash_value, :salt_value, :spin_count
    def set_workbook_protection(**opts)
      @workbook_protection = opts
    end

    # Returns workbook protection settings.
    def workbook_protection
      @workbook_protection&.dup
    end

    # Returns ordered sheet names.
    attr_reader :sheet_order

    # Adds a raw ZIP entry to be included in the output (for pass-through retention).
    def add_raw_entry(path, content, content_type: nil)
      @extra_entries[path] = content
      return unless content_type

      ext = File.extname(path).delete(".")
      if ext.empty? || path.include?("/")
        @extra_ct_overrides["/#{path}"] = content_type
      else
        @extra_ct_defaults[ext] = content_type
      end
    end

    # Copies all ZIP entries from an existing XLSX file as pass-through.
    # Generated parts override pass-through parts with the same path.
    def copy_entries_from(filepath)
      reader = Xlsxrb::Reader.new(filepath)
      reader.entry_names.each do |name|
        @extra_entries[name] = reader.raw_entry(name)
      end

      # Parse [Content_Types].xml for extra content types.
      ct_xml = reader.raw_entry("[Content_Types].xml")
      parse_extra_content_types(ct_xml) if ct_xml && !ct_xml.empty?
    end

    # Inserts an image from file data into the given sheet.
    # file_data: raw image bytes. ext: file extension (e.g. "png").
    # from_col/from_row: anchor start. to_col/to_row: anchor end.
    def insert_image(file_data, ext: "png", from_col: 0, from_row: 0, to_col: 5, to_row: 10, from_col_off: nil, from_row_off: nil, to_col_off: nil, to_row_off: nil, name: nil, description: nil, title: nil, hidden: nil, macro: nil, no_change_aspect: true, no_crop: nil, line_color: nil, line_width: nil, rotation: nil, edit_as: nil, published: nil, locks_with_sheet: nil, prints_with_sheet: nil, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @images.key?(sheet_name)

      img_name = name || "Picture #{@images[sheet_name].size + 1}"
      img = {
        file_data: file_data, ext: ext, name: img_name,
        from_col: from_col, from_row: from_row,
        to_col: to_col, to_row: to_row
      }
      img[:from_col_off] = from_col_off if from_col_off
      img[:from_row_off] = from_row_off if from_row_off
      img[:to_col_off] = to_col_off if to_col_off
      img[:to_row_off] = to_row_off if to_row_off
      img[:description] = description if description
      img[:title] = title if title
      img[:hidden] = hidden unless hidden.nil?
      img[:macro] = macro if macro
      img[:no_change_aspect] = no_change_aspect
      img[:no_crop] = no_crop unless no_crop.nil?
      img[:line_color] = line_color if line_color
      img[:line_width] = line_width if line_width
      img[:rotation] = rotation if rotation
      img[:edit_as] = edit_as if edit_as
      img[:published] = published unless published.nil?
      img[:locks_with_sheet] = locks_with_sheet unless locks_with_sheet.nil?
      img[:prints_with_sheet] = prints_with_sheet unless prints_with_sheet.nil?
      @images[sheet_name] << img
    end

    # Returns images for the first (or given) sheet.
    def images(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @images[sheet_name] || []
    end

    # Adds a chart to the given sheet.
    # type: :bar, :line, :pie. title: chart title string.
    # data_ref: e.g. "Sheet1!$A$1:$B$4". cat_ref/val_ref for explicit series.
    def add_chart(type: :bar, title: nil, auto_title_deleted: nil, cat_ref: nil, val_ref: nil, series: nil, legend: nil, data_labels: nil, cat_axis_title: nil, val_axis_title: nil, cat_axis_tick_lbl_pos: nil, val_axis_tick_lbl_pos: nil, cat_axis_major_gridlines: nil, val_axis_major_gridlines: nil, cat_axis_minor_gridlines: nil, val_axis_minor_gridlines: nil, cat_axis_delete: nil, val_axis_delete: nil, cat_axis_orientation: nil, val_axis_orientation: nil, cat_axis_num_fmt: nil, val_axis_num_fmt: nil, cat_axis_major_tick_mark: nil, cat_axis_minor_tick_mark: nil, val_axis_major_tick_mark: nil, val_axis_minor_tick_mark: nil, cat_axis_crosses: nil, val_axis_crosses: nil, cat_axis_crosses_at: nil, val_axis_crosses_at: nil, cat_axis_tick_lbl_skip: nil, cat_axis_tick_mark_skip: nil, cat_axis_lbl_offset: nil, cat_axis_no_multi_lvl_lbl: nil, val_axis_cross_between: nil, val_axis_major_unit: nil, val_axis_minor_unit: nil, cat_axis_scaling_max: nil, cat_axis_scaling_min: nil, val_axis_scaling_max: nil, val_axis_scaling_min: nil, cat_axis_log_base: nil, val_axis_log_base: nil, val_axis_disp_units: nil, gap_width: nil, gap_depth: nil, overlap: nil, first_slice_ang: nil, hole_size: nil, smooth: nil, marker: nil, scatter_style: nil, radar_style: nil, bar_shape: nil, bubble_3d: nil, bubble_scale: nil, show_neg_bubbles: nil, size_represents: nil, wireframe: nil, grouping: nil, bar_dir: nil, vary_colors: nil, style: nil, rounded_corners: nil, view_3d: nil, cat_axis_pos: nil, val_axis_pos: nil, name: nil, description: nil, frame_title: nil, frame_hidden: nil, frame_macro: nil, frame_no_grp: nil, from_col: 0, from_row: 0, to_col: 10, to_row: 15, from_col_off: nil, from_row_off: nil, to_col_off: nil, to_row_off: nil, edit_as: nil, published: nil, locks_with_sheet: nil, prints_with_sheet: nil, plot_vis_only: nil, disp_blanks_as: nil, show_d_lbls_over_max: nil, data_table: nil, plot_area_fill: nil, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @charts_data.key?(sheet_name)

      chart = { type: type, title: title,
                from_col: from_col, from_row: from_row,
                to_col: to_col, to_row: to_row }
      chart[:from_col_off] = from_col_off if from_col_off
      chart[:from_row_off] = from_row_off if from_row_off
      chart[:to_col_off] = to_col_off if to_col_off
      chart[:to_row_off] = to_row_off if to_row_off
      chart[:series] = (series || [{ cat_ref: cat_ref, val_ref: val_ref }])
      chart[:auto_title_deleted] = auto_title_deleted unless auto_title_deleted.nil?
      chart[:legend] = legend if legend
      chart[:data_labels] = data_labels if data_labels
      chart[:cat_axis_title] = cat_axis_title if cat_axis_title
      chart[:val_axis_title] = val_axis_title if val_axis_title
      chart[:cat_axis_tick_lbl_pos] = cat_axis_tick_lbl_pos if cat_axis_tick_lbl_pos
      chart[:val_axis_tick_lbl_pos] = val_axis_tick_lbl_pos if val_axis_tick_lbl_pos
      chart[:cat_axis_major_gridlines] = cat_axis_major_gridlines unless cat_axis_major_gridlines.nil?
      chart[:val_axis_major_gridlines] = val_axis_major_gridlines unless val_axis_major_gridlines.nil?
      chart[:cat_axis_minor_gridlines] = cat_axis_minor_gridlines unless cat_axis_minor_gridlines.nil?
      chart[:val_axis_minor_gridlines] = val_axis_minor_gridlines unless val_axis_minor_gridlines.nil?
      chart[:cat_axis_delete] = cat_axis_delete unless cat_axis_delete.nil?
      chart[:val_axis_delete] = val_axis_delete unless val_axis_delete.nil?
      chart[:cat_axis_orientation] = cat_axis_orientation if cat_axis_orientation
      chart[:val_axis_orientation] = val_axis_orientation if val_axis_orientation
      chart[:cat_axis_num_fmt] = cat_axis_num_fmt if cat_axis_num_fmt
      chart[:val_axis_num_fmt] = val_axis_num_fmt if val_axis_num_fmt
      chart[:cat_axis_major_tick_mark] = cat_axis_major_tick_mark if cat_axis_major_tick_mark
      chart[:cat_axis_minor_tick_mark] = cat_axis_minor_tick_mark if cat_axis_minor_tick_mark
      chart[:val_axis_major_tick_mark] = val_axis_major_tick_mark if val_axis_major_tick_mark
      chart[:val_axis_minor_tick_mark] = val_axis_minor_tick_mark if val_axis_minor_tick_mark
      chart[:cat_axis_crosses] = cat_axis_crosses if cat_axis_crosses
      chart[:val_axis_crosses] = val_axis_crosses if val_axis_crosses
      chart[:cat_axis_crosses_at] = cat_axis_crosses_at if cat_axis_crosses_at
      chart[:val_axis_crosses_at] = val_axis_crosses_at if val_axis_crosses_at
      chart[:cat_axis_tick_lbl_skip] = cat_axis_tick_lbl_skip if cat_axis_tick_lbl_skip
      chart[:cat_axis_tick_mark_skip] = cat_axis_tick_mark_skip if cat_axis_tick_mark_skip
      chart[:cat_axis_lbl_offset] = cat_axis_lbl_offset if cat_axis_lbl_offset
      chart[:cat_axis_no_multi_lvl_lbl] = cat_axis_no_multi_lvl_lbl unless cat_axis_no_multi_lvl_lbl.nil?
      chart[:val_axis_cross_between] = val_axis_cross_between if val_axis_cross_between
      chart[:val_axis_major_unit] = val_axis_major_unit if val_axis_major_unit
      chart[:val_axis_minor_unit] = val_axis_minor_unit if val_axis_minor_unit
      chart[:cat_axis_scaling_max] = cat_axis_scaling_max if cat_axis_scaling_max
      chart[:cat_axis_scaling_min] = cat_axis_scaling_min if cat_axis_scaling_min
      chart[:val_axis_scaling_max] = val_axis_scaling_max if val_axis_scaling_max
      chart[:val_axis_scaling_min] = val_axis_scaling_min if val_axis_scaling_min
      chart[:cat_axis_log_base] = cat_axis_log_base if cat_axis_log_base
      chart[:val_axis_log_base] = val_axis_log_base if val_axis_log_base
      chart[:val_axis_disp_units] = val_axis_disp_units if val_axis_disp_units
      chart[:first_slice_ang] = first_slice_ang if first_slice_ang
      chart[:hole_size] = hole_size if hole_size
      chart[:smooth] = smooth unless smooth.nil?
      chart[:marker] = marker unless marker.nil?
      chart[:scatter_style] = scatter_style if scatter_style
      chart[:radar_style] = radar_style if radar_style
      chart[:cat_axis_pos] = cat_axis_pos if cat_axis_pos
      chart[:val_axis_pos] = val_axis_pos if val_axis_pos
      chart[:gap_width] = gap_width if gap_width
      chart[:gap_depth] = gap_depth if gap_depth
      chart[:bar_shape] = bar_shape if bar_shape
      chart[:bubble_3d] = bubble_3d unless bubble_3d.nil?
      chart[:bubble_scale] = bubble_scale if bubble_scale
      chart[:show_neg_bubbles] = show_neg_bubbles unless show_neg_bubbles.nil?
      chart[:size_represents] = size_represents if size_represents
      chart[:overlap] = overlap if overlap
      chart[:grouping] = grouping if grouping
      chart[:bar_dir] = bar_dir if bar_dir
      chart[:vary_colors] = vary_colors unless vary_colors.nil?
      chart[:wireframe] = wireframe unless wireframe.nil?
      chart[:style] = style if style
      chart[:rounded_corners] = rounded_corners unless rounded_corners.nil?
      chart[:view_3d] = view_3d if view_3d
      chart[:name] = name if name
      chart[:description] = description if description
      chart[:frame_title] = frame_title if frame_title
      chart[:frame_hidden] = frame_hidden unless frame_hidden.nil?
      chart[:frame_macro] = frame_macro if frame_macro
      chart[:frame_no_grp] = frame_no_grp unless frame_no_grp.nil?
      chart[:edit_as] = edit_as if edit_as
      chart[:published] = published unless published.nil?
      chart[:locks_with_sheet] = locks_with_sheet unless locks_with_sheet.nil?
      chart[:prints_with_sheet] = prints_with_sheet unless prints_with_sheet.nil?
      chart[:plot_vis_only] = plot_vis_only unless plot_vis_only.nil?
      chart[:disp_blanks_as] = disp_blanks_as if disp_blanks_as
      chart[:show_d_lbls_over_max] = show_d_lbls_over_max unless show_d_lbls_over_max.nil?
      chart[:data_table] = data_table if data_table
      chart[:plot_area_fill] = plot_area_fill if plot_area_fill
      @charts_data[sheet_name] << chart
    end

    # Returns chart definitions for the first (or given) sheet.
    def charts(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @charts_data[sheet_name] || []
    end

    # Adds a shape to the given sheet.
    # preset: preset geometry name (e.g. "rect", "ellipse", "roundRect").
    # text: optional text body string.
    # from_col/from_row/to_col/to_row: anchor coordinates.
    def add_shape(preset: "rect", text: nil, name: nil, description: nil, title: nil, hidden: nil, macro: nil, textlink: nil, f_locks_text: nil, no_grp: nil, no_rot: nil, fill_color: nil, no_fill: nil, line_color: nil, line_width: nil, no_line: nil, rotation: nil, text_wrap: nil, text_anchor: nil, text_vert_overflow: nil, from_col: 0, from_row: 0, to_col: 5, to_row: 5, from_col_off: nil, from_row_off: nil, to_col_off: nil, to_row_off: nil, edit_as: nil, published: nil, locks_with_sheet: nil, prints_with_sheet: nil, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @shapes_data.key?(sheet_name)

      shape_name = name || "Shape #{@shapes_data[sheet_name].size + 1}"
      shape = {
        preset: preset, text: text, name: shape_name,
        from_col: from_col, from_row: from_row,
        to_col: to_col, to_row: to_row
      }
      shape[:from_col_off] = from_col_off if from_col_off
      shape[:from_row_off] = from_row_off if from_row_off
      shape[:to_col_off] = to_col_off if to_col_off
      shape[:to_row_off] = to_row_off if to_row_off
      shape[:description] = description if description
      shape[:title] = title if title
      shape[:hidden] = hidden unless hidden.nil?
      shape[:macro] = macro if macro
      shape[:textlink] = textlink if textlink
      shape[:f_locks_text] = f_locks_text unless f_locks_text.nil?
      shape[:no_grp] = no_grp unless no_grp.nil?
      shape[:no_rot] = no_rot unless no_rot.nil?
      shape[:fill_color] = fill_color if fill_color
      shape[:no_fill] = no_fill unless no_fill.nil?
      shape[:line_color] = line_color if line_color
      shape[:line_width] = line_width if line_width
      shape[:no_line] = no_line unless no_line.nil?
      shape[:text_wrap] = text_wrap if text_wrap
      shape[:text_anchor] = text_anchor if text_anchor
      shape[:text_vert_overflow] = text_vert_overflow if text_vert_overflow
      shape[:rotation] = rotation if rotation
      shape[:edit_as] = edit_as if edit_as
      shape[:published] = published unless published.nil?
      shape[:locks_with_sheet] = locks_with_sheet unless locks_with_sheet.nil?
      shape[:prints_with_sheet] = prints_with_sheet unless prints_with_sheet.nil?
      @shapes_data[sheet_name] << shape
    end

    # Returns shape definitions for the first (or given) sheet.
    def shapes(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @shapes_data[sheet_name] || []
    end

    # Adds a comment on a cell.
    def add_comment(cell_address, text, author: "Author", sheet: nil, guid: nil, shape_id: nil)
      validate_cell_address!(cell_address)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @comments_data.key?(sheet_name)

      entry = { ref: cell_address, text: text, author: author }
      entry[:guid] = guid if guid
      entry[:shape_id] = shape_id if shape_id
      @comments_data[sheet_name] << entry
    end

    # Returns comment definitions for the first (or given) sheet.
    def comments(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @comments_data[sheet_name] || []
    end

    # Adds a pivot table to the given sheet.
    # source_ref: data source range (e.g. "Sheet1!A1:C4").
    # row_fields: array of 0-based field indices for row axis.
    # data_fields: array of { fld:, name:, subtotal: } hashes.
    # col_fields: array of 0-based field indices for column axis.
    # field_names: array of field name strings (for cache definition).
    # items: hash mapping field index to array of item values.
    def add_pivot_table(source_ref, row_fields:, data_fields:, col_fields: [], dest_ref: "E1", name: nil, field_names: nil, items: nil, sheet: nil, **opts)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @pivot_tables_data.key?(sheet_name)

      pt_name = name || "PivotTable#{@pivot_tables_data.values.flatten.size + 1}"
      @pivot_tables_data[sheet_name] << {
        name: pt_name, source_ref: source_ref,
        row_fields: row_fields, col_fields: col_fields,
        data_fields: data_fields, dest_ref: dest_ref,
        field_names: field_names, items: items,
        data_caption: opts[:data_caption], data_on_rows: opts[:data_on_rows],
        row_grand_totals: opts[:row_grand_totals], col_grand_totals: opts[:col_grand_totals],
        compact: opts[:compact], outline: opts[:outline], show_headers: opts[:show_headers],
        source_name: opts[:source_name],
        grand_total_caption: opts[:grand_total_caption], error_caption: opts[:error_caption],
        show_error: opts[:show_error], missing_caption: opts[:missing_caption],
        show_missing: opts[:show_missing], tag: opts[:tag], indent: opts[:indent],
        published: opts[:published], created_version: opts[:created_version],
        updated_version: opts[:updated_version], min_refreshable_version: opts[:min_refreshable_version],
        pivot_table_style: opts[:pivot_table_style],
        cache_save_data: opts[:cache_save_data], cache_enable_refresh: opts[:cache_enable_refresh],
        cache_refreshed_by: opts[:cache_refreshed_by], cache_refreshed_version: opts[:cache_refreshed_version],
        cache_created_version: opts[:cache_created_version], cache_record_count: opts[:cache_record_count],
        cache_optimize_memory: opts[:cache_optimize_memory],
        row_page_count: opts[:row_page_count], col_page_count: opts[:col_page_count],
        field_attrs: opts[:field_attrs],
        apply_number_formats: opts[:apply_number_formats],
        apply_border_formats: opts[:apply_border_formats],
        apply_font_formats: opts[:apply_font_formats],
        apply_pattern_formats: opts[:apply_pattern_formats],
        apply_alignment_formats: opts[:apply_alignment_formats],
        apply_width_height_formats: opts[:apply_width_height_formats],
        multiple_field_filters: opts[:multiple_field_filters],
        show_drill: opts[:show_drill],
        show_data_tips: opts[:show_data_tips],
        enable_drill: opts[:enable_drill],
        show_member_property_tips: opts[:show_member_property_tips],
        item_print_titles: opts[:item_print_titles],
        field_print_titles: opts[:field_print_titles],
        preserve_formatting: opts[:preserve_formatting],
        page_over_then_down: opts[:page_over_then_down],
        page_wrap: opts[:page_wrap],
        compact_data: opts[:compact_data],
        outline_data: opts[:outline_data],
        show_multiple_label: opts[:show_multiple_label],
        show_data_drop_down: opts[:show_data_drop_down],
        edit_data: opts[:edit_data],
        disable_field_list: opts[:disable_field_list],
        visual_totals: opts[:visual_totals],
        print_drill: opts[:print_drill]
      }
    end

    # Returns pivot table definitions for the first (or given) sheet.
    def pivot_tables(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @pivot_tables_data[sheet_name] || []
    end

    # Adds an external link reference to another workbook.
    # target: path or URI to the external workbook (e.g. "Book2.xlsx").
    # sheet_names: array of sheet name strings in the external workbook.
    def add_external_link(target:, sheet_names: [])
      @external_links << { target: target, sheet_names: sheet_names }
    end

    # Returns external link definitions.
    def external_links
      @external_links.dup
    end

    # Enables macro preservation mode. Required when copy_entries_from loads a .xlsm file.
    def preserve_macros!
      @preserve_macros = true
    end

    # Returns whether macro preservation is enabled.
    def preserve_macros?
      @preserve_macros
    end

    # Writes the workbook as an XLSX file to the given path.
    def write(filepath)
      # Pre-register date/datetime formats if any sheet contains Date or Time values.
      needs_date = false
      needs_datetime = false
      @sheet_order.each do |sn|
        @sheets[sn].each_value do |v|
          needs_datetime = true if v.is_a?(Time)
          needs_date = true if v.is_a?(Date) && !v.is_a?(Time)
        end
        break if needs_date && needs_datetime
      end
      date_num_fmt_id if needs_date
      datetime_num_fmt_id if needs_datetime

      # Clear memoized xf index map so it picks up all registered formats.
      @xf_index_map = nil
      # Pre-populate xf entries for legacy num_fmt-based styles.
      @num_fmts.each { |nf| resolve_style_index(nf[:num_fmt_id]) }

      # Build shared string table if enabled.
      sst = build_shared_string_table if @use_shared_strings

      # Track generated drawing/image/chart/comment/pivot indices.
      @drawing_count = 0
      @chart_count = 0
      @comment_count = 0
      @pivot_cache_count = 0
      @pivot_table_count = 0
      @media_count = 0

      # Pre-compute total pivot cache count for workbook XML.
      total_pivot_caches = @pivot_tables_data.values.sum(&:size)

      entries = {
        "_rels/.rels" => generate_rels_root,
        "xl/workbook.xml" => generate_workbook_xml(total_pivot_caches),
        "xl/styles.xml" => generate_styles_xml
      }

      entries["docProps/core.xml"] = generate_core_properties_xml unless @core_properties.empty?
      entries["docProps/app.xml"] = generate_app_properties_xml unless @app_properties.empty?
      entries["docProps/custom.xml"] = generate_custom_properties_xml unless @custom_properties.empty?
      entries["xl/sharedStrings.xml"] = generate_shared_strings_xml(sst) if sst

      table_index = 0
      @sheet_order.each_with_index do |sheet_name, i|
        sheet_images = @images[sheet_name] || []
        sheet_charts = @charts_data[sheet_name] || []
        sheet_shapes = @shapes_data[sheet_name] || []
        sheet_comments = @comments_data[sheet_name] || []
        sheet_pivots = @pivot_tables_data[sheet_name] || []
        has_drawing = sheet_images.any? || sheet_charts.any? || sheet_shapes.any?
        has_comments = sheet_comments.any?

        # Pre-increment counters so rels reference correct paths.
        sheet_drawing_idx = has_drawing ? (@drawing_count += 1) : nil
        sheet_comment_idx = has_comments ? (@comment_count += 1) : nil
        sheet_pivot_start = @pivot_table_count
        sheet_cache_start = @pivot_cache_count
        sheet_pivots.each do
          @pivot_cache_count += 1
          @pivot_table_count += 1
        end

        # Build per-sheet rels first (needed for rId calculation in worksheet XML).
        sheet_tables = @tables[sheet_name] || []
        sheet_rels_parts = build_sheet_rels_parts_v2(
          sheet_name, sheet_tables, table_index,
          sheet_drawing_idx, sheet_comment_idx,
          sheet_pivot_start, sheet_pivots.size
        )

        # Calculate the legacyDrawing rId if comments exist.
        # The VML rel is always the one after the comments rel in rels.
        vml_rid = nil
        if has_comments
          vml_idx = sheet_rels_parts.index { |r| r[:type]&.end_with?("/vmlDrawing") }
          vml_rid = vml_idx + 1 if vml_idx # 1-based rId
        end

        entries["xl/worksheets/sheet#{i + 1}.xml"] = generate_worksheet_xml(
          @sheets[sheet_name], @column_widths[sheet_name], @column_attrs[sheet_name], @row_attrs[sheet_name],
          @auto_filters[sheet_name], @filter_columns[sheet_name], @sort_state[sheet_name],
          @merge_cells[sheet_name], @hyperlinks[sheet_name],
          @cell_styles[sheet_name], @sheet_properties[sheet_name], @sheet_formats[sheet_name],
          @sheet_views[sheet_name], @freeze_panes[sheet_name], @selections[sheet_name],
          @print_options[sheet_name], @page_margins[sheet_name], @page_setup[sheet_name],
          @header_footer[sheet_name], @row_breaks[sheet_name], @col_breaks[sheet_name],
          @data_validations[sheet_name], @conditional_formats[sheet_name], sst,
          @tables[sheet_name] || [], @hyperlinks[sheet_name].size,
          has_drawing:, has_comments:,
          sheet_prot: @sheet_protection[sheet_name], vml_rid:,
          phonetic_pr: @phonetic_properties[sheet_name],
          dv_options: @data_validations_options[sheet_name],
          prot_ranges: @protected_ranges[sheet_name],
          cell_watches: @cell_watches[sheet_name],
          ignored_errors: @ignored_errors[sheet_name],
          data_consol: @data_consolidate[sheet_name],
          sheet_scenarios: @scenarios[sheet_name],
          cell_phonetic: @cell_phonetic[sheet_name]
        )

        # Generate drawing XML + media + chart entries.
        if has_drawing
          drawing_rels_data = []
          drawing_parts = []

          sheet_images.each do |img|
            @media_count += 1
            media_path = "xl/media/image#{@media_count}.#{img[:ext]}"
            entries[media_path] = img[:file_data]
            drawing_rels_data << { type: :image, target: "../media/image#{@media_count}.#{img[:ext]}" }
            drawing_parts << { kind: :pic, img: img, rid_index: drawing_rels_data.size }
          end

          sheet_charts.each do |chart|
            @chart_count += 1
            chart_path = "xl/charts/chart#{@chart_count}.xml"
            entries[chart_path] = generate_chart_xml(chart)
            drawing_rels_data << { type: :chart, target: "../charts/chart#{@chart_count}.xml" }
            drawing_parts << { kind: :chart, chart: chart, rid_index: drawing_rels_data.size }
          end

          shape_id_base = drawing_parts.size + 1
          sheet_shapes.each_with_index do |shape, si|
            drawing_parts << { kind: :sp, shape: shape, id: shape_id_base + si + 1 }
          end

          entries["xl/drawings/drawing#{sheet_drawing_idx}.xml"] = generate_drawing_xml(drawing_parts)
          entries["xl/drawings/_rels/drawing#{sheet_drawing_idx}.xml.rels"] = generate_drawing_rels(drawing_rels_data) unless drawing_rels_data.empty?
        end

        # Generate comments XML and VML drawing.
        if has_comments
          entries["xl/comments#{sheet_comment_idx}.xml"] = generate_comments_xml(sheet_comments)
          entries["xl/drawings/vmlDrawing#{sheet_comment_idx}.vml"] = generate_vml_drawing_xml(sheet_comments)
        end

        # Generate pivot table + cache entries.
        sheet_pivots.each_with_index do |pt, pi|
          cache_idx = sheet_cache_start + pi + 1
          pt_idx = sheet_pivot_start + pi + 1
          entries["xl/pivotCache/pivotCacheDefinition#{cache_idx}.xml"] = generate_pivot_cache_definition_xml(pt, cache_idx)
          entries["xl/pivotCache/pivotCacheRecords#{cache_idx}.xml"] = generate_pivot_cache_records_xml(pt)
          entries["xl/pivotTables/pivotTable#{pt_idx}.xml"] = generate_pivot_table_xml(pt, cache_idx)
          entries["xl/pivotCache/_rels/pivotCacheDefinition#{cache_idx}.xml.rels"] = generate_pivot_cache_rels(cache_idx)
          entries["xl/pivotTables/_rels/pivotTable#{pt_idx}.xml.rels"] = generate_pivot_table_rels(cache_idx)
        end

        # Emit worksheet rels if any relationships exist.
        entries["xl/worksheets/_rels/sheet#{i + 1}.xml.rels"] = generate_generic_rels(sheet_rels_parts) unless sheet_rels_parts.empty?

        sheet_tables.each do |tbl|
          table_index += 1
          entries["xl/tables/table#{table_index}.xml"] = generate_table_xml(tbl)
        end
      end

      # Generate calcChain.xml if any cells have formulas.
      calc_chain_xml = generate_calc_chain_xml
      entries["xl/calcChain.xml"] = calc_chain_xml if calc_chain_xml

      # Generate external link entries.
      @external_links.each_with_index do |el, idx|
        link_num = idx + 1
        entries["xl/externalLinks/externalLink#{link_num}.xml"] = generate_external_link_xml(el)
        entries["xl/externalLinks/_rels/externalLink#{link_num}.xml.rels"] = generate_external_link_rels(el)
      end

      # Generate workbook rels (needs to know pivot cache count).
      entries["xl/_rels/workbook.xml.rels"] = generate_workbook_rels(has_calc_chain: entries.key?("xl/calcChain.xml"))

      # Content types must be generated after all entries are known.
      entries["[Content_Types].xml"] = generate_content_types_xml(entries)

      # Merge extra (pass-through) entries — generated entries take priority.
      merged = @extra_entries.merge(entries)

      generator = ZipGenerator.new(filepath)
      merged.each { |path, content| generator.add_entry(path, content) }
      generator.generate
    end

    CF_TYPE_MAP = {
      cell_is: "cellIs",
      expression: "expression",
      color_scale: "colorScale",
      data_bar: "dataBar",
      icon_set: "iconSet",
      above_average: "aboveAverage",
      top10: "top10",
      duplicate_values: "duplicateValues",
      unique_values: "uniqueValues",
      contains_text: "containsText",
      not_contains_text: "notContainsText",
      begins_with: "beginsWith",
      ends_with: "endsWith",
      contains_blanks: "containsBlanks",
      not_contains_blanks: "notContainsBlanks",
      time_period: "timePeriod"
    }.freeze

    XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
    C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    CHART_TYPE_MAP = {
      bar: "barChart", line: "lineChart", pie: "pieChart",
      area: "areaChart", scatter: "scatterChart", doughnut: "doughnutChart",
      radar: "radarChart", bar3d: "bar3DChart", line3d: "line3DChart",
      pie3d: "pie3DChart", area3d: "area3DChart", surface: "surfaceChart",
      stock: "stockChart", bubble: "bubbleChart"
    }.freeze

    NO_AXIS_CHARTS = %w[pieChart doughnutChart pie3DChart].freeze
    GROUPING_CHARTS = %w[barChart lineChart areaChart bar3DChart line3DChart area3DChart].freeze

    private

    def generate_content_types_xml(all_entries = {})
      defaults = {
        "rels" => "application/vnd.openxmlformats-package.relationships+xml",
        "xml" => "application/xml"
      }

      # Add image extension defaults.
      image_exts = {}
      @images.each_value do |imgs|
        imgs.each do |img|
          ext = img[:ext]
          mime = case ext
                 when "png" then "image/png"
                 when "jpg", "jpeg" then "image/jpeg"
                 when "gif" then "image/gif"
                 when "bmp" then "image/bmp"
                 else "image/#{ext}"
                 end
          image_exts[ext] = mime
        end
      end
      defaults.merge!(image_exts)

      # Add vml extension if comments exist.
      defaults["vml"] = "application/vnd.openxmlformats-officedocument.vmlDrawing" if @comment_count.to_i.positive?

      # Merge extra defaults from pass-through.
      defaults.merge!(@extra_ct_defaults)

      # Add bin extension if macros are preserved.
      defaults["bin"] = "application/vnd.ms-office.vbaProject" if @preserve_macros

      workbook_ct = if @preserve_macros
                      "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
                    else
                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
                    end

      overrides = {}
      overrides["/xl/workbook.xml"] = workbook_ct
      overrides["/xl/styles.xml"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"

      @sheet_order.each_with_index do |_, i|
        overrides["/xl/worksheets/sheet#{i + 1}.xml"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
      end
      overrides["/xl/sharedStrings.xml"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" if @use_shared_strings

      table_idx = 0
      @sheet_order.each do |sn|
        (@tables[sn] || []).each do
          table_idx += 1
          overrides["/xl/tables/table#{table_idx}.xml"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"
        end
      end

      # Drawings.
      (1..@drawing_count.to_i).each do |d|
        overrides["/xl/drawings/drawing#{d}.xml"] = "application/vnd.openxmlformats-officedocument.drawing+xml"
      end

      # Charts.
      (1..@chart_count.to_i).each do |c|
        overrides["/xl/charts/chart#{c}.xml"] = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
      end

      # Comments.
      (1..@comment_count.to_i).each do |c|
        overrides["/xl/comments#{c}.xml"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
      end

      # Pivot tables and cache.
      (1..@pivot_table_count.to_i).each do |p|
        overrides["/xl/pivotTables/pivotTable#{p}.xml"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"
      end
      (1..@pivot_cache_count.to_i).each do |p|
        overrides["/xl/pivotCache/pivotCacheDefinition#{p}.xml"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"
        overrides["/xl/pivotCache/pivotCacheRecords#{p}.xml"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"
      end

      # External links.
      @external_links.each_with_index do |_, idx|
        overrides["/xl/externalLinks/externalLink#{idx + 1}.xml"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"
      end

      # calcChain.
      overrides["/xl/calcChain.xml"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml" if all_entries.key?("xl/calcChain.xml")

      overrides["/docProps/core.xml"] = "application/vnd.openxmlformats-package.core-properties+xml" unless @core_properties.empty?
      overrides["/docProps/app.xml"] = "application/vnd.openxmlformats-officedocument.extended-properties+xml" unless @app_properties.empty?
      overrides["/docProps/custom.xml"] = "application/vnd.openxmlformats-officedocument.custom-properties+xml" unless @custom_properties.empty?

      # Merge extra overrides from pass-through.
      overrides.merge!(@extra_ct_overrides) { |_k, generated, _extra| generated }

      parts = [XML_HEADER, %(<Types xmlns="#{CT_NS}">)]
      defaults.each { |ext, ct| parts << %(<Default Extension="#{ext}" ContentType="#{ct}"/>) }
      overrides.each { |pn, ct| parts << %(<Override PartName="#{pn}" ContentType="#{ct}"/>) }
      parts << "</Types>"
      parts.join
    end

    def generate_rels_root
      rid_counter = 1
      parts = [
        XML_HEADER,
        %(<Relationships xmlns="#{REL_NS}">),
        %(<Relationship Id="rId#{rid_counter}" Type="#{DOC_REL_NS}/officeDocument" Target="xl/workbook.xml"/>)
      ]
      unless @core_properties.empty?
        rid_counter += 1
        parts << %(<Relationship Id="rId#{rid_counter}" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>)
      end
      if @app_properties.any?
        rid_counter += 1
        parts << %(<Relationship Id="rId#{rid_counter}" Type="#{DOC_REL_NS}/extended-properties" Target="docProps/app.xml"/>)
      end
      unless @custom_properties.empty?
        rid_counter += 1
        parts << %(<Relationship Id="rId#{rid_counter}" Type="#{DOC_REL_NS}/custom-properties" Target="docProps/custom.xml"/>)
      end
      parts << "</Relationships>"
      parts.join
    end

    def generate_workbook_xml(pivot_cache_count = 0)
      wb_attrs = %(xmlns="#{SSML_NS}" xmlns:r="#{DOC_REL_NS}")
      wb_attrs << %( conformance="#{@workbook_properties[:conformance]}") if @workbook_properties[:conformance]
      parts = [
        XML_HEADER,
        "<workbook #{wb_attrs}>"
      ]

      # fileVersion
      unless @file_version.empty?
        fv_attrs = []
        fv_attrs << %(appName="#{xml_escape(@file_version[:app_name])}") if @file_version[:app_name]
        fv_attrs << %(lastEdited="#{@file_version[:last_edited]}") if @file_version[:last_edited]
        fv_attrs << %(lowestEdited="#{@file_version[:lowest_edited]}") if @file_version[:lowest_edited]
        fv_attrs << %(rupBuild="#{@file_version[:rup_build]}") if @file_version[:rup_build]
        fv_attrs << %(codeName="#{xml_escape(@file_version[:code_name])}") if @file_version[:code_name]
        parts << "<fileVersion #{fv_attrs.join(" ")}/>" unless fv_attrs.empty?
      end

      # fileSharing
      unless @file_sharing.empty?
        fs_attrs = []
        fs_attrs << 'readOnlyRecommended="1"' if @file_sharing[:read_only_recommended]
        fs_attrs << %(userName="#{xml_escape(@file_sharing[:user_name])}") if @file_sharing[:user_name]
        fs_attrs << %(algorithmName="#{xml_escape(@file_sharing[:algorithm_name])}") if @file_sharing[:algorithm_name]
        fs_attrs << %(hashValue="#{@file_sharing[:hash_value]}") if @file_sharing[:hash_value]
        fs_attrs << %(saltValue="#{@file_sharing[:salt_value]}") if @file_sharing[:salt_value]
        fs_attrs << %(spinCount="#{@file_sharing[:spin_count]}") if @file_sharing[:spin_count]
        parts << "<fileSharing #{fs_attrs.join(" ")}/>" unless fs_attrs.empty?
      end

      # workbookPr
      unless @workbook_properties.empty?
        attrs = []
        attrs << %(date1904="#{@workbook_properties[:date1904] ? 1 : 0}") unless @workbook_properties[:date1904].nil?
        attrs << %(dateCompatibility="0") if @workbook_properties[:date_compatibility] == false
        attrs << %(defaultThemeVersion="#{@workbook_properties[:default_theme_version]}") if @workbook_properties[:default_theme_version]
        attrs << %(codeName="#{xml_escape(@workbook_properties[:code_name])}") if @workbook_properties[:code_name]
        attrs << %(filterPrivacy="1") if @workbook_properties[:filter_privacy]
        attrs << %(autoCompressPictures="0") if @workbook_properties[:auto_compress_pictures] == false
        attrs << %(backupFile="1") if @workbook_properties[:backup_file]
        attrs << %(showObjects="#{xml_escape(@workbook_properties[:show_objects])}") if @workbook_properties[:show_objects]
        attrs << %(updateLinks="#{xml_escape(@workbook_properties[:update_links])}") if @workbook_properties[:update_links]
        attrs << %(refreshAllConnections="1") if @workbook_properties[:refresh_all_connections]
        attrs << %(checkCompatibility="1") if @workbook_properties[:check_compatibility]
        attrs << %(hidePivotFieldList="1") if @workbook_properties[:hide_pivot_field_list]
        attrs << %(showBorderUnselectedTables="0") if @workbook_properties[:show_border_unselected_tables] == false
        attrs << %(promptedSolutions="1") if @workbook_properties[:prompted_solutions]
        attrs << %(showInkAnnotation="0") if @workbook_properties[:show_ink_annotation] == false
        attrs << %(saveExternalLinkValues="0") if @workbook_properties[:save_external_link_values] == false
        attrs << %(showPivotChartFilter="1") if @workbook_properties[:show_pivot_chart_filter]
        attrs << %(allowRefreshQuery="1") if @workbook_properties[:allow_refresh_query]
        attrs << %(publishItems="1") if @workbook_properties[:publish_items]
        parts << "<workbookPr #{attrs.join(" ")}/>" unless attrs.empty?
      end

      # bookViews/workbookView
      unless @workbook_views.empty?
        attrs = []
        attrs << %(visibility="#{@workbook_views[:visibility]}") if @workbook_views[:visibility]
        attrs << 'minimized="1"' if @workbook_views[:minimized]
        shs = @workbook_views[:show_horizontal_scroll]
        attrs << %(showHorizontalScroll="#{shs ? 1 : 0}") unless shs.nil?
        svs = @workbook_views[:show_vertical_scroll]
        attrs << %(showVerticalScroll="#{svs ? 1 : 0}") unless svs.nil?
        sst = @workbook_views[:show_sheet_tabs]
        attrs << %(showSheetTabs="#{sst ? 1 : 0}") unless sst.nil?
        attrs << %(xWindow="#{@workbook_views[:x_window]}") if @workbook_views[:x_window]
        attrs << %(yWindow="#{@workbook_views[:y_window]}") if @workbook_views[:y_window]
        attrs << %(windowWidth="#{@workbook_views[:window_width]}") if @workbook_views[:window_width]
        attrs << %(windowHeight="#{@workbook_views[:window_height]}") if @workbook_views[:window_height]
        attrs << %(tabRatio="#{@workbook_views[:tab_ratio]}") if @workbook_views[:tab_ratio]
        attrs << %(firstSheet="#{@workbook_views[:first_sheet]}") if @workbook_views[:first_sheet]
        attrs << %(activeTab="#{@workbook_views[:active_tab]}") if @workbook_views[:active_tab]
        afdg = @workbook_views[:auto_filter_date_grouping]
        attrs << %(autoFilterDateGrouping="#{afdg ? 1 : 0}") unless afdg.nil?
        parts << "<bookViews>"
        parts << "<workbookView #{attrs.join(" ")}/>" unless attrs.empty?
        parts << "</bookViews>"
      end

      parts << "<sheets>"
      @sheet_order.each_with_index do |name, i|
        state = @sheet_states[name]
        state_attr = case state
                     when :hidden then ' state="hidden"'
                     when :very_hidden then ' state="veryHidden"'
                     else ""
                     end
        parts << %(<sheet name="#{xml_escape(name)}" sheetId="#{i + 1}"#{state_attr} r:id="rId#{i + 1}"/>)
      end
      parts << "</sheets>"

      # workbookProtection
      if @workbook_protection
        wp_attrs = []
        wp_attrs << 'lockStructure="1"' if @workbook_protection[:lock_structure]
        wp_attrs << 'lockWindows="1"' if @workbook_protection[:lock_windows]
        wp_attrs << 'lockRevision="1"' if @workbook_protection[:lock_revision]
        if @workbook_protection[:algorithm_name]
          wp_attrs << %(workbookAlgorithmName="#{xml_escape(@workbook_protection[:algorithm_name])}")
          wp_attrs << %(workbookHashValue="#{xml_escape(@workbook_protection[:hash_value])}") if @workbook_protection[:hash_value]
          wp_attrs << %(workbookSaltValue="#{xml_escape(@workbook_protection[:salt_value])}") if @workbook_protection[:salt_value]
          wp_attrs << %(workbookSpinCount="#{@workbook_protection[:spin_count]}") if @workbook_protection[:spin_count]
        elsif @workbook_protection[:password]
          wp_attrs << %(workbookPassword="#{xml_escape(@workbook_protection[:password])}")
        end
        if @workbook_protection[:revisions_algorithm_name]
          wp_attrs << %(revisionsAlgorithmName="#{xml_escape(@workbook_protection[:revisions_algorithm_name])}")
          wp_attrs << %(revisionsHashValue="#{xml_escape(@workbook_protection[:revisions_hash_value])}") if @workbook_protection[:revisions_hash_value]
          wp_attrs << %(revisionsSaltValue="#{xml_escape(@workbook_protection[:revisions_salt_value])}") if @workbook_protection[:revisions_salt_value]
          wp_attrs << %(revisionsSpinCount="#{@workbook_protection[:revisions_spin_count]}") if @workbook_protection[:revisions_spin_count]
        elsif @workbook_protection[:revisions_password]
          wp_attrs << %(revisionsPassword="#{xml_escape(@workbook_protection[:revisions_password])}")
        end
        parts << "<workbookProtection #{wp_attrs.join(" ")}/>" unless wp_attrs.empty?
      end

      # externalReferences
      unless @external_links.empty?
        # rId for external links: after sheets + styles + optional SST + pivot caches
        el_rid_base = @sheet_order.size + 1 + (@use_shared_strings ? 1 : 0) + 1 + pivot_cache_count
        parts << "<externalReferences>"
        @external_links.each_with_index do |_, idx|
          parts << %(<externalReference r:id="rId#{el_rid_base + idx}"/>)
        end
        parts << "</externalReferences>"
      end

      # definedNames
      unless @defined_names.empty?
        parts << "<definedNames>"
        @defined_names.each do |dn|
          attrs = %(name="#{xml_escape(dn[:name])}")
          attrs << %( localSheetId="#{dn[:local_sheet_id]}") if dn[:local_sheet_id]
          attrs << ' hidden="1"' if dn[:hidden]
          attrs << %( comment="#{xml_escape(dn[:comment])}") if dn[:comment]
          attrs << %( description="#{xml_escape(dn[:description])}") if dn[:description]
          attrs << ' function="1"' if dn[:function]
          attrs << ' vbProcedure="1"' if dn[:vb_procedure]
          attrs << ' xlm="1"' if dn[:xlm]
          attrs << %( shortcutKey="#{xml_escape(dn[:shortcut_key])}") if dn[:shortcut_key]
          attrs << ' publishToServer="1"' if dn[:publish_to_server]
          attrs << ' workbookParameter="1"' if dn[:workbook_parameter]
          attrs << %( functionGroupId="#{dn[:function_group_id]}") if dn[:function_group_id]
          attrs << %( customMenu="#{xml_escape(dn[:custom_menu])}") if dn[:custom_menu]
          attrs << %( help="#{xml_escape(dn[:help])}") if dn[:help]
          attrs << %( statusBar="#{xml_escape(dn[:status_bar])}") if dn[:status_bar]
          parts << "<definedName #{attrs}>#{xml_escape(dn[:value])}</definedName>"
        end
        parts << "</definedNames>"
      end

      # calcPr
      unless @calc_properties.empty?
        attrs = []
        attrs << %(calcId="#{@calc_properties[:calc_id]}") if @calc_properties[:calc_id]
        attrs << %(calcMode="#{@calc_properties[:calc_mode]}") if @calc_properties[:calc_mode]
        attrs << %(fullCalcOnLoad="#{@calc_properties[:full_calc_on_load] ? 1 : 0}") unless @calc_properties[:full_calc_on_load].nil?
        attrs << %(iterate="#{@calc_properties[:iterate] ? 1 : 0}") unless @calc_properties[:iterate].nil?
        attrs << %(iterateCount="#{@calc_properties[:iterate_count]}") if @calc_properties[:iterate_count]
        attrs << %(iterateDelta="#{@calc_properties[:iterate_delta]}") if @calc_properties[:iterate_delta]
        attrs << %(refMode="#{@calc_properties[:ref_mode]}") if @calc_properties[:ref_mode]
        attrs << %(calcCompleted="#{@calc_properties[:calc_completed] ? 1 : 0}") unless @calc_properties[:calc_completed].nil?
        attrs << %(calcOnSave="#{@calc_properties[:calc_on_save] ? 1 : 0}") unless @calc_properties[:calc_on_save].nil?
        fp = @calc_properties[:full_precision]
        attrs << %(fullPrecision="#{fp ? 1 : 0}") unless fp.nil?
        conc = @calc_properties[:concurrent_calc]
        attrs << %(concurrentCalc="#{conc ? 1 : 0}") unless conc.nil?
        attrs << %(concurrentManualCount="#{@calc_properties[:concurrent_manual_count]}") if @calc_properties[:concurrent_manual_count]
        ffc = @calc_properties[:force_full_calc]
        attrs << %(forceFullCalc="#{ffc ? 1 : 0}") unless ffc.nil?
        parts << "<calcPr #{attrs.join(" ")}/>" unless attrs.empty?
      end

      # pivotCaches (reference from workbook to cache definition rels)
      if pivot_cache_count.positive?
        # rId layout: sheets(1..N), styles(N+1), optional SST(N+2), then pivot caches
        pivot_rid_base = @sheet_order.size + 1 + (@use_shared_strings ? 1 : 0) + 1
        parts << "<pivotCaches>"
        pivot_cache_count.times do |ci|
          parts << %(<pivotCache cacheId="#{ci + 1}" r:id="rId#{pivot_rid_base + ci}"/>)
        end
        parts << "</pivotCaches>"
      end

      # fileRecoveryPr
      unless @file_recovery_properties.empty?
        frp_attrs = []
        ar = @file_recovery_properties[:auto_recover]
        frp_attrs << %(autoRecover="#{ar ? 1 : 0}") unless ar.nil?
        cs = @file_recovery_properties[:crash_save]
        frp_attrs << %(crashSave="#{cs ? 1 : 0}") unless cs.nil?
        del = @file_recovery_properties[:data_extract_load]
        frp_attrs << %(dataExtractLoad="#{del ? 1 : 0}") unless del.nil?
        rl = @file_recovery_properties[:repair_load]
        frp_attrs << %(repairLoad="#{rl ? 1 : 0}") unless rl.nil?
        parts << "<fileRecoveryPr #{frp_attrs.join(" ")}/>" unless frp_attrs.empty?
      end

      parts << "</workbook>"
      parts.join
    end

    def generate_workbook_rels(has_calc_chain: false)
      parts = [
        XML_HEADER,
        %(<Relationships xmlns="#{REL_NS}">)
      ]
      @sheet_order.each_with_index do |_, i|
        parts << %(<Relationship Id="rId#{i + 1}" Type="#{DOC_REL_NS}/worksheet" Target="worksheets/sheet#{i + 1}.xml"/>)
      end
      next_rid = @sheet_order.size + 1
      parts << %(<Relationship Id="rId#{next_rid}" Type="#{DOC_REL_NS}/styles" Target="styles.xml"/>)
      next_rid += 1
      if @use_shared_strings
        parts << %(<Relationship Id="rId#{next_rid}" Type="#{DOC_REL_NS}/sharedStrings" Target="sharedStrings.xml"/>)
        next_rid += 1
      end
      (1..@pivot_cache_count.to_i).each do |c|
        parts << %(<Relationship Id="rId#{next_rid}" Type="#{DOC_REL_NS}/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition#{c}.xml"/>)
        next_rid += 1
      end
      @external_links.each_with_index do |_, idx|
        parts << %(<Relationship Id="rId#{next_rid}" Type="#{DOC_REL_NS}/externalLink" Target="externalLinks/externalLink#{idx + 1}.xml"/>)
        next_rid += 1
      end
      if has_calc_chain
        parts << %(<Relationship Id="rId#{next_rid}" Type="#{DOC_REL_NS}/calcChain" Target="calcChain.xml"/>)
        next_rid += 1
      end
      parts << "</Relationships>"
      parts.join
    end

    def generate_calc_chain_xml
      chain_entries = []
      @sheet_order.each_with_index do |sheet_name, i|
        @sheets[sheet_name].each do |address, value|
          chain_entries << { ref: address, sheet_id: i + 1 } if value.is_a?(Formula)
        end
      end
      return nil if chain_entries.empty?

      parts = [XML_HEADER, %(<calcChain xmlns="#{SSML_NS}">)]
      chain_entries.each do |entry|
        parts << %(<c r="#{entry[:ref]}" i="#{entry[:sheet_id]}"/>)
      end
      parts << "</calcChain>"
      parts.join
    end

    def generate_worksheet_xml(sheet_cells, sheet_col_widths, sheet_col_attrs, sheet_row_attrs, sheet_auto_filter, sheet_filter_cols, sheet_sort, sheet_merge_cells, sheet_hyperlinks, sheet_cell_styles, sheet_props, sheet_fmt, sheet_sv, sheet_fp, sheet_sel, sheet_po, sheet_pm, sheet_ps, sheet_hf, sheet_rb, sheet_cb, sheet_dv, sheet_cf, sst = nil, sheet_tables = [], hyperlink_count = 0, has_drawing: false, has_comments: false, sheet_prot: nil, vml_rid: nil, phonetic_pr: nil, dv_options: {}, prot_ranges: [], cell_watches: [], ignored_errors: [], data_consol: nil, sheet_scenarios: nil, cell_phonetic: {})
      needs_r_ns = !sheet_hyperlinks.empty? || sheet_tables.any? || has_drawing || has_comments
      worksheet_attrs = %(xmlns="#{SSML_NS}")
      worksheet_attrs << %( xmlns:r="#{DOC_REL_NS}") if needs_r_ns
      parts = [
        XML_HEADER,
        "<worksheet #{worksheet_attrs}>"
      ]

      # Emit <sheetPr> if sheet properties are defined.
      unless sheet_props.empty?
        sp_attrs = []
        sp_attrs << 'syncHorizontal="1"' if sheet_props[:sync_horizontal]
        sp_attrs << 'syncVertical="1"' if sheet_props[:sync_vertical]
        sp_attrs << %(syncRef="#{sheet_props[:sync_ref]}") if sheet_props[:sync_ref]
        sp_attrs << 'transitionEvaluation="1"' if sheet_props[:transition_evaluation]
        sp_attrs << 'transitionEntry="1"' if sheet_props[:transition_entry]
        sp_attrs << %(codeName="#{xml_escape(sheet_props[:code_name])}") if sheet_props[:code_name]
        sp_attrs << 'filterMode="1"' if sheet_props[:filter_mode]
        sp_attrs << 'published="0"' if sheet_props[:published] == false
        sp_attrs << 'enableFormatConditionsCalculation="0"' if sheet_props[:enable_format_conditions_calculation] == false
        sp_children = []
        if sheet_props[:tab_color]
          sp_children << %(<tabColor rgb="#{sheet_props[:tab_color]}"/>)
        elsif sheet_props[:tab_color_theme]
          tc_attrs = [%(theme="#{sheet_props[:tab_color_theme]}")]
          tc_attrs << %(tint="#{sheet_props[:tab_color_tint]}") if sheet_props[:tab_color_tint]
          sp_children << "<tabColor #{tc_attrs.join(" ")}/>"
        elsif sheet_props[:tab_color_indexed]
          sp_children << %(<tabColor indexed="#{sheet_props[:tab_color_indexed]}"/>)
        elsif sheet_props[:tab_color_auto]
          sp_children << '<tabColor auto="1"/>'
        end
        sb = sheet_props[:summary_below]
        sr = sheet_props[:summary_right]
        as = sheet_props[:apply_styles]
        sos = sheet_props[:show_outline_symbols]
        unless sb.nil? && sr.nil? && as.nil? && sos.nil?
          outline_attrs = []
          outline_attrs << %(applyStyles="#{as ? 1 : 0}") unless as.nil?
          outline_attrs << %(summaryBelow="#{sb ? 1 : 0}") unless sb.nil?
          outline_attrs << %(summaryRight="#{sr ? 1 : 0}") unless sr.nil?
          outline_attrs << %(showOutlineSymbols="#{sos ? 1 : 0}") unless sos.nil?
          sp_children << "<outlinePr #{outline_attrs.join(" ")}/>"
        end
        ftp = sheet_props[:fit_to_page]
        apb = sheet_props[:auto_page_breaks]
        unless ftp.nil? && apb.nil?
          psp_attrs = []
          psp_attrs << %(fitToPage="#{ftp ? 1 : 0}") unless ftp.nil?
          psp_attrs << %(autoPageBreaks="#{apb ? 1 : 0}") unless apb.nil?
          sp_children << "<pageSetUpPr #{psp_attrs.join(" ")}/>"
        end
        if !sp_children.empty? || !sp_attrs.empty?
          sp_open = sp_attrs.empty? ? "<sheetPr>" : "<sheetPr #{sp_attrs.join(" ")}>"
          parts << sp_open
          parts.concat(sp_children)
          parts << "</sheetPr>"
        end
      end

      # Emit <dimension> computed from cell addresses.
      parts << %(<dimension ref="#{compute_dimension(sheet_cells)}"/>) unless sheet_cells.empty?

      # Emit <sheetViews> if sheet view properties, freeze pane, or selection are defined.
      if !sheet_sv.empty? || sheet_fp || sheet_sel
        parts << "<sheetViews>"
        sv_attrs = []
        wp = sheet_sv[:window_protection]
        sv_attrs << %(windowProtection="#{wp ? 1 : 0}") unless wp.nil?
        sf = sheet_sv[:show_formulas]
        sv_attrs << %(showFormulas="#{sf ? 1 : 0}") unless sf.nil?
        sgl = sheet_sv[:show_grid_lines]
        sv_attrs << %(showGridLines="#{sgl ? 1 : 0}") unless sgl.nil?
        srch = sheet_sv[:show_row_col_headers]
        sv_attrs << %(showRowColHeaders="#{srch ? 1 : 0}") unless srch.nil?
        sz = sheet_sv[:show_zeros]
        sv_attrs << %(showZeros="#{sz ? 1 : 0}") unless sz.nil?
        rtl = sheet_sv[:right_to_left]
        sv_attrs << %(rightToLeft="#{rtl ? 1 : 0}") unless rtl.nil?
        sv_attrs << 'tabSelected="1"' if sheet_sv[:tab_selected]
        sr = sheet_sv[:show_ruler]
        sv_attrs << %(showRuler="#{sr ? 1 : 0}") unless sr.nil?
        sos = sheet_sv[:show_outline_symbols]
        sv_attrs << %(showOutlineSymbols="#{sos ? 1 : 0}") unless sos.nil?
        dgc = sheet_sv[:default_grid_color]
        sv_attrs << %(defaultGridColor="#{dgc ? 1 : 0}") unless dgc.nil?
        sws = sheet_sv[:show_white_space]
        sv_attrs << %(showWhiteSpace="#{sws ? 1 : 0}") unless sws.nil?
        sv_attrs << %(view="#{sheet_sv[:view]}") if sheet_sv[:view]
        sv_attrs << %(topLeftCell="#{sheet_sv[:top_left_cell]}") if sheet_sv[:top_left_cell]
        sv_attrs << %(colorId="#{sheet_sv[:color_id]}") if sheet_sv[:color_id]
        zs = sheet_sv[:zoom_scale]
        sv_attrs << %(zoomScale="#{zs}") if zs
        sv_attrs << %(zoomScaleNormal="#{sheet_sv[:zoom_scale_normal]}") if sheet_sv[:zoom_scale_normal]
        sv_attrs << %(zoomScaleSheetLayoutView="#{sheet_sv[:zoom_scale_sheet_layout_view]}") if sheet_sv[:zoom_scale_sheet_layout_view]
        sv_attrs << %(zoomScalePageLayoutView="#{sheet_sv[:zoom_scale_page_layout_view]}") if sheet_sv[:zoom_scale_page_layout_view]
        sv_attrs << 'workbookViewId="0"'
        parts << "<sheetView #{sv_attrs.join(" ")}>"

        if sheet_fp && sheet_fp[:state] == :split
          pane_attrs = []
          pane_attrs << %(xSplit="#{sheet_fp[:x_split]}") if sheet_fp[:x_split].to_i.positive?
          pane_attrs << %(ySplit="#{sheet_fp[:y_split]}") if sheet_fp[:y_split].to_i.positive?
          pane_attrs << %(topLeftCell="#{sheet_fp[:top_left_cell]}") if sheet_fp[:top_left_cell]
          has_x = sheet_fp[:x_split].to_i.positive?
          has_y = sheet_fp[:y_split].to_i.positive?
          active_pane = if has_y && has_x
                          "bottomRight"
                        elsif has_y
                          "bottomLeft"
                        else
                          "topRight"
                        end
          pane_attrs << %(activePane="#{active_pane}")
          parts << "<pane #{pane_attrs.join(" ")}/>"
        elsif sheet_fp && (sheet_fp[:row].to_i.positive? || sheet_fp[:col].to_i.positive?)
          top_left = "#{index_to_column_letter(sheet_fp[:col].to_i + 1)}#{sheet_fp[:row].to_i + 1}"
          pane_attrs = []
          pane_attrs << %(ySplit="#{sheet_fp[:row]}") if sheet_fp[:row].to_i.positive?
          pane_attrs << %(xSplit="#{sheet_fp[:col]}") if sheet_fp[:col].to_i.positive?
          pane_attrs << %(topLeftCell="#{top_left}")
          active_pane = if sheet_fp[:row].to_i.positive? && sheet_fp[:col].to_i.positive?
                          "bottomRight"
                        elsif sheet_fp[:row].to_i.positive?
                          "bottomLeft"
                        else
                          "topRight"
                        end
          pane_attrs << %(activePane="#{active_pane}")
          pane_attrs << 'state="frozen"'
          parts << "<pane #{pane_attrs.join(" ")}/>"
        end

        if sheet_sel
          sel_attrs = []
          sel_attrs << %(pane="#{sheet_sel[:pane]}") if sheet_sel[:pane]
          sel_attrs << %(activeCell="#{sheet_sel[:active_cell]}") if sheet_sel[:active_cell]
          sel_attrs << %(activeCellId="#{sheet_sel[:active_cell_id]}") if sheet_sel[:active_cell_id]
          sel_attrs << %(sqref="#{sheet_sel[:sqref]}") if sheet_sel[:sqref]
          parts << "<selection #{sel_attrs.join(" ")}/>"
        end

        parts << "</sheetView>"
        parts << "</sheetViews>"
      end

      # Emit <sheetFormatPr> if sheet format properties are defined.
      unless sheet_fmt.empty?
        fmt_attrs = []
        fmt_attrs << %(defaultRowHeight="#{sheet_fmt[:default_row_height]}") if sheet_fmt[:default_row_height]
        fmt_attrs << %(defaultColWidth="#{sheet_fmt[:default_col_width]}") if sheet_fmt[:default_col_width]
        fmt_attrs << %(baseColWidth="#{sheet_fmt[:base_col_width]}") if sheet_fmt[:base_col_width]
        fmt_attrs << %(outlineLevelRow="#{sheet_fmt[:outline_level_row]}") if sheet_fmt[:outline_level_row]
        fmt_attrs << %(outlineLevelCol="#{sheet_fmt[:outline_level_col]}") if sheet_fmt[:outline_level_col]
        fmt_attrs << 'customHeight="1"' if sheet_fmt[:custom_height]
        fmt_attrs << 'zeroHeight="1"' if sheet_fmt[:zero_height]
        fmt_attrs << 'thickTop="1"' if sheet_fmt[:thick_top]
        fmt_attrs << 'thickBottom="1"' if sheet_fmt[:thick_bottom]
        parts << "<sheetFormatPr #{fmt_attrs.join(" ")}/>"
      end

      # Emit <cols> if column widths or column attributes are defined.
      all_cols = sheet_col_widths.keys | sheet_col_attrs.keys
      unless all_cols.empty?
        parts << "<cols>"
        all_cols.sort_by { |col| column_letter_to_index(col) }.each do |col_letter|
          idx = column_letter_to_index(col_letter)
          width = sheet_col_widths[col_letter]
          ca = sheet_col_attrs[col_letter] || {}
          col_attrs = %(min="#{idx}" max="#{idx}")
          col_attrs << %( width="#{width}" customWidth="1") if width
          col_attrs << ' hidden="1"' if ca[:hidden]
          col_attrs << ' bestFit="1"' if ca[:best_fit]
          col_attrs << %( outlineLevel="#{ca[:outline_level]}") if ca[:outline_level]
          col_attrs << ' collapsed="1"' if ca[:collapsed]
          col_attrs << %( style="#{ca[:style]}") if ca[:style]
          col_attrs << ' phonetic="1"' if ca[:phonetic]
          parts << "<col #{col_attrs}/>"
        end
        parts << "</cols>"
      end

      parts << "<sheetData>"

      # Group cells by row number.
      cells_by_row = {}
      sheet_cells.each do |address, value|
        row_num = extract_row_number(address)
        col_letter = extract_column_letter(address)
        cells_by_row[row_num] ||= {}
        cells_by_row[row_num][col_letter] = value
      end

      # Include rows that have attributes but no cells.
      sheet_row_attrs.each_key { |rn| cells_by_row[rn] ||= {} }

      # Emit rows in ascending order.
      cells_by_row.sort.each do |row_num, row_cells|
        attrs = %(r="#{row_num}")
        ra = sheet_row_attrs[row_num]
        if ra
          attrs << %( ht="#{ra[:height]}" customHeight="1") if ra[:height]
          attrs << ' hidden="1"' if ra[:hidden]
          attrs << %( outlineLevel="#{ra[:outline_level]}") if ra[:outline_level]
          attrs << ' collapsed="1"' if ra[:collapsed]
          attrs << %( s="#{ra[:style]}" customFormat="1") if ra.key?(:style)
          attrs << ' thickTop="1"' if ra[:thick_top]
          attrs << ' thickBot="1"' if ra[:thick_bot]
          attrs << ' ph="1"' if ra[:ph]
        end
        parts << "<row #{attrs}>"
        row_cells.sort_by { |col, _| column_letter_to_index(col) }.each do |col_letter, value|
          cell_ref = "#{col_letter}#{row_num}"
          style_idx = resolve_style_index(sheet_cell_styles[cell_ref])
          parts << cell_xml(cell_ref, value, style_idx, sst, ph: cell_phonetic[cell_ref])
        end
        parts << "</row>"
      end

      parts << "</sheetData>"

      # Emit <sheetCalcPr> if fullCalcOnLoad is set.
      parts << '<sheetCalcPr fullCalcOnLoad="1"/>' if sheet_props[:full_calc_on_load]

      # Emit <sheetProtection> if defined.
      if sheet_prot
        sp_attrs = []
        sp_attrs << 'sheet="1"' if sheet_prot[:sheet] != false
        sp_attrs << 'objects="1"' if sheet_prot[:objects]
        sp_attrs << 'scenarios="1"' if sheet_prot[:scenarios]
        sp_attrs << 'formatCells="0"' if sheet_prot[:format_cells] == false
        sp_attrs << 'formatColumns="0"' if sheet_prot[:format_columns] == false
        sp_attrs << 'formatRows="0"' if sheet_prot[:format_rows] == false
        sp_attrs << 'insertColumns="0"' if sheet_prot[:insert_columns] == false
        sp_attrs << 'insertRows="0"' if sheet_prot[:insert_rows] == false
        sp_attrs << 'insertHyperlinks="0"' if sheet_prot[:insert_hyperlinks] == false
        sp_attrs << 'deleteColumns="0"' if sheet_prot[:delete_columns] == false
        sp_attrs << 'deleteRows="0"' if sheet_prot[:delete_rows] == false
        sp_attrs << 'selectLockedCells="1"' if sheet_prot[:select_locked_cells]
        sp_attrs << 'sort="0"' if sheet_prot[:sort] == false
        sp_attrs << 'autoFilter="0"' if sheet_prot[:auto_filter] == false
        sp_attrs << 'pivotTables="0"' if sheet_prot[:pivot_tables] == false
        sp_attrs << 'selectUnlockedCells="1"' if sheet_prot[:select_unlocked_cells]
        if sheet_prot[:algorithm_name]
          sp_attrs << %(algorithmName="#{xml_escape(sheet_prot[:algorithm_name])}")
          sp_attrs << %(hashValue="#{xml_escape(sheet_prot[:hash_value])}") if sheet_prot[:hash_value]
          sp_attrs << %(saltValue="#{xml_escape(sheet_prot[:salt_value])}") if sheet_prot[:salt_value]
          sp_attrs << %(spinCount="#{sheet_prot[:spin_count]}") if sheet_prot[:spin_count]
        elsif sheet_prot[:password]
          sp_attrs << %(password="#{xml_escape(sheet_prot[:password])}")
        end
        parts << "<sheetProtection #{sp_attrs.join(" ")}/>"
      end

      # Emit <protectedRanges> if defined.
      unless prot_ranges.empty?
        parts << "<protectedRanges>"
        prot_ranges.each do |pr|
          pr_attrs = [%(sqref="#{pr[:sqref]}"), %(name="#{xml_escape(pr[:name])}")]
          if pr[:algorithm_name]
            pr_attrs << %(algorithmName="#{xml_escape(pr[:algorithm_name])}")
            pr_attrs << %(hashValue="#{xml_escape(pr[:hash_value])}") if pr[:hash_value]
            pr_attrs << %(saltValue="#{xml_escape(pr[:salt_value])}") if pr[:salt_value]
            pr_attrs << %(spinCount="#{pr[:spin_count]}") if pr[:spin_count]
          end
          sds = pr[:security_descriptors] || []
          if sds.empty?
            parts << "<protectedRange #{pr_attrs.join(" ")}/>"
          else
            parts << "<protectedRange #{pr_attrs.join(" ")}>"
            sds.each { |sd| parts << "<securityDescriptor>#{xml_escape(sd)}</securityDescriptor>" }
            parts << "</protectedRange>"
          end
        end
        parts << "</protectedRanges>"
      end

      # Emit <scenarios> if defined.
      if sheet_scenarios
        sc_attrs = []
        sc_attrs << %(current="#{sheet_scenarios[:current]}") if sheet_scenarios[:current]
        sc_attrs << %(show="#{sheet_scenarios[:show]}") if sheet_scenarios[:show]
        sc_attrs << %(sqref="#{sheet_scenarios[:sqref]}") if sheet_scenarios[:sqref]
        items = sheet_scenarios[:scenarios] || []
        sc_attr_str = sc_attrs.empty? ? "" : " #{sc_attrs.join(" ")}"
        parts << "<scenarios#{sc_attr_str}>"
        items.each do |sc|
          s_attrs = [%(name="#{xml_escape(sc[:name])}")]
          s_attrs << 'locked="1"' if sc[:locked]
          s_attrs << 'hidden="1"' if sc[:hidden]
          s_attrs << %(count="#{sc[:input_cells].size}") if sc[:input_cells]&.any?
          s_attrs << %(user="#{xml_escape(sc[:user])}") if sc[:user]
          s_attrs << %(comment="#{xml_escape(sc[:comment])}") if sc[:comment]
          parts << "<scenario #{s_attrs.join(" ")}>"
          (sc[:input_cells] || []).each do |ic|
            ic_attrs = [%(r="#{ic[:r]}"), %(val="#{xml_escape(ic[:val].to_s)}")]
            ic_attrs << 'deleted="1"' if ic[:deleted]
            ic_attrs << 'undone="1"' if ic[:undone]
            ic_attrs << %(numFmtId="#{ic[:num_fmt_id]}") if ic[:num_fmt_id]
            parts << "<inputCells #{ic_attrs.join(" ")}/>"
          end
          parts << "</scenario>"
        end
        parts << "</scenarios>"
      end

      # Emit <autoFilter> with optional filterColumns.
      if sheet_auto_filter
        if sheet_filter_cols.empty?
          parts << %(<autoFilter ref="#{sheet_auto_filter}"/>)
        else
          parts << %(<autoFilter ref="#{sheet_auto_filter}">)
          sheet_filter_cols.sort.each do |col_id, filter|
            fc_attrs = %( colId="#{col_id}")
            fc_attrs << ' hiddenButton="1"' if filter[:hidden_button]
            fc_attrs << ' showButton="0"' if filter[:show_button] == false
            parts << "<filterColumn#{fc_attrs}>"
            parts << emit_filter_xml(filter)
            parts << "</filterColumn>"
          end
          parts << "</autoFilter>"
        end
      end

      # Emit <sortState> if defined.
      if sheet_sort
        ss_attrs = %( ref="#{sheet_sort[:ref]}")
        ss_attrs << ' columnSort="1"' if sheet_sort[:column_sort]
        ss_attrs << ' caseSensitive="1"' if sheet_sort[:case_sensitive]
        ss_attrs << %( sortMethod="#{sheet_sort[:sort_method]}") if sheet_sort[:sort_method]
        parts << "<sortState#{ss_attrs}>"
        sheet_sort[:sort_conditions].each do |sc|
          sc_attrs = %(ref="#{sc[:ref]}")
          sc_attrs << ' descending="1"' if sc[:descending]
          sc_attrs << %( sortBy="#{sc[:sort_by]}") if sc[:sort_by]
          sc_attrs << %( customList="#{xml_escape(sc[:custom_list])}") if sc[:custom_list]
          sc_attrs << %( dxfId="#{sc[:dxf_id]}") if sc[:dxf_id]
          sc_attrs << %( iconSet="#{sc[:icon_set]}") if sc[:icon_set]
          sc_attrs << %( iconId="#{sc[:icon_id]}") if sc[:icon_id]
          parts << "<sortCondition #{sc_attrs}/>"
        end
        parts << "</sortState>"
      end

      # Emit <dataConsolidate> if defined.
      if data_consol
        dc_attrs = []
        dc_attrs << %(function="#{data_consol[:function]}") if data_consol[:function] && data_consol[:function] != "sum"
        dc_attrs << 'startLabels="1"' if data_consol[:start_labels]
        dc_attrs << 'leftLabels="1"' if data_consol[:left_labels]
        dc_attrs << 'topLabels="1"' if data_consol[:top_labels]
        dc_attrs << 'link="1"' if data_consol[:link]
        refs = data_consol[:data_refs] || []
        dc_attr_str = dc_attrs.empty? ? "" : " #{dc_attrs.join(" ")}"
        if refs.empty?
          parts << "<dataConsolidate#{dc_attr_str}/>"
        else
          parts << "<dataConsolidate#{dc_attr_str}>"
          parts << %(<dataRefs count="#{refs.size}">)
          refs.each do |r|
            ref_attrs = []
            ref_attrs << %(ref="#{r[:ref]}") if r[:ref]
            ref_attrs << %(name="#{xml_escape(r[:name])}") if r[:name]
            ref_attrs << %(sheet="#{xml_escape(r[:sheet])}") if r[:sheet]
            parts << "<dataRef #{ref_attrs.join(" ")}/>"
          end
          parts << "</dataRefs>"
          parts << "</dataConsolidate>"
        end
      end

      # Emit <mergeCells> if merge ranges are defined.
      unless sheet_merge_cells.empty?
        parts << %(<mergeCells count="#{sheet_merge_cells.size}">)
        sheet_merge_cells.each { |ref| parts << %(<mergeCell ref="#{ref}"/>) }
        parts << "</mergeCells>"
      end

      # Emit <phoneticPr> if defined.
      if phonetic_pr
        pp_attrs = []
        pp_attrs << %(fontId="#{phonetic_pr[:font_id]}") if phonetic_pr[:font_id]
        pp_attrs << %(type="#{phonetic_pr[:type]}") if phonetic_pr[:type]
        pp_attrs << %(alignment="#{phonetic_pr[:alignment]}") if phonetic_pr[:alignment]
        parts << "<phoneticPr #{pp_attrs.join(" ")}/>"
      end

      # Emit <hyperlinks> if hyperlinks are defined.
      unless sheet_hyperlinks.empty?
        parts << "<hyperlinks>"
        ext_rid = 0
        sheet_hyperlinks.each do |(cell_ref, link)|
          attrs = %(ref="#{cell_ref}")
          if link[:url]
            ext_rid += 1
            attrs << %( r:id="rId#{ext_rid}")
          end
          attrs << %( display="#{xml_escape(link[:display])}") if link[:display]
          attrs << %( tooltip="#{xml_escape(link[:tooltip])}") if link[:tooltip]
          attrs << %( location="#{xml_escape(link[:location])}") if link[:location]
          parts << %(<hyperlink #{attrs}/>)
        end
        parts << "</hyperlinks>"
      end

      # Emit <conditionalFormatting> if defined.
      unless sheet_cf.empty?
        sheet_cf.group_by { |cf| cf[:sqref] }.each do |sqref, rules|
          cf_attrs = %( sqref="#{sqref}")
          cf_attrs << ' pivot="1"' if rules.any? { |r| r[:pivot] }
          parts << "<conditionalFormatting#{cf_attrs}>"
          rules.each do |cf|
            emit_cf_rule(parts, cf)
          end
          parts << "</conditionalFormatting>"
        end
      end

      # Emit <dataValidations> if defined.
      unless sheet_dv.empty?
        dv_container_attrs = %( count="#{sheet_dv.size}")
        dv_container_attrs << ' disablePrompts="1"' if dv_options[:disable_prompts]
        dv_container_attrs << %( xWindow="#{dv_options[:x_window]}") if dv_options[:x_window]
        dv_container_attrs << %( yWindow="#{dv_options[:y_window]}") if dv_options[:y_window]
        parts << "<dataValidations#{dv_container_attrs}>"
        sheet_dv.each do |dv|
          dv_attrs = %(sqref="#{dv[:sqref]}")
          dv_attrs << %( type="#{dv[:type]}") if dv[:type]
          dv_attrs << %( operator="#{dv[:operator]}") if dv[:operator]
          dv_attrs << %( errorStyle="#{dv[:error_style]}") if dv[:error_style]
          dv_attrs << ' allowBlank="1"' if dv[:allow_blank]
          dv_attrs << ' showInputMessage="1"' if dv[:show_input_message]
          dv_attrs << ' showErrorMessage="1"' if dv[:show_error_message]
          dv_attrs << %( errorTitle="#{xml_escape(dv[:error_title])}") if dv[:error_title]
          dv_attrs << %( error="#{xml_escape(dv[:error])}") if dv[:error]
          dv_attrs << %( promptTitle="#{xml_escape(dv[:prompt_title])}") if dv[:prompt_title]
          dv_attrs << %( prompt="#{xml_escape(dv[:prompt])}") if dv[:prompt]
          dv_attrs << ' showDropDown="1"' if dv[:show_drop_down]
          dv_attrs << %( imeMode="#{dv[:ime_mode]}") if dv[:ime_mode]
          if dv[:formula1] || dv[:formula2]
            parts << "<dataValidation #{dv_attrs}>"
            parts << "<formula1>#{xml_escape(dv[:formula1])}</formula1>" if dv[:formula1]
            parts << "<formula2>#{xml_escape(dv[:formula2])}</formula2>" if dv[:formula2]
            parts << "</dataValidation>"
          else
            parts << "<dataValidation #{dv_attrs}/>"
          end
        end
        parts << "</dataValidations>"
      end

      # Emit <printOptions> if defined.
      unless sheet_po.empty?
        po_attrs = []
        po_attrs << 'gridLines="1"' if sheet_po[:grid_lines]
        po_attrs << 'headings="1"' if sheet_po[:headings]
        po_attrs << 'horizontalCentered="1"' if sheet_po[:horizontal_centered]
        po_attrs << 'verticalCentered="1"' if sheet_po[:vertical_centered]
        po_attrs << 'gridLinesSet="0"' if sheet_po[:grid_lines_set] == false
        parts << "<printOptions #{po_attrs.join(" ")}/>" unless po_attrs.empty?
      end

      # Emit <pageMargins> if defined.
      if sheet_pm
        pm_attrs = %w[left right top bottom header footer].filter_map do |k|
          v = sheet_pm[k.to_sym]
          %(#{k}="#{v}") if v
        end
        parts << "<pageMargins #{pm_attrs.join(" ")}/>" unless pm_attrs.empty?
      end

      # Emit <pageSetup> if defined.
      unless sheet_ps.empty?
        ps_attrs = []
        ps_attrs << %(orientation="#{sheet_ps[:orientation]}") if sheet_ps[:orientation]
        ps_attrs << %(paperSize="#{sheet_ps[:paper_size]}") if sheet_ps[:paper_size]
        ps_attrs << %(scale="#{sheet_ps[:scale]}") if sheet_ps[:scale]
        ps_attrs << %(fitToWidth="#{sheet_ps[:fit_to_width]}") if sheet_ps[:fit_to_width]
        ps_attrs << %(fitToHeight="#{sheet_ps[:fit_to_height]}") if sheet_ps[:fit_to_height]
        ps_attrs << %(pageOrder="#{sheet_ps[:page_order]}") if sheet_ps[:page_order]
        ps_attrs << 'blackAndWhite="1"' if sheet_ps[:black_and_white]
        ps_attrs << 'draft="1"' if sheet_ps[:draft]
        ps_attrs << %(cellComments="#{sheet_ps[:cell_comments]}") if sheet_ps[:cell_comments]
        ps_attrs << %(firstPageNumber="#{sheet_ps[:first_page_number]}") if sheet_ps[:first_page_number]
        ps_attrs << 'useFirstPageNumber="1"' if sheet_ps[:use_first_page_number]
        ps_attrs << %(horizontalDpi="#{sheet_ps[:horizontal_dpi]}") if sheet_ps[:horizontal_dpi]
        ps_attrs << %(verticalDpi="#{sheet_ps[:vertical_dpi]}") if sheet_ps[:vertical_dpi]
        ps_attrs << %(copies="#{sheet_ps[:copies]}") if sheet_ps[:copies]
        ps_attrs << %(paperHeight="#{sheet_ps[:paper_height]}") if sheet_ps[:paper_height]
        ps_attrs << %(paperWidth="#{sheet_ps[:paper_width]}") if sheet_ps[:paper_width]
        ps_attrs << %(errors="#{sheet_ps[:errors]}") if sheet_ps[:errors]
        ps_attrs << 'usePrinterDefaults="0"' if sheet_ps[:use_printer_defaults] == false
        parts << "<pageSetup #{ps_attrs.join(" ")}/>" unless ps_attrs.empty?
      end

      # Emit <headerFooter> if defined.
      unless sheet_hf.empty?
        hf_attrs = []
        hf_attrs << 'differentFirst="1"' if sheet_hf[:different_first]
        hf_attrs << 'differentOddEven="1"' if sheet_hf[:different_odd_even]
        hf_attrs << 'scaleWithDoc="0"' if sheet_hf[:scale_with_doc] == false
        hf_attrs << 'alignWithMargins="0"' if sheet_hf[:align_with_margins] == false
        hf_tag = hf_attrs.empty? ? "<headerFooter>" : "<headerFooter #{hf_attrs.join(" ")}>"
        parts << hf_tag
        parts << "<oddHeader>#{xml_escape(sheet_hf[:odd_header])}</oddHeader>" if sheet_hf[:odd_header]
        parts << "<oddFooter>#{xml_escape(sheet_hf[:odd_footer])}</oddFooter>" if sheet_hf[:odd_footer]
        parts << "<evenHeader>#{xml_escape(sheet_hf[:even_header])}</evenHeader>" if sheet_hf[:even_header]
        parts << "<evenFooter>#{xml_escape(sheet_hf[:even_footer])}</evenFooter>" if sheet_hf[:even_footer]
        parts << "<firstHeader>#{xml_escape(sheet_hf[:first_header])}</firstHeader>" if sheet_hf[:first_header]
        parts << "<firstFooter>#{xml_escape(sheet_hf[:first_footer])}</firstFooter>" if sheet_hf[:first_footer]
        parts << "</headerFooter>"
      end

      # Emit <rowBreaks> if defined.
      unless sheet_rb.empty?
        manual_count = sheet_rb.count { |r| r.is_a?(Hash) ? r.fetch(:man, true) : true }
        parts << %(<rowBreaks count="#{sheet_rb.size}" manualBreakCount="#{manual_count}">)
        sheet_rb.each { |r| parts << emit_brk_xml(r, default_max: 16_383) }
        parts << "</rowBreaks>"
      end

      # Emit <colBreaks> if defined.
      unless sheet_cb.empty?
        manual_count = sheet_cb.count { |c| c.is_a?(Hash) ? c.fetch(:man, true) : true }
        parts << %(<colBreaks count="#{sheet_cb.size}" manualBreakCount="#{manual_count}">)
        sheet_cb.each { |c| parts << emit_brk_xml(c, default_max: 1_048_575) }
        parts << "</colBreaks>"
      end

      # Emit <cellWatches> if defined.
      unless cell_watches.empty?
        parts << "<cellWatches>"
        cell_watches.each { |r| parts << %(<cellWatch r="#{r}"/>) }
        parts << "</cellWatches>"
      end

      # Emit <ignoredErrors> if defined.
      unless ignored_errors.empty?
        parts << "<ignoredErrors>"
        ignored_errors.each do |ie|
          ie_attrs = %( sqref="#{ie[:sqref]}")
          ie_attrs << ' evalError="1"' if ie[:eval_error]
          ie_attrs << ' twoDigitTextYear="1"' if ie[:two_digit_text_year]
          ie_attrs << ' numberStoredAsText="1"' if ie[:number_stored_as_text]
          ie_attrs << ' formula="1"' if ie[:formula]
          ie_attrs << ' formulaRange="1"' if ie[:formula_range]
          ie_attrs << ' unlockedFormula="1"' if ie[:unlocked_formula]
          ie_attrs << ' emptyCellReference="1"' if ie[:empty_cell_reference]
          ie_attrs << ' listDataValidation="1"' if ie[:list_data_validation]
          ie_attrs << ' calculatedColumn="1"' if ie[:calculated_column]
          parts << "<ignoredError#{ie_attrs}/>"
        end
        parts << "</ignoredErrors>"
      end

      # Emit <tableParts> if tables are defined.
      unless sheet_tables.empty?
        parts << %(<tableParts count="#{sheet_tables.size}">)
        sheet_tables.each_with_index do |_tbl, i|
          rid = hyperlink_count + i + 1
          parts << %(<tablePart r:id="rId#{rid}"/>)
        end
        parts << "</tableParts>"
      end

      # Emit <drawing> reference if images or charts exist.
      if has_drawing
        # The drawing rId is after hyperlinks + tables + comments + vml
        drawing_rid = hyperlink_count + sheet_tables.size + (has_comments ? 2 : 0) + 1
        parts << %(<drawing r:id="rId#{drawing_rid}"/>)
      end

      # Emit <legacyDrawing> reference if comments exist (VML shapes).
      parts << %(<legacyDrawing r:id="rId#{vml_rid}"/>) if vml_rid

      parts << "</worksheet>"
      parts.join
    end

    def build_shared_string_table
      rt_sst = {}.compare_by_identity # RichText object -> global index
      str_sst = {} # String -> global index
      @sheets.each_value do |sheet_cells|
        sheet_cells.each_value do |value|
          total = rt_sst.size + str_sst.size
          case value
          when RichText
            rt_sst[value] = total unless rt_sst.key?(value)
          when String
            str_sst[value] = total unless str_sst.key?(value)
          else
            next if value.is_a?(Numeric) || value.is_a?(Date) || value == true || value == false || value.is_a?(Formula)

            str = value.to_s
            str_sst[str] = total unless str_sst.key?(str)
          end
        end
      end
      [rt_sst, str_sst]
    end

    def generate_shared_strings_xml(sst)
      rt_sst, str_sst = sst
      total = rt_sst.size + str_sst.size
      entries = Array.new(total)
      rt_sst.each { |rt, idx| entries[idx] = { type: :rt, value: rt } }
      str_sst.each { |str, idx| entries[idx] = { type: :str, value: str } }

      parts = [XML_HEADER, %(<sst xmlns="#{SSML_NS}" count="#{total}" uniqueCount="#{total}">)]
      entries.each do |entry|
        parts << if entry[:type] == :rt
                   "<si>#{rich_text_xml(entry[:value])}</si>"
                 else
                   "<si><t>#{xml_escape(entry[:value])}</t></si>"
                 end
      end
      parts << "</sst>"
      parts.join
    end

    def generate_table_xml(tbl)
      trc = tbl[:totals_row_count].to_i
      table_attrs = %(xmlns="#{SSML_NS}" id="#{tbl[:id]}" name="#{xml_escape(tbl[:name])}" displayName="#{xml_escape(tbl[:display_name])}" ref="#{tbl[:ref]}")
      table_attrs << %( comment="#{xml_escape(tbl[:comment])}") if tbl[:comment]
      hrc = tbl[:header_row_count]
      table_attrs << %( headerRowCount="#{hrc}") if hrc && hrc != 1
      table_attrs << ' insertRow="1"' if tbl[:insert_row]
      table_attrs << ' insertRowShift="1"' if tbl[:insert_row_shift]
      table_attrs << %( totalsRowCount="#{trc}") if trc.positive?
      table_attrs << ' totalsRowShown="0"' if trc.zero?
      table_attrs << ' published="1"' if tbl[:published]
      table_attrs << %( headerRowDxfId="#{tbl[:header_row_dxf_id]}") if tbl[:header_row_dxf_id]
      table_attrs << %( dataDxfId="#{tbl[:data_dxf_id]}") if tbl[:data_dxf_id]
      table_attrs << %( totalsRowDxfId="#{tbl[:totals_row_dxf_id]}") if tbl[:totals_row_dxf_id]
      table_attrs << %( headerRowBorderDxfId="#{tbl[:header_row_border_dxf_id]}") if tbl[:header_row_border_dxf_id]
      table_attrs << %( tableBorderDxfId="#{tbl[:table_border_dxf_id]}") if tbl[:table_border_dxf_id]
      table_attrs << %( totalsRowBorderDxfId="#{tbl[:totals_row_border_dxf_id]}") if tbl[:totals_row_border_dxf_id]
      table_attrs << %( headerRowCellStyle="#{xml_escape(tbl[:header_row_cell_style])}") if tbl[:header_row_cell_style]
      table_attrs << %( totalsRowCellStyle="#{xml_escape(tbl[:totals_row_cell_style])}") if tbl[:totals_row_cell_style]
      table_attrs << %( connectionId="#{tbl[:connection_id]}") if tbl[:connection_id]
      table_attrs << %( tableType="#{xml_escape(tbl[:table_type])}") if tbl[:table_type]
      parts = [
        XML_HEADER,
        "<table #{table_attrs}>",
        %(<autoFilter ref="#{tbl[:ref]}"/>),
        %(<tableColumns count="#{tbl[:columns].size}">)
      ]
      tbl[:columns].each_with_index do |col, i|
        col_name = col.is_a?(Hash) ? col[:name] : col
        col_attrs = %(id="#{i + 1}" name="#{xml_escape(col_name)}")
        if col.is_a?(Hash)
          col_attrs << %( totalsRowFunction="#{col[:totals_row_function]}") if col[:totals_row_function]
          col_attrs << %( totalsRowLabel="#{xml_escape(col[:totals_row_label])}") if col[:totals_row_label]
          col_attrs << %( dataDxfId="#{col[:data_dxf_id]}") if col[:data_dxf_id]
          col_attrs << %( totalsRowDxfId="#{col[:totals_row_dxf_id]}") if col[:totals_row_dxf_id]
          col_attrs << %( headerRowDxfId="#{col[:header_row_dxf_id]}") if col[:header_row_dxf_id]
          col_attrs << %( dataCellStyle="#{xml_escape(col[:data_cell_style])}") if col[:data_cell_style]
          if col[:calculated_column_formula] || col[:totals_row_formula]
            parts << "<tableColumn #{col_attrs}>"
            parts << "<calculatedColumnFormula>#{xml_escape(col[:calculated_column_formula])}</calculatedColumnFormula>" if col[:calculated_column_formula]
            parts << "<totalsRowFormula>#{xml_escape(col[:totals_row_formula])}</totalsRowFormula>" if col[:totals_row_formula]
            parts << "</tableColumn>"
          else
            parts << "<tableColumn #{col_attrs}/>"
          end
        else
          parts << "<tableColumn #{col_attrs}/>"
        end
      end
      parts << "</tableColumns>"
      style = tbl[:style] || {}
      style_name = style[:name] || "TableStyleMedium2"
      sfc = style[:show_first_column] ? "1" : "0"
      slc = style[:show_last_column] ? "1" : "0"
      srs = if style.key?(:show_row_stripes)
              style[:show_row_stripes] ? "1" : "0"
            else
              "1"
            end
      scs = style[:show_column_stripes] ? "1" : "0"
      parts << %(<tableStyleInfo name="#{xml_escape(style_name)}" showFirstColumn="#{sfc}" showLastColumn="#{slc}" showRowStripes="#{srs}" showColumnStripes="#{scs}"/>)
      parts << "</table>"
      parts.join
    end

    def generate_worksheet_rels(sheet_hyperlinks, sheet_tables = [], table_start_index = 0)
      parts = [
        XML_HEADER,
        %(<Relationships xmlns="#{REL_NS}">)
      ]
      rid = 0
      sheet_hyperlinks.each do |(_cell_ref, link)|
        next unless link[:url]

        rid += 1
        parts << %(<Relationship Id="rId#{rid}" Type="#{DOC_REL_NS}/hyperlink" Target="#{xml_escape(link[:url])}" TargetMode="External"/>)
      end
      sheet_tables.each_with_index do |_tbl, i|
        rid += 1
        parts << %(<Relationship Id="rId#{rid}" Type="#{DOC_REL_NS}/table" Target="../tables/table#{table_start_index + i + 1}.xml"/>)
      end
      parts << "</Relationships>"
      parts.join
    end

    def build_sheet_rels_parts_v2(sheet_name, sheet_tables, table_start_index, drawing_idx, comment_idx, pivot_start, pivot_count)
      rels = []
      @hyperlinks[sheet_name].each do |(_cell_ref, link)|
        next unless link[:url]

        rels << { type: "#{DOC_REL_NS}/hyperlink", target: link[:url], external: true }
      end
      sheet_tables.each_with_index do |_tbl, i|
        rels << { type: "#{DOC_REL_NS}/table", target: "../tables/table#{table_start_index + i + 1}.xml" }
      end
      if comment_idx
        rels << { type: "#{DOC_REL_NS}/comments", target: "../comments#{comment_idx}.xml" }
        rels << { type: "#{DOC_REL_NS}/vmlDrawing", target: "../drawings/vmlDrawing#{comment_idx}.vml" }
      end
      rels << { type: "#{DOC_REL_NS}/drawing", target: "../drawings/drawing#{drawing_idx}.xml" } if drawing_idx
      pivot_count.times do |i|
        pt_idx = pivot_start + i + 1
        rels << { type: "#{DOC_REL_NS}/pivotTable", target: "../pivotTables/pivotTable#{pt_idx}.xml" }
      end
      rels
    end

    def generate_generic_rels(rels_data)
      parts = [XML_HEADER, %(<Relationships xmlns="#{REL_NS}">)]
      rels_data.each_with_index do |rel, i|
        ext_attr = rel[:external] ? ' TargetMode="External"' : ""
        parts << %(<Relationship Id="rId#{i + 1}" Type="#{rel[:type]}" Target="#{xml_escape(rel[:target])}"#{ext_attr}/>)
      end
      parts << "</Relationships>"
      parts.join
    end

    def generate_drawing_xml(drawing_parts)
      parts = [
        XML_HEADER,
        %(<xdr:wsDr xmlns:xdr="#{XDR_NS}" xmlns:a="#{A_NS}" xmlns:r="#{DOC_REL_NS}">)
      ]

      drawing_parts.each do |dp|
        case dp[:kind]
        when :pic
          img = dp[:img]
          rid = "rId#{dp[:rid_index]}"
          ea = img[:edit_as] || "oneCell"
          img_pub_attr = img[:published] ? ' fPublished="1"' : ""
          parts << %(<xdr:twoCellAnchor editAs="#{xml_escape(ea)}"#{img_pub_attr}>)
          parts << anchor_xml("from", img[:from_col], img[:from_row], col_off: img[:from_col_off] || 0, row_off: img[:from_row_off] || 0)
          parts << anchor_xml("to", img[:to_col], img[:to_row], col_off: img[:to_col_off] || 0, row_off: img[:to_row_off] || 0)
          macro_attr = img[:macro] ? %( macro="#{xml_escape(img[:macro])}") : ""
          parts << "<xdr:pic#{macro_attr}>"
          descr_attr = img[:description] ? %( descr="#{xml_escape(img[:description])}") : ""
          title_attr = img[:title] ? %( title="#{xml_escape(img[:title])}") : ""
          hidden_attr = img[:hidden] ? ' hidden="1"' : ""
          pic_lock_attrs = +""
          pic_lock_attrs << ' noChangeAspect="1"' if img[:no_change_aspect]
          pic_lock_attrs << ' noCrop="1"' if img[:no_crop]
          pic_locks = pic_lock_attrs.empty? ? "<a:picLocks/>" : "<a:picLocks#{pic_lock_attrs}/>"
          parts << %(<xdr:nvPicPr><xdr:cNvPr id="#{dp[:rid_index] + 1}" name="#{xml_escape(img[:name])}"#{descr_attr}#{title_attr}#{hidden_attr}/><xdr:cNvPicPr>#{pic_locks}</xdr:cNvPicPr></xdr:nvPicPr>)
          parts << %(<xdr:blipFill><a:blip r:embed="#{rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>)
          img_line_xml = if img[:line_color]
                           ln_w_attr = img[:line_width] ? %( w="#{img[:line_width].to_i}") : ""
                           %(<a:ln#{ln_w_attr}><a:solidFill><a:srgbClr val="#{xml_escape(img[:line_color])}"/></a:solidFill></a:ln>)
                         else
                           ""
                         end
          img_xfrm_xml = img[:rotation] ? %(<a:xfrm rot="#{img[:rotation].to_i}"><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>) : ""
          parts << %(<xdr:spPr>#{img_xfrm_xml}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>#{img_line_xml}</xdr:spPr>)
          parts << "</xdr:pic>"
          parts << client_data_xml(img)
          parts << "</xdr:twoCellAnchor>"
        when :chart
          chart = dp[:chart]
          rid = "rId#{dp[:rid_index]}"
          chart_ea_attr = chart[:edit_as] ? %( editAs="#{xml_escape(chart[:edit_as])}") : ""
          chart_pub_attr = chart[:published] ? ' fPublished="1"' : ""
          parts << "<xdr:twoCellAnchor#{chart_ea_attr}#{chart_pub_attr}>"
          parts << anchor_xml("from", chart[:from_col], chart[:from_row], col_off: chart[:from_col_off] || 0, row_off: chart[:from_row_off] || 0)
          parts << anchor_xml("to", chart[:to_col], chart[:to_row], col_off: chart[:to_col_off] || 0, row_off: chart[:to_row_off] || 0)
          gf_macro = chart[:frame_macro] ? xml_escape(chart[:frame_macro]) : ""
          parts << %(<xdr:graphicFrame macro="#{gf_macro}">)
          chart_frame_name = chart[:name] || chart[:title] || "Chart"
          chart_descr_attr = chart[:description] ? %( descr="#{xml_escape(chart[:description])}") : ""
          chart_title_attr = chart[:frame_title] ? %( title="#{xml_escape(chart[:frame_title])}") : ""
          chart_hidden_attr = chart[:frame_hidden] ? ' hidden="1"' : ""
          cnv_gf_pr = if chart[:frame_no_grp]
                        '<xdr:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></xdr:cNvGraphicFramePr>'
                      else
                        "<xdr:cNvGraphicFramePr/>"
                      end
          parts << %(<xdr:nvGraphicFramePr><xdr:cNvPr id="#{dp[:rid_index] + 1}" name="#{xml_escape(chart_frame_name)}"#{chart_descr_attr}#{chart_title_attr}#{chart_hidden_attr}/>#{cnv_gf_pr}</xdr:nvGraphicFramePr>)
          parts << '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="5000000" cy="3000000"/></xdr:xfrm>'
          parts << %(<a:graphic><a:graphicData uri="#{C_NS}"><c:chart xmlns:c="#{C_NS}" r:id="#{rid}"/></a:graphicData></a:graphic>)
          parts << "</xdr:graphicFrame>"
          parts << client_data_xml(chart)
          parts << "</xdr:twoCellAnchor>"
        when :sp
          shape = dp[:shape]
          shape_ea_attr = shape[:edit_as] ? %( editAs="#{xml_escape(shape[:edit_as])}") : ""
          shape_pub_attr = shape[:published] ? ' fPublished="1"' : ""
          parts << "<xdr:twoCellAnchor#{shape_ea_attr}#{shape_pub_attr}>"
          parts << anchor_xml("from", shape[:from_col], shape[:from_row], col_off: shape[:from_col_off] || 0, row_off: shape[:from_row_off] || 0)
          parts << anchor_xml("to", shape[:to_col], shape[:to_row], col_off: shape[:to_col_off] || 0, row_off: shape[:to_row_off] || 0)
          sp_macro_attr = shape[:macro] ? %( macro="#{xml_escape(shape[:macro])}") : ""
          sp_textlink_attr = shape[:textlink] ? %( textlink="#{xml_escape(shape[:textlink])}") : ""
          parts << "<xdr:sp#{sp_macro_attr}#{sp_textlink_attr}>"
          shape_descr_attr = shape[:description] ? %( descr="#{xml_escape(shape[:description])}") : ""
          shape_title_attr = shape[:title] ? %( title="#{xml_escape(shape[:title])}") : ""
          shape_hidden_attr = shape[:hidden] ? ' hidden="1"' : ""
          sp_lock_attrs = +""
          sp_lock_attrs << ' noGrp="1"' if shape[:no_grp]
          sp_lock_attrs << ' noRot="1"' if shape[:no_rot]
          sp_lock_attrs << ' fLocksText="1"' if shape[:f_locks_text]
          cnv_sp_pr = if sp_lock_attrs.empty?
                        "<xdr:cNvSpPr/>"
                      else
                        "<xdr:cNvSpPr><a:spLocks#{sp_lock_attrs}/></xdr:cNvSpPr>"
                      end
          parts << %(<xdr:nvSpPr><xdr:cNvPr id="#{dp[:id]}" name="#{xml_escape(shape[:name])}"#{shape_descr_attr}#{shape_title_attr}#{shape_hidden_attr}/>#{cnv_sp_pr}</xdr:nvSpPr>)
          shape_fill_xml = if shape[:no_fill]
                             "<a:noFill/>"
                           elsif shape[:fill_color]
                             %(<a:solidFill><a:srgbClr val="#{xml_escape(shape[:fill_color])}"/></a:solidFill>)
                           else
                             ""
                           end
          shape_line_xml = if shape[:no_line]
                             "<a:ln><a:noFill/></a:ln>"
                           elsif shape[:line_color]
                             ln_w_attr = shape[:line_width] ? %( w="#{shape[:line_width].to_i}") : ""
                             %(<a:ln#{ln_w_attr}><a:solidFill><a:srgbClr val="#{xml_escape(shape[:line_color])}"/></a:solidFill></a:ln>)
                           else
                             ""
                           end
          shape_xfrm_xml = shape[:rotation] ? %(<a:xfrm rot="#{shape[:rotation].to_i}"><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>) : ""
          parts << %(<xdr:spPr>#{shape_xfrm_xml}#{shape_fill_xml}<a:prstGeom prst="#{xml_escape(shape[:preset])}"><a:avLst/></a:prstGeom>#{shape_line_xml}</xdr:spPr>)
          if shape[:text]
            body_pr_attrs = +""
            body_pr_attrs << %( wrap="#{xml_escape(shape[:text_wrap])}") if shape[:text_wrap]
            body_pr_attrs << %( anchor="#{xml_escape(shape[:text_anchor])}") if shape[:text_anchor]
            body_pr_attrs << %( vertOverflow="#{xml_escape(shape[:text_vert_overflow])}") if shape[:text_vert_overflow]
            parts << "<xdr:txBody><a:bodyPr#{body_pr_attrs}/><a:lstStyle/>"
            parts << "<a:p><a:r><a:t>#{xml_escape(shape[:text])}</a:t></a:r></a:p>"
            parts << "</xdr:txBody>"
          end
          parts << "</xdr:sp>"
          parts << client_data_xml(shape)
          parts << "</xdr:twoCellAnchor>"
        end
      end

      parts << "</xdr:wsDr>"
      parts.join
    end

    def anchor_xml(tag, col, row, col_off: 0, row_off: 0)
      "<xdr:#{tag}><xdr:col>#{col}</xdr:col><xdr:colOff>#{col_off}</xdr:colOff><xdr:row>#{row}</xdr:row><xdr:rowOff>#{row_off}</xdr:rowOff></xdr:#{tag}>"
    end

    def client_data_xml(obj)
      cd_attrs = +""
      cd_attrs << ' fLocksWithSheet="0"' if obj[:locks_with_sheet] == false
      cd_attrs << ' fLocksWithSheet="1"' if obj[:locks_with_sheet] == true
      cd_attrs << ' fPrintsWithSheet="0"' if obj[:prints_with_sheet] == false
      cd_attrs << ' fPrintsWithSheet="1"' if obj[:prints_with_sheet] == true
      "<xdr:clientData#{cd_attrs}/>"
    end

    def generate_drawing_rels(rels_data)
      parts = [XML_HEADER, %(<Relationships xmlns="#{REL_NS}">)]
      rels_data.each_with_index do |rel, i|
        rel_type = case rel[:type]
                   when :image then "#{DOC_REL_NS}/image"
                   when :chart then "#{DOC_REL_NS}/chart"
                   end
        parts << %(<Relationship Id="rId#{i + 1}" Type="#{rel_type}" Target="#{rel[:target]}"/>)
      end
      parts << "</Relationships>"
      parts.join
    end

    def generate_chart_xml(chart)
      chart_type = CHART_TYPE_MAP[chart[:type]] || "barChart"
      no_axes = NO_AXIS_CHARTS.include?(chart_type)
      parts = [
        XML_HEADER,
        %(<c:chartSpace xmlns:c="#{C_NS}" xmlns:a="#{A_NS}" xmlns:r="#{DOC_REL_NS}">)
      ]
      rc = chart[:rounded_corners]
      parts << %(<c:roundedCorners val="#{rc ? 1 : 0}"/>) unless rc.nil?
      parts << %(<c:style val="#{chart[:style]}"/>) if chart[:style]
      parts << "<c:chart>"

      parts << "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>#{xml_escape(chart[:title])}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>" if chart[:title]
      atd = chart[:auto_title_deleted]
      parts << %(<c:autoTitleDeleted val="#{atd ? 1 : 0}"/>) unless atd.nil?

      if (v3d = chart[:view_3d])
        parts << "<c:view3D>"
        parts << %(<c:rotX val="#{v3d[:rot_x]}"/>) if v3d[:rot_x]
        parts << %(<c:hPercent val="#{v3d[:h_percent]}"/>) if v3d[:h_percent]
        parts << %(<c:rotY val="#{v3d[:rot_y]}"/>) if v3d[:rot_y]
        parts << %(<c:depthPercent val="#{v3d[:depth_percent]}"/>) if v3d[:depth_percent]
        r_ang = v3d[:r_ang_ax]
        parts << %(<c:rAngAx val="#{r_ang ? 1 : 0}"/>) unless r_ang.nil?
        parts << %(<c:perspective val="#{v3d[:perspective]}"/>) if v3d[:perspective]
        parts << "</c:view3D>"
      end

      parts << "<c:plotArea><c:layout/>"
      parts << "<c:#{chart_type}>"
      if %w[barChart bar3DChart].include?(chart_type)
        bd = chart[:bar_dir] || "col"
        gr = chart[:grouping] || "clustered"
        parts << %(<c:barDir val="#{bd}"/><c:grouping val="#{gr}"/>)
      elsif GROUPING_CHARTS.include?(chart_type)
        parts << %(<c:grouping val="#{chart[:grouping] || "standard"}"/>)
      end
      parts << %(<c:scatterStyle val="#{chart[:scatter_style] || "lineMarker"}"/>) if chart_type == "scatterChart"
      parts << %(<c:radarStyle val="#{chart[:radar_style] || "standard"}"/>) if chart_type == "radarChart"
      vc = chart[:vary_colors]
      parts << %(<c:varyColors val="#{vc ? 1 : 0}"/>) unless vc.nil?
      wf = chart[:wireframe]
      parts << %(<c:wireframe val="#{wf ? 1 : 0}"/>) unless wf.nil?

      all_series = chart[:series] || []
      all_series.each_with_index do |ser, idx|
        parts << "<c:ser><c:idx val=\"#{idx}\"/><c:order val=\"#{idx}\"/>"
        parts << "<c:tx><c:strRef><c:f>#{xml_escape(ser[:name])}</c:f></c:strRef></c:tx>" if ser[:name]
        if ser[:fill_color] || ser[:line_color] || ser[:line_width]
          parts << "<c:spPr>"
          parts << %(<a:solidFill><a:srgbClr val="#{xml_escape(ser[:fill_color])}"/></a:solidFill>) if ser[:fill_color]
          if ser[:line_color] || ser[:line_width]
            lw = ser[:line_width] ? %( w="#{(ser[:line_width] * 12_700).to_i}") : ""
            parts << "<a:ln#{lw}>"
            parts << %(<a:solidFill><a:srgbClr val="#{xml_escape(ser[:line_color])}"/></a:solidFill>) if ser[:line_color]
            parts << "</a:ln>"
          end
          parts << "</c:spPr>"
        end
        if chart[:data_labels]
          dl = chart[:data_labels]
          parts << "<c:dLbls>"
          parts << %(<c:dLblPos val="#{dl[:position]}"/>) if dl[:position]
          parts << "<c:showLegendKey val=\"#{dl[:show_legend_key] ? 1 : 0}\"/>" unless dl[:show_legend_key].nil?
          parts << "<c:showVal val=\"#{dl[:show_val] ? 1 : 0}\"/>" unless dl[:show_val].nil?
          parts << "<c:showCatName val=\"#{dl[:show_cat_name] ? 1 : 0}\"/>" unless dl[:show_cat_name].nil?
          parts << "<c:showSerName val=\"#{dl[:show_ser_name] ? 1 : 0}\"/>" unless dl[:show_ser_name].nil?
          parts << "<c:showPercent val=\"#{dl[:show_percent] ? 1 : 0}\"/>" unless dl[:show_percent].nil?
          parts << "<c:showBubbleSize val=\"#{dl[:show_bubble_size] ? 1 : 0}\"/>" unless dl[:show_bubble_size].nil?
          parts << "<c:separator>#{xml_escape(dl[:separator])}</c:separator>" if dl[:separator]
          parts << "</c:dLbls>"
        end
        parts << "<c:cat><c:strRef><c:f>#{xml_escape(ser[:cat_ref])}</c:f></c:strRef></c:cat>" if ser[:cat_ref]
        parts << "<c:val><c:numRef><c:f>#{xml_escape(ser[:val_ref])}</c:f></c:numRef></c:val>" if ser[:val_ref]
        parts << "</c:ser>"
      end

      parts << %(<c:gapWidth val="#{chart[:gap_width]}"/>) if chart[:gap_width]
      parts << %(<c:gapDepth val="#{chart[:gap_depth]}"/>) if chart[:gap_depth]
      parts << %(<c:overlap val="#{chart[:overlap]}"/>) if chart[:overlap]
      parts << %(<c:shape val="#{chart[:bar_shape]}"/>) if chart[:bar_shape]
      b3d = chart[:bubble_3d]
      parts << %(<c:bubble3D val="#{b3d ? 1 : 0}"/>) unless b3d.nil?
      parts << %(<c:bubbleScale val="#{chart[:bubble_scale]}"/>) if chart[:bubble_scale]
      snb = chart[:show_neg_bubbles]
      parts << %(<c:showNegBubbles val="#{snb ? 1 : 0}"/>) unless snb.nil?
      parts << %(<c:sizeRepresents val="#{chart[:size_represents]}"/>) if chart[:size_represents]
      parts << %(<c:firstSliceAng val="#{chart[:first_slice_ang]}"/>) if chart[:first_slice_ang]
      parts << %(<c:holeSize val="#{chart[:hole_size]}"/>) if chart[:hole_size]
      mk = chart[:marker]
      parts << %(<c:marker val="#{mk ? 1 : 0}"/>) unless mk.nil?
      sm = chart[:smooth]
      parts << %(<c:smooth val="#{sm ? 1 : 0}"/>) unless sm.nil?
      parts << '<c:axId val="1"/><c:axId val="2"/>' unless no_axes
      parts << "</c:#{chart_type}>"

      unless no_axes
        cat_del = chart[:cat_axis_delete] ? 1 : 0
        cat_orient = chart[:cat_axis_orientation] || "minMax"
        parts << %(<c:catAx><c:axId val="1"/><c:scaling>)
        parts << %(<c:logBase val="#{chart[:cat_axis_log_base]}"/>) if chart[:cat_axis_log_base]
        parts << %(<c:orientation val="#{cat_orient}"/>)
        parts << %(<c:max val="#{chart[:cat_axis_scaling_max]}"/>) if chart[:cat_axis_scaling_max]
        parts << %(<c:min val="#{chart[:cat_axis_scaling_min]}"/>) if chart[:cat_axis_scaling_min]
        parts << %(</c:scaling><c:delete val="#{cat_del}"/><c:axPos val="#{chart[:cat_axis_pos] || "b"}"/>)
        parts << "<c:majorGridlines/>" if chart[:cat_axis_major_gridlines]
        parts << "<c:minorGridlines/>" if chart[:cat_axis_minor_gridlines]
        parts << "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>#{xml_escape(chart[:cat_axis_title])}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>" if chart[:cat_axis_title]
        if (cnf = chart[:cat_axis_num_fmt])
          sl = cnf[:source_linked] ? 1 : 0
          parts << %(<c:numFmt formatCode="#{xml_escape(cnf[:format_code])}" sourceLinked="#{sl}"/>)
        end
        parts << %(<c:majorTickMark val="#{chart[:cat_axis_major_tick_mark]}"/>) if chart[:cat_axis_major_tick_mark]
        parts << %(<c:minorTickMark val="#{chart[:cat_axis_minor_tick_mark]}"/>) if chart[:cat_axis_minor_tick_mark]
        parts << %(<c:tickLblPos val="#{chart[:cat_axis_tick_lbl_pos]}"/>) if chart[:cat_axis_tick_lbl_pos]
        parts << '<c:crossAx val="2"/>'
        parts << %(<c:crosses val="#{chart[:cat_axis_crosses]}"/>) if chart[:cat_axis_crosses]
        parts << %(<c:crossesAt val="#{chart[:cat_axis_crosses_at]}"/>) if !chart[:cat_axis_crosses] && chart[:cat_axis_crosses_at]
        parts << %(<c:lblOffset val="#{chart[:cat_axis_lbl_offset]}"/>) if chart[:cat_axis_lbl_offset]
        parts << %(<c:tickLblSkip val="#{chart[:cat_axis_tick_lbl_skip]}"/>) if chart[:cat_axis_tick_lbl_skip]
        parts << %(<c:tickMarkSkip val="#{chart[:cat_axis_tick_mark_skip]}"/>) if chart[:cat_axis_tick_mark_skip]
        nml = chart[:cat_axis_no_multi_lvl_lbl]
        parts << %(<c:noMultiLvlLbl val="#{nml ? 1 : 0}"/>) unless nml.nil?
        parts << "</c:catAx>"
        val_del = chart[:val_axis_delete] ? 1 : 0
        val_orient = chart[:val_axis_orientation] || "minMax"
        parts << %(<c:valAx><c:axId val="2"/><c:scaling>)
        parts << %(<c:logBase val="#{chart[:val_axis_log_base]}"/>) if chart[:val_axis_log_base]
        parts << %(<c:orientation val="#{val_orient}"/>)
        parts << %(<c:max val="#{chart[:val_axis_scaling_max]}"/>) if chart[:val_axis_scaling_max]
        parts << %(<c:min val="#{chart[:val_axis_scaling_min]}"/>) if chart[:val_axis_scaling_min]
        parts << %(</c:scaling><c:delete val="#{val_del}"/><c:axPos val="#{chart[:val_axis_pos] || "l"}"/>)
        parts << "<c:majorGridlines/>" if chart[:val_axis_major_gridlines]
        parts << "<c:minorGridlines/>" if chart[:val_axis_minor_gridlines]
        parts << "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>#{xml_escape(chart[:val_axis_title])}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>" if chart[:val_axis_title]
        if (vnf = chart[:val_axis_num_fmt])
          sl = vnf[:source_linked] ? 1 : 0
          parts << %(<c:numFmt formatCode="#{xml_escape(vnf[:format_code])}" sourceLinked="#{sl}"/>)
        end
        parts << %(<c:majorTickMark val="#{chart[:val_axis_major_tick_mark]}"/>) if chart[:val_axis_major_tick_mark]
        parts << %(<c:minorTickMark val="#{chart[:val_axis_minor_tick_mark]}"/>) if chart[:val_axis_minor_tick_mark]
        parts << %(<c:tickLblPos val="#{chart[:val_axis_tick_lbl_pos]}"/>) if chart[:val_axis_tick_lbl_pos]
        parts << '<c:crossAx val="1"/>'
        parts << %(<c:crosses val="#{chart[:val_axis_crosses]}"/>) if chart[:val_axis_crosses]
        parts << %(<c:crossesAt val="#{chart[:val_axis_crosses_at]}"/>) if !chart[:val_axis_crosses] && chart[:val_axis_crosses_at]
        parts << %(<c:crossBetween val="#{chart[:val_axis_cross_between]}"/>) if chart[:val_axis_cross_between]
        parts << %(<c:majorUnit val="#{chart[:val_axis_major_unit]}"/>) if chart[:val_axis_major_unit]
        parts << %(<c:minorUnit val="#{chart[:val_axis_minor_unit]}"/>) if chart[:val_axis_minor_unit]
        parts << %(<c:dispUnits><c:builtInUnit val="#{chart[:val_axis_disp_units]}"/></c:dispUnits>) if chart[:val_axis_disp_units]
        parts << "</c:valAx>"
      end

      if chart[:data_table]
        dt = chart[:data_table]
        parts << "<c:dTable>"
        parts << %(<c:showHorzBorder val="#{dt[:show_horz_border] ? 1 : 0}"/>) unless dt[:show_horz_border].nil?
        parts << %(<c:showVertBorder val="#{dt[:show_vert_border] ? 1 : 0}"/>) unless dt[:show_vert_border].nil?
        parts << %(<c:showOutline val="#{dt[:show_outline] ? 1 : 0}"/>) unless dt[:show_outline].nil?
        parts << %(<c:showKeys val="#{dt[:show_keys] ? 1 : 0}"/>) unless dt[:show_keys].nil?
        parts << "</c:dTable>"
      end

      parts << %(<c:spPr><a:solidFill><a:srgbClr val="#{xml_escape(chart[:plot_area_fill])}"/></a:solidFill></c:spPr>) if chart[:plot_area_fill]

      parts << "</c:plotArea>"
      legend_pos = chart.dig(:legend, :position) || "r"
      legend_overlay = chart.dig(:legend, :overlay)
      parts << %(<c:legend><c:legendPos val="#{legend_pos}"/>)
      legend_entries = chart.dig(:legend, :entries)
      legend_entries&.each do |entry|
        parts << %(<c:legendEntry><c:idx val="#{entry[:idx]}"/>)
        parts << %(<c:delete val="#{entry[:delete] ? 1 : 0}"/>) unless entry[:delete].nil?
        parts << "</c:legendEntry>"
      end
      parts << %(<c:overlay val="#{legend_overlay ? 1 : 0}"/>) unless legend_overlay.nil?
      parts << %(</c:legend>)
      pvo = chart[:plot_vis_only]
      parts << %(<c:plotVisOnly val="#{pvo ? 1 : 0}"/>) unless pvo.nil?
      parts << %(<c:dispBlanksAs val="#{chart[:disp_blanks_as]}"/>) if chart[:disp_blanks_as]
      sdom = chart[:show_d_lbls_over_max]
      parts << %(<c:showDLblsOverMax val="#{sdom ? 1 : 0}"/>) unless sdom.nil?
      parts << "</c:chart></c:chartSpace>"
      parts.join
    end

    def generate_comments_xml(sheet_comments)
      authors = sheet_comments.map { |c| c[:author] }.uniq
      parts = [
        XML_HEADER,
        %(<comments xmlns="#{SSML_NS}">),
        "<authors>"
      ]
      authors.each { |a| parts << "<author>#{xml_escape(a)}</author>" }
      parts << "</authors><commentList>"
      sheet_comments.each do |c|
        aid = authors.index(c[:author]) || 0
        text_xml = if c[:text].is_a?(RichText)
                     rich_text_xml(c[:text])
                   else
                     "<r><t>#{xml_escape(c[:text])}</t></r>"
                   end
        comment_attrs = %(ref="#{c[:ref]}" authorId="#{aid}")
        comment_attrs << %( guid="#{c[:guid]}") if c[:guid]
        comment_attrs << %( shapeId="#{c[:shape_id]}") if c[:shape_id]
        parts << "<comment #{comment_attrs}><text>#{text_xml}</text></comment>"
      end
      parts << "</commentList></comments>"
      parts.join
    end

    def generate_vml_drawing_xml(sheet_comments)
      parts = [
        '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">',
        '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>',
        '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">',
        '<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/>',
        "</v:shapetype>"
      ]
      sheet_comments.each_with_index do |c, idx|
        col, row = cell_to_col_row(c[:ref])
        shape_id = 1025 + idx
        parts << %(<v:shape id="_x0000_s#{shape_id}" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:#{idx + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">)
        parts << '<v:fill color2="#ffffe1"/>'
        parts << '<v:shadow on="t" color="black" obscured="t"/>'
        parts << '<v:path o:connecttype="none"/>'
        parts << '<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>'
        parts << '<x:ClientData ObjectType="Note">'
        parts << "<x:MoveWithCells/><x:SizeWithCells/>"
        parts << "<x:Anchor>#{col + 1}, 15, #{row}, 10, #{col + 3}, 15, #{row + 4}, 4</x:Anchor>"
        parts << "<x:AutoFill>False</x:AutoFill>"
        parts << "<x:Row>#{row}</x:Row>"
        parts << "<x:Column>#{col}</x:Column>"
        parts << "</x:ClientData></v:shape>"
      end
      parts << "</xml>"
      parts.join
    end

    def cell_to_col_row(cell_ref)
      m = cell_ref.match(/\A([A-Z]+)(\d+)\z/)
      return [0, 0] unless m

      col = column_letter_to_index(m[1]) - 1
      row = m[2].to_i - 1
      [col, row]
    end

    def generate_pivot_table_xml(pivot_table, cache_id)
      data_caption = pivot_table[:data_caption] || (pivot_table[:data_fields].first ? pivot_table[:data_fields].first[:name] : "Values")
      pt_attrs = %( xmlns="#{SSML_NS}" name="#{xml_escape(pivot_table[:name])}" cacheId="#{cache_id}" dataCaption="#{xml_escape(data_caption)}")
      pt_attrs << %( grandTotalCaption="#{xml_escape(pivot_table[:grand_total_caption])}") if pivot_table[:grand_total_caption]
      pt_attrs << %( errorCaption="#{xml_escape(pivot_table[:error_caption])}") if pivot_table[:error_caption]
      pt_attrs << ' showError="1"' if pivot_table[:show_error]
      pt_attrs << %( missingCaption="#{xml_escape(pivot_table[:missing_caption])}") if pivot_table[:missing_caption]
      pt_attrs << ' showMissing="0"' if pivot_table[:show_missing] == false
      pt_attrs << %( tag="#{xml_escape(pivot_table[:tag])}") if pivot_table[:tag]
      pt_attrs << %( dataOnRows="1") if pivot_table[:data_on_rows]
      pt_attrs << %( dataOnRows="0") unless pivot_table[:data_on_rows]
      pt_attrs << %( rowGrandTotals="0") if pivot_table[:row_grand_totals] == false
      pt_attrs << %( colGrandTotals="0") if pivot_table[:col_grand_totals] == false
      pt_attrs << %( compact="0") if pivot_table[:compact] == false
      pt_attrs << %( outline="0") if pivot_table[:outline] == false
      pt_attrs << %( outlineData="1") if pivot_table[:outline_data]
      pt_attrs << %( compactData="0") if pivot_table[:compact_data] == false
      pt_attrs << %( showHeaders="0") if pivot_table[:show_headers] == false
      pt_attrs << %( showMultipleLabel="0") if pivot_table[:show_multiple_label] == false
      pt_attrs << %( showDataDropDown="0") if pivot_table[:show_data_drop_down] == false
      pt_attrs << %( indent="#{pivot_table[:indent]}") if pivot_table[:indent]
      pt_attrs << ' published="1"' if pivot_table[:published]
      pt_attrs << ' editData="1"' if pivot_table[:edit_data]
      pt_attrs << ' disableFieldList="1"' if pivot_table[:disable_field_list]
      pt_attrs << ' visualTotals="0"' if pivot_table[:visual_totals] == false
      pt_attrs << ' printDrill="1"' if pivot_table[:print_drill]
      pt_attrs << %( createdVersion="#{pivot_table[:created_version]}") if pivot_table[:created_version]
      pt_attrs << %( updatedVersion="#{pivot_table[:updated_version]}") if pivot_table[:updated_version]
      pt_attrs << %( minRefreshableVersion="#{pivot_table[:min_refreshable_version]}") if pivot_table[:min_refreshable_version]
      anf = pivot_table.fetch(:apply_number_formats, false)
      abf = pivot_table.fetch(:apply_border_formats, false)
      aff = pivot_table.fetch(:apply_font_formats, false)
      apf = pivot_table.fetch(:apply_pattern_formats, false)
      aaf = pivot_table.fetch(:apply_alignment_formats, false)
      awf = pivot_table.fetch(:apply_width_height_formats, true)
      pt_attrs << %( applyNumberFormats="#{anf ? 1 : 0}")
      pt_attrs << %( applyBorderFormats="#{abf ? 1 : 0}")
      pt_attrs << %( applyFontFormats="#{aff ? 1 : 0}")
      pt_attrs << %( applyPatternFormats="#{apf ? 1 : 0}")
      pt_attrs << %( applyAlignmentFormats="#{aaf ? 1 : 0}")
      pt_attrs << %( applyWidthHeightFormats="#{awf ? 1 : 0}")
      pt_attrs << ' multipleFieldFilters="0"' if pivot_table[:multiple_field_filters] == false
      pt_attrs << ' showDrill="0"' if pivot_table[:show_drill] == false
      pt_attrs << ' showDataTips="0"' if pivot_table[:show_data_tips] == false
      pt_attrs << ' enableDrill="0"' if pivot_table[:enable_drill] == false
      pt_attrs << ' showMemberPropertyTips="0"' if pivot_table[:show_member_property_tips] == false
      pt_attrs << ' itemPrintTitles="1"' if pivot_table[:item_print_titles]
      pt_attrs << ' fieldPrintTitles="1"' if pivot_table[:field_print_titles]
      pt_attrs << ' preserveFormatting="0"' if pivot_table[:preserve_formatting] == false
      pt_attrs << ' pageOverThenDown="1"' if pivot_table[:page_over_then_down]
      pt_attrs << %( pageWrap="#{pivot_table[:page_wrap]}") if pivot_table[:page_wrap]
      parts = [
        XML_HEADER,
        "<pivotTableDefinition#{pt_attrs}>"
      ]

      # Compute field count from source range or explicit field_names.
      field_count = if pivot_table[:field_names]
                      pivot_table[:field_names].size
                    else
                      (pivot_table[:row_fields].size + pivot_table[:col_fields].size + pivot_table[:data_fields].size).clamp(1, 100)
                    end
      loc_attrs = %(<location ref="#{pivot_table[:dest_ref]}" firstHeaderRow="1" firstDataRow="1" firstDataCol="1")
      loc_attrs << %( rowPageCount="#{pivot_table[:row_page_count]}") if pivot_table[:row_page_count]
      loc_attrs << %( colPageCount="#{pivot_table[:col_page_count]}") if pivot_table[:col_page_count]
      loc_attrs << "/>"
      parts << loc_attrs
      parts << %(<pivotFields count="#{field_count}">)
      field_count.times do |fi|
        attrs = +""
        if pivot_table[:row_fields].include?(fi)
          attrs << ' axis="axisRow"'
        elsif pivot_table[:col_fields].include?(fi)
          attrs << ' axis="axisCol"'
        end
        attrs << ' dataField="1"' if pivot_table[:data_fields].any? { |df| df[:fld] == fi }
        fa = pivot_table[:field_attrs] && pivot_table[:field_attrs][fi]
        attrs << %( compact="#{fa[:compact] ? "1" : "0"}") if fa && !fa[:compact].nil?
        attrs << %( outline="#{fa[:outline] ? "1" : "0"}") if fa && !fa[:outline].nil?
        attrs << %( subtotalTop="#{fa[:subtotal_top] ? "1" : "0"}") if fa && !fa[:subtotal_top].nil?
        attrs << %( showAll="#{fa && fa[:show_all] == true ? "1" : "0"}")
        attrs << %( numFmtId="#{fa[:num_fmt_id]}") if fa && fa[:num_fmt_id]
        attrs << %( sortType="#{xml_escape(fa[:sort_type])}") if fa && fa[:sort_type]
        attrs << ' defaultSubtotal="0"' if fa && fa[:default_subtotal] == false
        attrs << ' insertBlankRow="1"' if fa && fa[:insert_blank_row]
        attrs << ' insertPageBreak="1"' if fa && fa[:insert_page_break]
        attrs << ' includeNewItemsInFilter="1"' if fa && fa[:include_new_items_in_filter]

        field_items = pivot_table[:items] && pivot_table[:items][fi]
        if field_items
          parts << "<pivotField#{attrs}>"
          parts << %(<items count="#{field_items.size + 1}">)
          field_items.size.times { |ix| parts << %(<item x="#{ix}"/>) }
          parts << '<item t="default"/>'
          parts << "</items>"
          parts << "</pivotField>"
        else
          parts << "<pivotField#{attrs}/>"
        end
      end
      parts << "</pivotFields>"

      unless pivot_table[:row_fields].empty?
        parts << %(<rowFields count="#{pivot_table[:row_fields].size}">)
        pivot_table[:row_fields].each { |f| parts << %(<field x="#{f}"/>) }
        parts << "</rowFields>"
      end

      unless pivot_table[:col_fields].empty?
        parts << %(<colFields count="#{pivot_table[:col_fields].size}">)
        pivot_table[:col_fields].each { |f| parts << %(<field x="#{f}"/>) }
        parts << "</colFields>"
      end

      unless pivot_table[:data_fields].empty?
        parts << %(<dataFields count="#{pivot_table[:data_fields].size}">)
        pivot_table[:data_fields].each do |df|
          df_attrs = %( name="#{xml_escape(df[:name])}" fld="#{df[:fld]}" subtotal="#{df[:subtotal] || "sum"}")
          df_attrs << %( showDataAs="#{xml_escape(df[:show_data_as])}") if df[:show_data_as]
          df_attrs << %( baseField="#{df[:base_field]}") if df[:base_field]
          df_attrs << %( baseItem="#{df[:base_item]}") if df[:base_item]
          df_attrs << %( numFmtId="#{df[:num_fmt_id]}") if df[:num_fmt_id]
          parts << "<dataField#{df_attrs}/>"
        end
        parts << "</dataFields>"
      end

      if pivot_table[:pivot_table_style]
        psi = pivot_table[:pivot_table_style]
        psi_attrs = +""
        psi_attrs << %( name="#{xml_escape(psi[:name])}") if psi[:name]
        psi_attrs << %( showRowHeaders="#{psi[:show_row_headers] ? "1" : "0"}") unless psi[:show_row_headers].nil?
        psi_attrs << %( showColHeaders="#{psi[:show_col_headers] ? "1" : "0"}") unless psi[:show_col_headers].nil?
        psi_attrs << %( showRowStripes="#{psi[:show_row_stripes] ? "1" : "0"}") unless psi[:show_row_stripes].nil?
        psi_attrs << %( showColStripes="#{psi[:show_col_stripes] ? "1" : "0"}") unless psi[:show_col_stripes].nil?
        psi_attrs << %( showLastColumn="#{psi[:show_last_column] ? "1" : "0"}") unless psi[:show_last_column].nil?
        parts << "<pivotTableStyleInfo#{psi_attrs}/>"
      end

      parts << "</pivotTableDefinition>"
      parts.join
    end

    def generate_pivot_cache_definition_xml(pivot_table, _cache_id)
      pcd_attrs = %( xmlns="#{SSML_NS}" xmlns:r="#{DOC_REL_NS}" r:id="rId1" refreshOnLoad="1")
      pcd_attrs << ' saveData="0"' if pivot_table[:cache_save_data] == false
      pcd_attrs << ' enableRefresh="0"' if pivot_table[:cache_enable_refresh] == false
      pcd_attrs << %( refreshedBy="#{xml_escape(pivot_table[:cache_refreshed_by])}") if pivot_table[:cache_refreshed_by]
      pcd_attrs << %( refreshedVersion="#{pivot_table[:cache_refreshed_version]}") if pivot_table[:cache_refreshed_version]
      pcd_attrs << %( createdVersion="#{pivot_table[:cache_created_version]}") if pivot_table[:cache_created_version]
      pcd_attrs << %( recordCount="#{pivot_table[:cache_record_count]}") if pivot_table[:cache_record_count]
      pcd_attrs << ' optimizeMemory="1"' if pivot_table[:cache_optimize_memory]
      parts = [
        XML_HEADER,
        "<pivotCacheDefinition#{pcd_attrs}>"
      ]

      # Parse source ref: "Sheet1!A1:C4" => sheet name + range.
      source = pivot_table[:source_ref]
      ws_name_attr = pivot_table[:source_name] ? %( name="#{xml_escape(pivot_table[:source_name])}") : ""
      if source.include?("!")
        sname, srange = source.split("!", 2)
        sname = sname.delete("'")
        parts << %(<cacheSource type="worksheet"><worksheetSource ref="#{srange}" sheet="#{xml_escape(sname)}"#{ws_name_attr}/></cacheSource>)
      else
        parts << %(<cacheSource type="worksheet"><worksheetSource ref="#{source}"#{ws_name_attr}/></cacheSource>)
      end

      field_count = pivot_table[:field_names] ? pivot_table[:field_names].size : (pivot_table[:row_fields].size + pivot_table[:col_fields].size + pivot_table[:data_fields].size)
      parts << %(<cacheFields count="#{field_count}">)
      field_count.times do |fi|
        fname = if pivot_table[:field_names] && pivot_table[:field_names][fi]
                  pivot_table[:field_names][fi]
                else
                  df = pivot_table[:data_fields].find { |d| d[:fld] == fi }
                  df ? df[:name] : "Field#{fi + 1}"
                end
        fa = pivot_table[:field_attrs] && pivot_table[:field_attrs][fi]
        cf_num_fmt = (fa && fa[:cache_num_fmt_id]) || 0
        cf_attrs = %( name="#{xml_escape(fname)}" numFmtId="#{cf_num_fmt}")
        cf_attrs << %( caption="#{xml_escape(fa[:cache_caption])}") if fa && fa[:cache_caption]
        cf_attrs << %( formula="#{xml_escape(fa[:cache_formula])}") if fa && fa[:cache_formula]
        field_items = pivot_table[:items] && pivot_table[:items][fi]
        if field_items
          parts << "<cacheField#{cf_attrs}>"
          parts << %(<sharedItems count="#{field_items.size}">)
          field_items.each { |v| parts << %(<s v="#{xml_escape(v.to_s)}"/>) }
          parts << "</sharedItems>"
          parts << "</cacheField>"
        else
          parts << "<cacheField#{cf_attrs}><sharedItems/></cacheField>"
        end
      end
      parts << "</cacheFields>"
      parts << "</pivotCacheDefinition>"
      parts.join
    end

    def generate_pivot_cache_records_xml(pivot_table)
      items = pivot_table[:items]
      if items&.values&.any? { |v| v && !v.empty? }
        max_len = items.values.map { |v| v ? v.size : 0 }.max
        parts = [XML_HEADER, %(<pivotCacheRecords xmlns="#{SSML_NS}" count="#{max_len}">)]
        max_len.times do |ri|
          parts << "<r>"
          field_count = pivot_table[:field_names] ? pivot_table[:field_names].size : (pivot_table[:row_fields].size + pivot_table[:col_fields].size + pivot_table[:data_fields].size)
          field_count.times do |fi|
            field_items = items[fi]
            parts << if field_items
                       %(<x v="#{ri < field_items.size ? ri : 0}"/>)
                     else
                       %(<n v="0"/>)
                     end
          end
          parts << "</r>"
        end
        parts << "</pivotCacheRecords>"
        parts.join
      else
        [XML_HEADER, %(<pivotCacheRecords xmlns="#{SSML_NS}" count="0"/>)].join
      end
    end

    def generate_pivot_cache_rels(cache_id)
      [
        XML_HEADER,
        %(<Relationships xmlns="#{REL_NS}">),
        %(<Relationship Id="rId1" Type="#{DOC_REL_NS}/pivotCacheRecords" Target="pivotCacheRecords#{cache_id}.xml"/>),
        "</Relationships>"
      ].join
    end

    def generate_pivot_table_rels(cache_id)
      [
        XML_HEADER,
        %(<Relationships xmlns="#{REL_NS}">),
        %(<Relationship Id="rId1" Type="#{DOC_REL_NS}/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition#{cache_id}.xml"/>),
        "</Relationships>"
      ].join
    end

    def generate_external_link_xml(ext_link)
      parts = [
        XML_HEADER,
        %(<externalLink xmlns="#{SSML_NS}" xmlns:r="#{DOC_REL_NS}">),
        '<externalBook r:id="rId1">'
      ]
      unless ext_link[:sheet_names].empty?
        parts << "<sheetNames>"
        ext_link[:sheet_names].each { |sn| parts << %(<sheetName val="#{xml_escape(sn)}"/>) }
        parts << "</sheetNames>"
      end
      parts << "</externalBook>"
      parts << "</externalLink>"
      parts.join
    end

    def generate_external_link_rels(ext_link)
      [
        XML_HEADER,
        %(<Relationships xmlns="#{REL_NS}">),
        %(<Relationship Id="rId1" Type="#{DOC_REL_NS}/externalLinkPath" Target="#{xml_escape(ext_link[:target])}" TargetMode="External"/>),
        "</Relationships>"
      ].join
    end

    def parse_extra_content_types(ct_xml)
      ct_xml.scan(/<Default\s+Extension="([^"]+)"\s+ContentType="([^"]+)"/).each do |ext, ct|
        @extra_ct_defaults[ext] ||= ct
      end
      ct_xml.scan(/<Override\s+PartName="([^"]+)"\s+ContentType="([^"]+)"/).each do |pn, ct|
        @extra_ct_overrides[pn] ||= ct
      end
    end

    def xml_escape(value)
      value.to_s
           .gsub("&", "&amp;")
           .gsub("<", "&lt;")
           .gsub(">", "&gt;")
           .gsub('"', "&quot;")
           .gsub("'", "&apos;")
    end

    def rich_text_xml(rich_text)
      rich_text.runs.map do |run|
        font = run[:font]
        if font && !font.empty?
          rpr = +""
          rpr << "<b/>" if font[:bold]
          rpr << "<i/>" if font[:italic]
          rpr << "<strike/>" if font[:strike]
          if font[:underline]
            rpr << if font[:underline] == true
                     "<u/>"
                   else
                     %(<u val="#{font[:underline]}"/>)
                   end
          end
          rpr << %(<vertAlign val="#{font[:vert_align]}"/>) if font[:vert_align]
          rpr << %(<sz val="#{font[:sz]}"/>) if font[:sz]
          rpr << emit_color_xml(font)
          rpr << %(<rFont val="#{xml_escape(font[:name])}"/>) if font[:name]
          rpr << %(<family val="#{font[:family]}"/>) if font[:family]
          rpr << %(<scheme val="#{font[:scheme]}"/>) if font[:scheme]
          "<r><rPr>#{rpr}</rPr><t>#{xml_escape(run[:text])}</t></r>"
        else
          "<r><t>#{xml_escape(run[:text])}</t></r>"
        end
      end.join
    end

    def cell_xml(cell_ref, value, style_idx, sst = nil, ph: nil) # rubocop:disable Naming/MethodParameterName
      s_attr = style_idx ? %( s="#{style_idx}") : ""
      ph_attr = ph ? ' ph="1"' : ""
      case value
      when CellError
        %(<c r="#{cell_ref}" t="e"#{s_attr}#{ph_attr}><v>#{xml_escape(value.code)}</v></c>)
      when Formula
        f_attrs = +""
        case value.type
        when :shared
          f_attrs << %( t="shared" si="#{value.shared_index}")
          f_attrs << %( ref="#{value.ref}") if value.ref
        when :array
          f_attrs << %( t="array" ref="#{value.ref}") if value.ref
        when :data_table
          f_attrs << ' t="dataTable"'
          f_attrs << ' dt2D="1"' if value.dt2d
          f_attrs << ' dtr="1"' if value.dtr
          f_attrs << %( r1="#{value.r1}") if value.r1
          f_attrs << %( r2="#{value.r2}") if value.r2
        end
        f_attrs << ' ca="1"' if value.calculate_always
        f_attrs << ' aca="1"' if value.aca
        f_attrs << ' bx="1"' if value.bx
        parts = %(<c r="#{cell_ref}"#{s_attr}#{ph_attr}><f#{f_attrs}>#{xml_escape(value.expression)}</f>)
        parts << "<v>#{xml_escape(value.cached_value.to_s)}</v>" unless value.cached_value.nil?
        parts << "</c>"
        parts
      when RichText
        if sst
          rt_sst, = sst
          idx = rt_sst[value]
          %(<c r="#{cell_ref}" t="s"#{s_attr}#{ph_attr}><v>#{idx}</v></c>)
        else
          %(<c r="#{cell_ref}" t="inlineStr"#{s_attr}#{ph_attr}><is>#{rich_text_xml(value)}</is></c>)
        end
      when true, false
        %(<c r="#{cell_ref}" t="b"#{s_attr}#{ph_attr}><v>#{value ? 1 : 0}</v></c>)
      when Time
        serial = Xlsxrb.datetime_to_serial(value)
        dt_style = resolve_style_index(datetime_num_fmt_id)
        dt_attr = dt_style ? %( s="#{dt_style}") : ""
        %(<c r="#{cell_ref}"#{dt_attr}#{ph_attr}><v>#{serial}</v></c>)
      when Date
        serial = Xlsxrb.date_to_serial(value)
        date_style = resolve_style_index(date_num_fmt_id)
        ds_attr = date_style ? %( s="#{date_style}") : ""
        %(<c r="#{cell_ref}"#{ds_attr}#{ph_attr}><v>#{serial}</v></c>)
      when Numeric
        %(<c r="#{cell_ref}"#{s_attr}#{ph_attr}><v>#{value}</v></c>)
      else
        if sst
          _, str_sst = sst
          idx = str_sst[value.to_s]
          %(<c r="#{cell_ref}" t="s"#{s_attr}#{ph_attr}><v>#{idx}</v></c>)
        else
          %(<c r="#{cell_ref}" t="inlineStr"#{s_attr}#{ph_attr}><is><t>#{xml_escape(value)}</t></is></c>)
        end
      end
    end

    # Returns the numFmtId for dates, registering it on first use.
    def date_num_fmt_id
      @date_num_fmt_id ||= add_number_format(DEFAULT_DATE_FORMAT)
    end

    # Returns the numFmtId for datetime, registering it on first use.
    def datetime_num_fmt_id
      @datetime_num_fmt_id ||= add_number_format(DEFAULT_DATETIME_FORMAT)
    end

    # Maps a numFmtId to a cellXfs index. Index 0 is the default (no format).
    def resolve_style_index(style_value)
      return nil if style_value.nil?

      # New-style: { xf_index: N } from set_cell_style.
      return style_value[:xf_index] if style_value.is_a?(Hash) && style_value.key?(:xf_index)

      # Legacy: raw num_fmt_id from set_cell_format — find or create matching xf entry.
      num_fmt_id = style_value
      @xf_index_map ||= begin
        map = {}
        @num_fmts.each_with_index do |nf, _i|
          entry = { num_fmt_id: nf[:num_fmt_id], font_id: 0, fill_id: 0, border_id: 0 }
          idx = @xf_entries.index(entry)
          unless idx
            @xf_entries << entry
            idx = @xf_entries.size - 1
          end
          map[nf[:num_fmt_id]] = idx
        end
        map
      end
      @xf_index_map[num_fmt_id]
    end

    def emit_filter_xml(filter)
      case filter[:type]
      when :filters
        attrs = filter[:blank] ? +' blank="1"' : +""
        attrs << %( calendarType="#{filter[:calendar_type]}") if filter[:calendar_type]
        has_values = filter[:values]&.any?
        has_date_groups = filter[:date_group_items]&.any?
        if has_values || has_date_groups
          parts = ["<filters#{attrs}>"]
          filter[:values]&.each { |v| parts << %(<filter val="#{xml_escape(v)}"/>) }
          filter[:date_group_items]&.each do |dg|
            dg_attrs = %(dateTimeGrouping="#{dg[:date_time_grouping]}")
            dg_attrs << %( year="#{dg[:year]}") if dg[:year]
            dg_attrs << %( month="#{dg[:month]}") if dg[:month]
            dg_attrs << %( day="#{dg[:day]}") if dg[:day]
            dg_attrs << %( hour="#{dg[:hour]}") if dg[:hour]
            dg_attrs << %( minute="#{dg[:minute]}") if dg[:minute]
            dg_attrs << %( second="#{dg[:second]}") if dg[:second]
            parts << "<dateGroupItem #{dg_attrs}/>"
          end
          parts << "</filters>"
          parts.join
        else
          "<filters#{attrs}/>"
        end
      when :custom
        if filter[:filters]
          and_attr = filter[:and] ? ' and="1"' : ""
          parts = ["<customFilters#{and_attr}>"]
          filter[:filters].each do |cf|
            parts << %(<customFilter operator="#{cf[:operator]}" val="#{xml_escape(cf[:val])}"/>)
          end
          parts << "</customFilters>"
          parts.join
        else
          %(<customFilters><customFilter operator="#{filter[:operator]}" val="#{xml_escape(filter[:val])}"/></customFilters>)
        end
      when :dynamic
        dyn_attrs = %( type="#{filter[:dynamic_type]}")
        dyn_attrs << %( val="#{filter[:val]}") if filter[:val]
        dyn_attrs << %( valIso="#{filter[:val_iso]}") if filter[:val_iso]
        dyn_attrs << %( maxVal="#{filter[:max_val]}") if filter[:max_val]
        dyn_attrs << %( maxValIso="#{filter[:max_val_iso]}") if filter[:max_val_iso]
        "<dynamicFilter#{dyn_attrs}/>"
      when :top10
        top_attr = filter[:top] ? ' top="1"' : ""
        pct_attr = filter[:percent] ? ' percent="1"' : ""
        fv_attr = filter[:filter_val] ? %( filterVal="#{filter[:filter_val]}") : ""
        %(<top10#{top_attr}#{pct_attr} val="#{filter[:val]}"#{fv_attr}/>)
      when :color_filter
        cf_attrs = %(dxfId="#{filter[:dxf_id]}")
        cf_attrs << ' cellColor="0"' if filter[:cell_color] == false
        %(<colorFilter #{cf_attrs}/>)
      when :icon_filter
        if_attrs = %(iconSet="#{filter[:icon_set]}")
        if_attrs << %( iconId="#{filter[:icon_id]}") if filter[:icon_id]
        %(<iconFilter #{if_attrs}/>)
      else
        ""
      end
    end

    def emit_cf_rule(parts, rule)
      type = rule[:type]
      rule_type = CF_TYPE_MAP[type] || type.to_s
      rule_attrs = %(type="#{rule_type}")
      rule_attrs << %( priority="#{rule[:priority]}") if rule[:priority]
      rule_attrs << %( operator="#{rule[:operator]}") if rule[:operator]
      rule_attrs << %( dxfId="#{rule[:format_id]}") if rule[:format_id]
      rule_attrs << %( stopIfTrue="1") if rule[:stop_if_true]
      rule_attrs << %( aboveAverage="0") if rule[:above_average] == false
      rule_attrs << %( equalAverage="1") if rule[:equal_average]
      rule_attrs << %( rank="#{rule[:rank]}") if rule[:rank]
      rule_attrs << %( percent="1") if rule[:percent]
      rule_attrs << %( bottom="1") if rule[:bottom]
      rule_attrs << %( text="#{xml_escape(rule[:text])}") if rule[:text]
      rule_attrs << %( timePeriod="#{rule[:time_period]}") if rule[:time_period]
      rule_attrs << %( stdDev="#{rule[:std_dev]}") if rule[:std_dev]

      case type
      when :cell_is, :expression, :above_average, :top10, :duplicate_values, :unique_values,
           :contains_text, :not_contains_text, :begins_with, :ends_with,
           :contains_blanks, :not_contains_blanks, :time_period
        formulas = rule[:formulas] || [rule[:formula]].compact
        if formulas.empty?
          parts << "<cfRule #{rule_attrs}/>"
        else
          parts << "<cfRule #{rule_attrs}>"
          formulas.each { |f| parts << "<formula>#{xml_escape(f)}</formula>" }
          parts << "</cfRule>"
        end
      when :color_scale
        cs = rule[:color_scale]
        parts << "<cfRule #{rule_attrs}>"
        parts << "<colorScale>"
        cs[:cfvo]&.each do |cfvo|
          cfvo_attrs = %(type="#{cfvo[:type]}")
          cfvo_attrs << %( val="#{cfvo[:val]}") if cfvo[:val]
          cfvo_attrs << ' gte="0"' if cfvo[:gte] == false
          parts << "<cfvo #{cfvo_attrs}/>"
        end
        cs[:colors]&.each { |c| parts << emit_cf_color_xml(c) }
        parts << "</colorScale>"
        parts << "</cfRule>"
      when :data_bar
        db = rule[:data_bar]
        parts << "<cfRule #{rule_attrs}>"
        db_attrs = +""
        db_attrs << %( minLength="#{db[:min_length]}") if db[:min_length]
        db_attrs << %( maxLength="#{db[:max_length]}") if db[:max_length]
        db_attrs << %( showValue="#{db[:show_value] ? 1 : 0}") unless db[:show_value].nil?
        parts << "<dataBar#{db_attrs}>"
        db[:cfvo]&.each do |cfvo|
          cfvo_attrs = %(type="#{cfvo[:type]}")
          cfvo_attrs << %( val="#{cfvo[:val]}") if cfvo[:val]
          cfvo_attrs << ' gte="0"' if cfvo[:gte] == false
          parts << "<cfvo #{cfvo_attrs}/>"
        end
        parts << emit_cf_color_xml(db[:color]) if db[:color]
        parts << "</dataBar>"
        parts << "</cfRule>"
      when :icon_set
        is = rule[:icon_set]
        parts << "<cfRule #{rule_attrs}>"
        is_attrs = +""
        is_attrs << %( iconSet="#{is[:icon_set]}") if is[:icon_set]
        is_attrs << %( reverse="#{is[:reverse] ? 1 : 0}") unless is[:reverse].nil?
        is_attrs << %( percent="#{is[:percent] ? 1 : 0}") unless is[:percent].nil?
        is_attrs << %( showValue="#{is[:show_value] ? 1 : 0}") unless is[:show_value].nil?
        parts << "<iconSet#{is_attrs}>"
        is[:cfvo]&.each do |cfvo|
          cfvo_attrs = %(type="#{cfvo[:type]}")
          cfvo_attrs << %( val="#{cfvo[:val]}") if cfvo[:val]
          cfvo_attrs << ' gte="0"' if cfvo[:gte] == false
          parts << "<cfvo #{cfvo_attrs}/>"
        end
        parts << "</iconSet>"
        parts << "</cfRule>"
      else
        parts << "<cfRule #{rule_attrs}/>"
      end
    end

    # Emits a <color> element for CF rules, accepting either a plain RGB string or a hash.
    def emit_cf_color_xml(color)
      if color.is_a?(Hash)
        emit_color_xml(color)
      else
        %(<color rgb="#{color}"/>)
      end
    end

    def compute_dimension(sheet_cells)
      return "A1" if sheet_cells.empty?

      min_col = Float::INFINITY
      max_col = 0
      min_row = Float::INFINITY
      max_row = 0
      sheet_cells.each_key do |addr|
        col_letter = extract_column_letter(addr)
        row_num = extract_row_number(addr)
        col_idx = column_letter_to_index(col_letter)
        min_col = col_idx if col_idx < min_col
        max_col = col_idx if col_idx > max_col
        min_row = row_num if row_num < min_row
        max_row = row_num if row_num > max_row
      end
      start_col = index_to_column_letter(min_col)
      end_col = index_to_column_letter(max_col)
      "#{start_col}#{min_row}:#{end_col}#{max_row}"
    end

    def index_to_column_letter(index)
      result = +""
      while index.positive?
        index -= 1
        result.prepend(("A".ord + (index % 26)).chr)
        index /= 26
      end
      result
    end

    def generate_core_properties_xml
      parts = [
        XML_HEADER,
        %(<cp:coreProperties xmlns:cp="#{CP_NS}" xmlns:dc="#{DC_NS}" xmlns:dcterms="#{DCTERMS_NS}" xmlns:xsi="#{XSI_NS}">)
      ]
      parts << "<dc:title>#{xml_escape(@core_properties[:title])}</dc:title>" if @core_properties[:title]
      parts << "<dc:subject>#{xml_escape(@core_properties[:subject])}</dc:subject>" if @core_properties[:subject]
      parts << "<dc:creator>#{xml_escape(@core_properties[:creator])}</dc:creator>" if @core_properties[:creator]
      parts << "<cp:keywords>#{xml_escape(@core_properties[:keywords])}</cp:keywords>" if @core_properties[:keywords]
      parts << "<dc:description>#{xml_escape(@core_properties[:description])}</dc:description>" if @core_properties[:description]
      parts << "<cp:lastModifiedBy>#{xml_escape(@core_properties[:last_modified_by])}</cp:lastModifiedBy>" if @core_properties[:last_modified_by]
      parts << "<cp:revision>#{xml_escape(@core_properties[:revision])}</cp:revision>" if @core_properties[:revision]
      parts << %(<dcterms:created xsi:type="dcterms:W3CDTF">#{xml_escape(@core_properties[:created])}</dcterms:created>) if @core_properties[:created]
      parts << %(<dcterms:modified xsi:type="dcterms:W3CDTF">#{xml_escape(@core_properties[:modified])}</dcterms:modified>) if @core_properties[:modified]
      parts << "<cp:category>#{xml_escape(@core_properties[:category])}</cp:category>" if @core_properties[:category]
      parts << "<cp:contentStatus>#{xml_escape(@core_properties[:content_status])}</cp:contentStatus>" if @core_properties[:content_status]
      parts << "<dc:language>#{xml_escape(@core_properties[:language])}</dc:language>" if @core_properties[:language]
      parts << "</cp:coreProperties>"
      parts.join
    end

    def generate_app_properties_xml
      parts = [
        XML_HEADER,
        %(<Properties xmlns="#{APP_NS}" xmlns:vt="#{VT_NS}">)
      ]
      parts << "<Application>#{xml_escape(@app_properties[:application])}</Application>" if @app_properties[:application]
      parts << "<AppVersion>#{xml_escape(@app_properties[:app_version])}</AppVersion>" if @app_properties[:app_version]
      if @app_properties[:heading_pairs] && @app_properties[:titles_of_parts]
        hp = @app_properties[:heading_pairs]
        tp = @app_properties[:titles_of_parts]
        parts << "<HeadingPairs>"
        parts << %(<vt:vector size="#{hp.size * 2}" baseType="variant">)
        hp.each do |label, count|
          parts << "<vt:variant><vt:lpstr>#{xml_escape(label)}</vt:lpstr></vt:variant>"
          parts << "<vt:variant><vt:i4>#{count}</vt:i4></vt:variant>"
        end
        parts << "</vt:vector>"
        parts << "</HeadingPairs>"
        parts << "<TitlesOfParts>"
        parts << %(<vt:vector size="#{tp.size}" baseType="lpstr">)
        tp.each { |t| parts << "<vt:lpstr>#{xml_escape(t)}</vt:lpstr>" }
        parts << "</vt:vector>"
        parts << "</TitlesOfParts>"
      end
      parts << "</Properties>"
      parts.join
    end

    def generate_custom_properties_xml
      custom_ns = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
      parts = [
        XML_HEADER,
        %(<Properties xmlns="#{custom_ns}" xmlns:vt="#{VT_NS}">)
      ]
      @custom_properties.each_with_index do |prop, idx|
        fmtid = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
        pid = idx + 2 # pids start at 2
        parts << %(<property fmtid="#{fmtid}" pid="#{pid}" name="#{xml_escape(prop[:name])}">)
        parts << case prop[:type]
                 when :number, :int, :i4
                   "<vt:i4>#{prop[:value]}</vt:i4>"
                 when :float, :r8
                   "<vt:r8>#{prop[:value]}</vt:r8>"
                 when :bool
                   "<vt:bool>#{prop[:value] ? "true" : "false"}</vt:bool>"
                 when :date, :filetime
                   "<vt:filetime>#{prop[:value]}</vt:filetime>"
                 else
                   "<vt:lpwstr>#{xml_escape(prop[:value].to_s)}</vt:lpwstr>"
                 end
        parts << "</property>"
      end
      parts << "</Properties>"
      parts.join
    end

    def generate_styles_xml
      parts = [
        XML_HEADER,
        %(<styleSheet xmlns="#{SSML_NS}">)
      ]

      # numFmts
      unless @num_fmts.empty?
        parts << %(<numFmts count="#{@num_fmts.size}">)
        @num_fmts.each do |nf|
          parts << %(<numFmt numFmtId="#{nf[:num_fmt_id]}" formatCode="#{xml_escape(nf[:format_code])}"/>)
        end
        parts << "</numFmts>"
      end

      # fonts
      parts << %(<fonts count="#{@fonts.size}">)
      @fonts.each { |f| parts << emit_font_xml(f) }
      parts << "</fonts>"

      # fills
      parts << %(<fills count="#{@fills.size}">)
      @fills.each { |f| parts << emit_fill_xml(f) }
      parts << "</fills>"

      # borders
      parts << %(<borders count="#{@borders.size}">)
      @borders.each { |b| parts << emit_border_xml(b) }
      parts << "</borders>"

      # cellStyleXfs
      parts << %(<cellStyleXfs count="#{@cell_style_xfs.size}">)
      @cell_style_xfs.each do |xf|
        parts << %(<xf numFmtId="#{xf[:num_fmt_id]}" fontId="#{xf[:font_id]}" fillId="#{xf[:fill_id]}" borderId="#{xf[:border_id]}"/>)
      end
      parts << "</cellStyleXfs>"

      # cellXfs
      parts << %(<cellXfs count="#{@xf_entries.size}">)
      @xf_entries.each do |xf|
        apply_attrs = []
        apply_attrs << ' applyNumberFormat="1"' if xf[:num_fmt_id].positive?
        apply_attrs << ' applyFont="1"' if xf[:font_id].positive?
        apply_attrs << ' applyFill="1"' if xf[:fill_id].positive?
        apply_attrs << ' applyBorder="1"' if xf[:border_id].positive?
        apply_attrs << ' applyAlignment="1"' if xf[:alignment]
        apply_attrs << ' applyProtection="1"' if xf[:protection]
        apply_attrs << ' quotePrefix="1"' if xf[:quote_prefix]
        apply_attrs << ' pivotButton="1"' if xf[:pivot_button]
        children = []
        children << emit_alignment_xml(xf[:alignment]) if xf[:alignment]
        children << emit_protection_xml(xf[:protection]) if xf[:protection]
        xf_id = xf[:xf_id] || 0
        parts << if children.any?
                   %(<xf numFmtId="#{xf[:num_fmt_id]}" fontId="#{xf[:font_id]}" fillId="#{xf[:fill_id]}" borderId="#{xf[:border_id]}" xfId="#{xf_id}"#{apply_attrs.join}>#{children.join}</xf>)
                 else
                   %(<xf numFmtId="#{xf[:num_fmt_id]}" fontId="#{xf[:font_id]}" fillId="#{xf[:fill_id]}" borderId="#{xf[:border_id]}" xfId="#{xf_id}"#{apply_attrs.join}/>)
                 end
      end
      parts << "</cellXfs>"

      # cellStyles
      parts << %(<cellStyles count="#{@cell_style_names.size}">)
      @cell_style_names.each do |cs|
        cs_attrs = %(name="#{xml_escape(cs[:name])}" xfId="#{cs[:xf_id]}")
        cs_attrs << %( builtinId="#{cs[:builtin_id]}") if cs[:builtin_id]
        cs_attrs << %( iLevel="#{cs[:i_level]}") if cs[:i_level]
        cs_attrs << ' hidden="1"' if cs[:hidden]
        cs_attrs << ' customBuiltin="1"' if cs[:custom_builtin]
        parts << "<cellStyle #{cs_attrs}/>"
      end
      parts << "</cellStyles>"

      # dxfs
      unless @dxfs.empty?
        parts << %(<dxfs count="#{@dxfs.size}">)
        @dxfs.each { |d| parts << emit_dxf_xml(d) }
        parts << "</dxfs>"
      end

      # tableStyles
      ts_styles = @table_styles[:styles] || []
      unless ts_styles.empty? && @table_styles[:default_table_style].nil? && @table_styles[:default_pivot_style].nil?
        ts_attrs = [%(count="#{ts_styles.size}")]
        ts_attrs << %(defaultTableStyle="#{xml_escape(@table_styles[:default_table_style])}") if @table_styles[:default_table_style]
        ts_attrs << %(defaultPivotStyle="#{xml_escape(@table_styles[:default_pivot_style])}") if @table_styles[:default_pivot_style]
        if ts_styles.empty?
          parts << "<tableStyles #{ts_attrs.join(" ")}/>"
        else
          parts << "<tableStyles #{ts_attrs.join(" ")}>"
          ts_styles.each do |ts|
            s_attrs = [%(name="#{xml_escape(ts[:name])}")]
            s_attrs << %(pivot="0") if ts[:pivot] == false
            s_attrs << %(table="0") if ts[:table] == false
            s_attrs << %(count="#{ts[:elements].size}") unless ts[:elements].empty?
            if ts[:elements].empty?
              parts << "<tableStyle #{s_attrs.join(" ")}/>"
            else
              parts << "<tableStyle #{s_attrs.join(" ")}>"
              ts[:elements].each do |el|
                el_attrs = [%(type="#{el[:type]}")]
                el_attrs << %(size="#{el[:size]}") if el[:size] && el[:size] != 1
                el_attrs << %(dxfId="#{el[:dxf_id]}") if el[:dxf_id]
                parts << "<tableStyleElement #{el_attrs.join(" ")}/>"
              end
              parts << "</tableStyle>"
            end
          end
          parts << "</tableStyles>"
        end
      end

      # colors
      unless @indexed_colors.empty? && @mru_colors.empty?
        parts << "<colors>"
        unless @indexed_colors.empty?
          parts << "<indexedColors>"
          @indexed_colors.each { |c| parts << %(<rgbColor rgb="#{c}"/>) }
          parts << "</indexedColors>"
        end
        unless @mru_colors.empty?
          parts << "<mruColors>"
          @mru_colors.each { |c| parts << emit_color_xml(c) }
          parts << "</mruColors>"
        end
        parts << "</colors>"
      end

      parts << "</styleSheet>"
      parts.join
    end

    def emit_alignment_xml(alignment)
      attrs = []
      attrs << %(horizontal="#{alignment[:horizontal]}") if alignment[:horizontal]
      attrs << %(vertical="#{alignment[:vertical]}") if alignment[:vertical]
      attrs << %(wrapText="1") if alignment[:wrap_text]
      attrs << %(textRotation="#{alignment[:text_rotation]}") if alignment[:text_rotation]
      attrs << %(indent="#{alignment[:indent]}") if alignment[:indent]
      attrs << %(relativeIndent="#{alignment[:relative_indent]}") if alignment[:relative_indent]
      attrs << %(shrinkToFit="1") if alignment[:shrink_to_fit]
      attrs << %(readingOrder="#{alignment[:reading_order]}") if alignment[:reading_order]
      attrs << %(justifyLastLine="1") if alignment[:justify_last_line]
      "<alignment #{attrs.join(" ")}/>"
    end

    def emit_protection_xml(protection)
      attrs = []
      attrs << %(locked="#{protection[:locked] ? "1" : "0"}") unless protection[:locked].nil?
      attrs << %(hidden="#{protection[:hidden] ? "1" : "0"}") unless protection[:hidden].nil?
      "<protection #{attrs.join(" ")}/>"
    end

    def emit_color_xml(source, tag: "color")
      if source[:auto]
        %(<#{tag} auto="1"/>)
      elsif source[:color] || source[:rgb]
        %(<#{tag} rgb="#{source[:color] || source[:rgb]}"/>)
      elsif source[:theme]
        attrs = [%(theme="#{source[:theme]}")]
        attrs << %(tint="#{source[:tint]}") if source[:tint]
        %(<#{tag} #{attrs.join(" ")}/>)
      elsif source[:indexed]
        %(<#{tag} indexed="#{source[:indexed]}"/>)
      else
        ""
      end
    end

    def emit_font_xml(font)
      parts = ["<font>"]
      parts << "<b/>" if font[:bold]
      parts << "<i/>" if font[:italic]
      parts << "<strike/>" if font[:strike]
      parts << "<shadow/>" if font[:shadow]
      parts << "<outline/>" if font[:outline]
      parts << "<condense/>" if font[:condense]
      parts << "<extend/>" if font[:extend]
      if font[:underline]
        parts << if font[:underline] == true
                   "<u/>"
                 else
                   %(<u val="#{font[:underline]}"/>)
                 end
      end
      parts << %(<vertAlign val="#{font[:vert_align]}"/>) if font[:vert_align]
      parts << %(<sz val="#{font[:sz]}"/>) if font[:sz]
      parts << emit_color_xml(font)
      parts << %(<name val="#{xml_escape(font[:name])}"/>) if font[:name]
      parts << %(<family val="#{font[:family]}"/>) if font[:family]
      parts << %(<charset val="#{font[:charset]}"/>) if font[:charset]
      parts << %(<scheme val="#{font[:scheme]}"/>) if font[:scheme]
      parts << "</font>"
      parts.join
    end

    def emit_fill_xml(fill)
      return emit_gradient_fill_xml(fill[:gradient]) if fill[:gradient]

      has_fg = fill[:fg_color] || fill[:fg_color_theme] || fill[:fg_color_indexed] || fill[:fg_color_auto]
      has_bg = fill[:bg_color] || fill[:bg_color_theme] || fill[:bg_color_indexed] || fill[:bg_color_auto]
      return "<fill><patternFill patternType=\"#{fill[:pattern]}\"/></fill>" if fill[:pattern] && !has_fg && !has_bg

      parts = ["<fill>"]
      pt = fill[:pattern] || "solid"
      parts << %(<patternFill patternType="#{pt}">)
      parts << emit_fill_color_xml("fgColor", fill, :fg)
      parts << emit_fill_color_xml("bgColor", fill, :bg)
      parts << "</patternFill>"
      parts << "</fill>"
      parts.join
    end

    def emit_fill_color_xml(tag, fill, prefix)
      if fill[:"#{prefix}_color"]
        %(<#{tag} rgb="#{fill[:"#{prefix}_color"]}"/>)
      elsif fill[:"#{prefix}_color_theme"]
        attrs = [%(theme="#{fill[:"#{prefix}_color_theme"]}")]
        attrs << %(tint="#{fill[:"#{prefix}_color_tint"]}") if fill[:"#{prefix}_color_tint"]
        %(<#{tag} #{attrs.join(" ")}/>)
      elsif fill[:"#{prefix}_color_indexed"]
        %(<#{tag} indexed="#{fill[:"#{prefix}_color_indexed"]}"/>)
      elsif fill[:"#{prefix}_color_auto"]
        %(<#{tag} auto="1"/>)
      else
        ""
      end
    end

    def emit_gradient_fill_xml(gradient)
      attrs = []
      attrs << %(type="#{gradient[:type]}") if gradient[:type]
      attrs << %(degree="#{gradient[:degree]}") if gradient[:degree]
      attrs << %(left="#{gradient[:left]}") if gradient[:left]
      attrs << %(right="#{gradient[:right]}") if gradient[:right]
      attrs << %(top="#{gradient[:top]}") if gradient[:top]
      attrs << %(bottom="#{gradient[:bottom]}") if gradient[:bottom]
      parts = ["<fill>"]
      parts << "<gradientFill#{" #{attrs.join(" ")}" unless attrs.empty?}"
      if gradient[:stops]&.any?
        parts[-1] = "#{parts[-1]}>"
        gradient[:stops].each do |stop|
          parts << %(<stop position="#{stop[:position]}">#{emit_color_xml(stop)}</stop>)
        end
        parts << "</gradientFill>"
      else
        parts[-1] = "#{parts[-1]}/>"
      end
      parts << "</fill>"
      parts.join
    end

    def emit_brk_xml(brk, default_max:)
      if brk.is_a?(Hash)
        attrs = %(id="#{brk[:id]}")
        attrs << %( min="#{brk[:min]}") if brk[:min]
        attrs << %( max="#{brk.fetch(:max, default_max)}")
        attrs << ' man="1"' if brk.fetch(:man, true)
        attrs << ' pt="1"' if brk[:pt]
        "<brk #{attrs}/>"
      else
        %(<brk id="#{brk}" max="#{default_max}" man="1"/>)
      end
    end

    def emit_border_xml(bdr)
      border_attrs = []
      border_attrs << ' diagonalUp="1"' if bdr[:diagonal_up]
      border_attrs << ' diagonalDown="1"' if bdr[:diagonal_down]
      border_attrs << ' outline="0"' if bdr[:outline] == false
      parts = ["<border#{border_attrs.join}>"]
      %i[left right top bottom diagonal vertical horizontal].each do |side|
        s = bdr[side]
        if s.is_a?(Hash)
          parts << %(<#{side} style="#{s[:style]}">)
          parts << emit_color_xml(s)
          parts << "</#{side}>"
        else
          parts << "<#{side}/>"
        end
      end
      parts << "</border>"
      parts.join
    end

    def emit_dxf_xml(dxf)
      parts = ["<dxf>"]
      parts << emit_font_xml(dxf[:font]) if dxf[:font]
      if dxf[:num_fmt]
        nf = dxf[:num_fmt]
        parts << %(<numFmt numFmtId="#{nf[:num_fmt_id]}" formatCode="#{xml_escape(nf[:format_code])}"/>)
      end
      parts << emit_fill_xml(dxf[:fill]) if dxf[:fill]
      parts << emit_alignment_xml(dxf[:alignment]) if dxf[:alignment]
      parts << emit_border_xml(dxf[:border]) if dxf[:border]
      parts << emit_protection_xml(dxf[:protection]) if dxf[:protection]
      parts << "</dxf>"
      parts.join
    end

    def validate_cell_address!(cell_address)
      raise ArgumentError, "cell address must be a String" unless cell_address.is_a?(String)

      match = cell_address.match(CELL_ADDRESS_PATTERN)
      raise ArgumentError, "invalid cell address: #{cell_address.inspect}" unless match

      row_num = match[2].to_i
      raise ArgumentError, "row out of range: #{row_num}" unless row_num.between?(1, MAX_ROW)

      col_index = column_letter_to_index(match[1])
      raise ArgumentError, "column out of range: #{match[1]}" unless col_index.between?(1, MAX_COLUMN_INDEX)
    end

    def column_letter_to_index(letters)
      letters.chars.reduce(0) { |sum, char| (sum * 26) + (char.ord - "A".ord + 1) }
    end

    # Extracts the column letter(s) from a cell address, e.g. "A" from "A1".
    def extract_column_letter(cell_address)
      cell_address.match(/^([A-Z]+)/)[1]
    end

    # Extracts the row number from a cell address, e.g. 1 from "A1".
    def extract_row_number(cell_address)
      cell_address.match(/(\d+)$/)[1].to_i
    end

    # Converts a range like "A1:D20" to absolute "$A$1:$D$20".
    def absolute_range(range)
      range.split(":").map { |part| absolute_cell(part) }.join(":")
    end

    # Converts "A1" to "$A$1".
    def absolute_cell(cell_ref)
      m = cell_ref.match(/\A([A-Z]+)(\d+)\z/)
      raise ArgumentError, "invalid cell reference: #{cell_ref}" unless m

      "$#{m[1]}$#{m[2]}"
    end
  end
end
