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
      @sheet_order = ["Sheet1"]
      @core_properties = {}
      @app_properties = {}
      @workbook_properties = {}
      @workbook_views = {}
      @calc_properties = {}
      @sheet_states = {}
      @defined_names = []
      @sheet_properties = { "Sheet1" => {} }
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
      @conditional_formats = { "Sheet1" => [] }
      @tables = { "Sheet1" => [] }
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
      @conditional_formats[name] = []
      @tables[name] = []
      @images[name] = []
      @charts_data[name] = []
      @shapes_data[name] = []
      @comments_data[name] = []
      @pivot_tables_data[name] = []
      @sheet_protection[name] = nil
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
    def add_hyperlink(cell_address, url, sheet: nil)
      validate_cell_address!(cell_address)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @hyperlinks.key?(sheet_name)

      @hyperlinks[sheet_name][cell_address] = url
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
    def set_sort_state(ref, sort_conditions, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @sort_state.key?(sheet_name)

      @sort_state[sheet_name] = { ref: ref, sort_conditions: sort_conditions }
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
        border_id: opts[:border_id] || 0
      }
      entry[:alignment] = opts[:alignment] if opts[:alignment]
      existing = @xf_entries.index(entry)
      return existing if existing

      @xf_entries << entry
      @xf_entries.size - 1
    end

    # Registers a base style definition (cellStyleXf) and a named cellStyle.
    # Returns the xfId for the new base style.
    def add_named_cell_style(name:, num_fmt_id: 0, font_id: 0, fill_id: 0, border_id: 0, builtin_id: nil)
      entry = { num_fmt_id: num_fmt_id, font_id: font_id, fill_id: fill_id, border_id: border_id }
      @cell_style_xfs << entry
      xf_id = @cell_style_xfs.size - 1
      cs = { name: name, xf_id: xf_id }
      cs[:builtin_id] = builtin_id if builtin_id
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

    # Registers a differential format (dxf) for conditional formatting. Returns dxf_id.
    # Opts: font (hash), fill (hash), border (hash), num_fmt (hash).
    def add_dxf(**opts)
      @dxfs << opts
      @dxfs.size - 1
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

      @freeze_panes[sheet_name] = { row: row, col: col }
    end

    # Returns freeze pane settings for the first (or given) sheet.
    def freeze_pane(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @freeze_panes[sheet_name]
    end

    # Sets the active cell selection.
    def set_selection(active_cell, sqref: nil, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @selections.key?(sheet_name)

      @selections[sheet_name] = { active_cell: active_cell, sqref: sqref || active_cell }
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
    def add_table(ref, columns:, name: nil, display_name: nil, sheet: nil, totals_row_count: 0, style: nil)
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
    def add_defined_name(name, value, sheet: nil, hidden: false)
      entry = { name: name, value: value, hidden: hidden }
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

    # Sets a workbook property (e.g. :date1904, :default_theme_version).
    def set_workbook_property(name, value)
      @workbook_properties[name] = value
    end

    # Returns workbook properties hash.
    def workbook_properties
      @workbook_properties.dup
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
      if content_type
        ext = File.extname(path).delete(".")
        if ext.empty? || path.include?("/")
          @extra_ct_overrides["/#{path}"] = content_type
        else
          @extra_ct_defaults[ext] = content_type
        end
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
    def insert_image(file_data, ext: "png", from_col: 0, from_row: 0, to_col: 5, to_row: 10, name: nil, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @images.key?(sheet_name)

      img_name = name || "Picture #{@images[sheet_name].size + 1}"
      @images[sheet_name] << {
        file_data: file_data, ext: ext, name: img_name,
        from_col: from_col, from_row: from_row,
        to_col: to_col, to_row: to_row
      }
    end

    # Returns images for the first (or given) sheet.
    def images(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @images[sheet_name] || []
    end

    # Adds a chart to the given sheet.
    # type: :bar, :line, :pie. title: chart title string.
    # data_ref: e.g. "Sheet1!$A$1:$B$4". cat_ref/val_ref for explicit series.
    def add_chart(type: :bar, title: nil, cat_ref: nil, val_ref: nil, series: nil, legend: nil, data_labels: nil, cat_axis_title: nil, val_axis_title: nil, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @charts_data.key?(sheet_name)

      chart = { type: type, title: title }
      if series
        chart[:series] = series
      else
        chart[:series] = [{ cat_ref: cat_ref, val_ref: val_ref }]
      end
      chart[:legend] = legend if legend
      chart[:data_labels] = data_labels if data_labels
      chart[:cat_axis_title] = cat_axis_title if cat_axis_title
      chart[:val_axis_title] = val_axis_title if val_axis_title
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
    def add_shape(preset: "rect", text: nil, name: nil, from_col: 0, from_row: 0, to_col: 5, to_row: 5, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @shapes_data.key?(sheet_name)

      shape_name = name || "Shape #{@shapes_data[sheet_name].size + 1}"
      @shapes_data[sheet_name] << {
        preset: preset, text: text, name: shape_name,
        from_col: from_col, from_row: from_row,
        to_col: to_col, to_row: to_row
      }
    end

    # Returns shape definitions for the first (or given) sheet.
    def shapes(sheet: nil)
      sheet_name = sheet || @sheet_order.first
      @shapes_data[sheet_name] || []
    end

    # Adds a comment on a cell.
    def add_comment(cell_address, text, author: "Author", sheet: nil)
      validate_cell_address!(cell_address)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @comments_data.key?(sheet_name)

      @comments_data[sheet_name] << { ref: cell_address, text: text, author: author }
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
    def add_pivot_table(source_ref, row_fields:, data_fields:, col_fields: [], dest_ref: "E1", name: nil, field_names: nil, items: nil, sheet: nil)
      sheet_name = sheet || @sheet_order.first
      raise ArgumentError, "unknown sheet: #{sheet_name}" unless @pivot_tables_data.key?(sheet_name)

      pt_name = name || "PivotTable#{@pivot_tables_data.values.flatten.size + 1}"
      @pivot_tables_data[sheet_name] << {
        name: pt_name, source_ref: source_ref,
        row_fields: row_fields, col_fields: col_fields,
        data_fields: data_fields, dest_ref: dest_ref,
        field_names: field_names, items: items
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
      # Pre-register date format if any sheet contains Date values.
      @sheet_order.each do |sn|
        if @sheets[sn].each_value.any?(Date)
          date_num_fmt_id
          break
        end
      end

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
      total_pivot_caches = @pivot_tables_data.values.sum { |v| v.size }

      entries = {
        "_rels/.rels" => generate_rels_root,
        "xl/workbook.xml" => generate_workbook_xml(total_pivot_caches),
        "xl/styles.xml" => generate_styles_xml
      }

      entries["docProps/core.xml"] = generate_core_properties_xml unless @core_properties.empty?
      entries["docProps/app.xml"] = generate_app_properties_xml unless @app_properties.empty?
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
        sheet_pivots.each { @pivot_cache_count += 1; @pivot_table_count += 1 }

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
          has_drawing, has_comments,
          @sheet_protection[sheet_name], vml_rid
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
        unless sheet_rels_parts.empty?
          entries["xl/worksheets/_rels/sheet#{i + 1}.xml.rels"] = generate_generic_rels(sheet_rels_parts)
        end

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
      entries["xl/_rels/workbook.xml.rels"] = generate_workbook_rels(entries.key?("xl/calcChain.xml"))

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
      icon_set: "iconSet"
    }.freeze

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

      # Merge extra overrides from pass-through.
      overrides.merge!(@extra_ct_overrides) { |_k, generated, _extra| generated }

      parts = [XML_HEADER, %(<Types xmlns="#{CT_NS}">)]
      defaults.each { |ext, ct| parts << %(<Default Extension="#{ext}" ContentType="#{ct}"/>) }
      overrides.each { |pn, ct| parts << %(<Override PartName="#{pn}" ContentType="#{ct}"/>) }
      parts << "</Types>"
      parts.join
    end

    def generate_rels_root
      parts = [
        XML_HEADER,
        %(<Relationships xmlns="#{REL_NS}">),
        %(<Relationship Id="rId1" Type="#{DOC_REL_NS}/officeDocument" Target="xl/workbook.xml"/>)
      ]
      parts << %(<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>) unless @core_properties.empty?
      if @app_properties.any?
        rid = @core_properties.empty? ? "rId2" : "rId3"
        parts << %(<Relationship Id="#{rid}" Type="#{DOC_REL_NS}/extended-properties" Target="docProps/app.xml"/>)
      end
      parts << "</Relationships>"
      parts.join
    end

    def generate_workbook_xml(pivot_cache_count = 0)
      parts = [
        XML_HEADER,
        %(<workbook xmlns="#{SSML_NS}" xmlns:r="#{DOC_REL_NS}">)
      ]

      # workbookPr
      unless @workbook_properties.empty?
        attrs = []
        attrs << %(date1904="#{@workbook_properties[:date1904] ? 1 : 0}") unless @workbook_properties[:date1904].nil?
        attrs << %(defaultThemeVersion="#{@workbook_properties[:default_theme_version]}") if @workbook_properties[:default_theme_version]
        parts << "<workbookPr #{attrs.join(" ")}/>" unless attrs.empty?
      end

      # bookViews/workbookView
      unless @workbook_views.empty?
        attrs = []
        attrs << %(activeTab="#{@workbook_views[:active_tab]}") if @workbook_views[:active_tab]
        attrs << %(firstSheet="#{@workbook_views[:first_sheet]}") if @workbook_views[:first_sheet]
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
        if @workbook_protection[:algorithm_name]
          wp_attrs << %(workbookAlgorithmName="#{xml_escape(@workbook_protection[:algorithm_name])}")
          wp_attrs << %(workbookHashValue="#{xml_escape(@workbook_protection[:hash_value])}") if @workbook_protection[:hash_value]
          wp_attrs << %(workbookSaltValue="#{xml_escape(@workbook_protection[:salt_value])}") if @workbook_protection[:salt_value]
          wp_attrs << %(workbookSpinCount="#{@workbook_protection[:spin_count]}") if @workbook_protection[:spin_count]
        elsif @workbook_protection[:password]
          wp_attrs << %(workbookPassword="#{xml_escape(@workbook_protection[:password])}")
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
        parts << "<calcPr #{attrs.join(" ")}/>" unless attrs.empty?
      end

      # pivotCaches (reference from workbook to cache definition rels)
      if pivot_cache_count > 0
        # rId layout: sheets(1..N), styles(N+1), optional SST(N+2), then pivot caches
        pivot_rid_base = @sheet_order.size + 1 + (@use_shared_strings ? 1 : 0) + 1
        parts << "<pivotCaches>"
        pivot_cache_count.times do |ci|
          parts << %(<pivotCache cacheId="#{ci + 1}" r:id="rId#{pivot_rid_base + ci}"/>)
        end
        parts << "</pivotCaches>"
      end

      parts << "</workbook>"
      parts.join
    end

    def generate_workbook_rels(has_calc_chain = false)
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

    def generate_worksheet_xml(sheet_cells, sheet_col_widths, sheet_col_attrs, sheet_row_attrs, sheet_auto_filter, sheet_filter_cols, sheet_sort, sheet_merge_cells, sheet_hyperlinks, sheet_cell_styles, sheet_props, sheet_fmt, sheet_sv, sheet_fp, sheet_sel, sheet_po, sheet_pm, sheet_ps, sheet_hf, sheet_rb, sheet_cb, sheet_dv, sheet_cf, sst = nil, sheet_tables = [], hyperlink_count = 0, has_drawing = false, has_comments = false, sheet_prot = nil, vml_rid = nil)
      needs_r_ns = !sheet_hyperlinks.empty? || sheet_tables.any? || has_drawing
      worksheet_attrs = %(xmlns="#{SSML_NS}")
      worksheet_attrs << %( xmlns:r="#{DOC_REL_NS}") if needs_r_ns
      parts = [
        XML_HEADER,
        "<worksheet #{worksheet_attrs}>"
      ]

      # Emit <sheetPr> if sheet properties are defined.
      unless sheet_props.empty?
        sp_children = []
        sp_children << %(<tabColor rgb="#{sheet_props[:tab_color]}"/>) if sheet_props[:tab_color]
        sb = sheet_props[:summary_below]
        sr = sheet_props[:summary_right]
        unless sb.nil? && sr.nil?
          outline_attrs = []
          outline_attrs << %(summaryBelow="#{sb ? 1 : 0}") unless sb.nil?
          outline_attrs << %(summaryRight="#{sr ? 1 : 0}") unless sr.nil?
          sp_children << "<outlinePr #{outline_attrs.join(" ")}/>"
        end
        unless sp_children.empty?
          parts << "<sheetPr>"
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
        sgl = sheet_sv[:show_grid_lines]
        sv_attrs << %(showGridLines="#{sgl ? 1 : 0}") unless sgl.nil?
        srch = sheet_sv[:show_row_col_headers]
        sv_attrs << %(showRowColHeaders="#{srch ? 1 : 0}") unless srch.nil?
        rtl = sheet_sv[:right_to_left]
        sv_attrs << %(rightToLeft="#{rtl ? 1 : 0}") unless rtl.nil?
        zs = sheet_sv[:zoom_scale]
        sv_attrs << %(zoomScale="#{zs}") if zs
        sv_attrs << 'tabSelected="1"' if sheet_sv[:tab_selected]
        sv_attrs << 'workbookViewId="0"'
        parts << "<sheetView #{sv_attrs.join(" ")}>"

        if sheet_fp && (sheet_fp[:row].to_i.positive? || sheet_fp[:col].to_i.positive?)
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
          sel_attrs << %(activeCell="#{sheet_sel[:active_cell]}") if sheet_sel[:active_cell]
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
        end
        parts << "<row #{attrs}>"
        row_cells.sort_by { |col, _| column_letter_to_index(col) }.each do |col_letter, value|
          cell_ref = "#{col_letter}#{row_num}"
          style_idx = resolve_style_index(sheet_cell_styles[cell_ref])
          parts << cell_xml(cell_ref, value, style_idx, sst)
        end
        parts << "</row>"
      end

      parts << "</sheetData>"

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

      # Emit <autoFilter> with optional filterColumns.
      if sheet_auto_filter
        if sheet_filter_cols.empty?
          parts << %(<autoFilter ref="#{sheet_auto_filter}"/>)
        else
          parts << %(<autoFilter ref="#{sheet_auto_filter}">)
          sheet_filter_cols.sort.each do |col_id, filter|
            parts << %(<filterColumn colId="#{col_id}">)
            parts << emit_filter_xml(filter)
            parts << "</filterColumn>"
          end
          parts << "</autoFilter>"
        end
      end

      # Emit <sortState> if defined.
      if sheet_sort
        parts << %(<sortState ref="#{sheet_sort[:ref]}">)
        sheet_sort[:sort_conditions].each do |sc|
          sc_attrs = %(ref="#{sc[:ref]}")
          sc_attrs << ' descending="1"' if sc[:descending]
          parts << "<sortCondition #{sc_attrs}/>"
        end
        parts << "</sortState>"
      end

      # Emit <mergeCells> if merge ranges are defined.
      unless sheet_merge_cells.empty?
        parts << %(<mergeCells count="#{sheet_merge_cells.size}">)
        sheet_merge_cells.each { |ref| parts << %(<mergeCell ref="#{ref}"/>) }
        parts << "</mergeCells>"
      end

      # Emit <hyperlinks> if hyperlinks are defined.
      unless sheet_hyperlinks.empty?
        parts << "<hyperlinks>"
        sheet_hyperlinks.each_with_index do |(cell_ref, _url), idx|
          parts << %(<hyperlink ref="#{cell_ref}" r:id="rId#{idx + 1}"/>)
        end
        parts << "</hyperlinks>"
      end

      # Emit <conditionalFormatting> if defined.
      unless sheet_cf.empty?
        sheet_cf.group_by { |cf| cf[:sqref] }.each do |sqref, rules|
          parts << %(<conditionalFormatting sqref="#{sqref}">)
          rules.each do |cf|
            emit_cf_rule(parts, cf)
          end
          parts << "</conditionalFormatting>"
        end
      end

      # Emit <dataValidations> if defined.
      unless sheet_dv.empty?
        parts << %(<dataValidations count="#{sheet_dv.size}">)
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
        parts << "<pageSetup #{ps_attrs.join(" ")}/>" unless ps_attrs.empty?
      end

      # Emit <headerFooter> if defined.
      unless sheet_hf.empty?
        parts << "<headerFooter>"
        parts << "<oddHeader>#{xml_escape(sheet_hf[:odd_header])}</oddHeader>" if sheet_hf[:odd_header]
        parts << "<oddFooter>#{xml_escape(sheet_hf[:odd_footer])}</oddFooter>" if sheet_hf[:odd_footer]
        parts << "<evenHeader>#{xml_escape(sheet_hf[:even_header])}</evenHeader>" if sheet_hf[:even_header]
        parts << "<evenFooter>#{xml_escape(sheet_hf[:even_footer])}</evenFooter>" if sheet_hf[:even_footer]
        parts << "</headerFooter>"
      end

      # Emit <rowBreaks> if defined.
      unless sheet_rb.empty?
        parts << %(<rowBreaks count="#{sheet_rb.size}" manualBreakCount="#{sheet_rb.size}">)
        sheet_rb.each { |r| parts << %(<brk id="#{r}" max="16383" man="1"/>) }
        parts << "</rowBreaks>"
      end

      # Emit <colBreaks> if defined.
      unless sheet_cb.empty?
        parts << %(<colBreaks count="#{sheet_cb.size}" manualBreakCount="#{sheet_cb.size}">)
        sheet_cb.each { |c| parts << %(<brk id="#{c}" max="1048575" man="1"/>) }
        parts << "</colBreaks>"
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
      if vml_rid
        parts << %(<legacyDrawing r:id="rId#{vml_rid}"/>)
      end

      parts << "</worksheet>"
      parts.join
    end

    def build_shared_string_table
      sst = {}
      @sheets.each_value do |sheet_cells|
        sheet_cells.each_value do |value|
          case value
          when RichText
            sst[value.object_id] = sst.size unless sst.key?(value.object_id)
          when String
            sst[value] = sst.size unless sst.key?(value)
          else
            next if value.is_a?(Numeric) || value.is_a?(Date) || value == true || value == false || value.is_a?(Formula)

            str = value.to_s
            sst[str] = sst.size unless sst.key?(str)
          end
        end
      end
      sst
    end

    def generate_shared_strings_xml(sst)
      # Collect RichText objects for lookup by object_id.
      rt_by_id = {}
      @sheets.each_value do |sheet_cells|
        sheet_cells.each_value do |value|
          rt_by_id[value.object_id] = value if value.is_a?(RichText)
        end
      end

      parts = [XML_HEADER, %(<sst xmlns="#{SSML_NS}" count="#{sst.size}" uniqueCount="#{sst.size}">)]
      sst.each_key do |key|
        rt = rt_by_id[key]
        if rt
          parts << "<si>#{rich_text_xml(rt)}</si>"
        else
          parts << "<si><t>#{xml_escape(key)}</t></si>"
        end
      end
      parts << "</sst>"
      parts.join
    end

    def generate_table_xml(tbl)
      trc = tbl[:totals_row_count].to_i
      table_attrs = %(xmlns="#{SSML_NS}" id="#{tbl[:id]}" name="#{xml_escape(tbl[:name])}" displayName="#{xml_escape(tbl[:display_name])}" ref="#{tbl[:ref]}")
      table_attrs << %( totalsRowCount="#{trc}") if trc.positive?
      table_attrs << ' totalsRowShown="0"' if trc.zero?
      parts = [
        XML_HEADER,
        "<table #{table_attrs}>",
        %(<autoFilter ref="#{tbl[:ref]}"/>),
        %(<tableColumns count="#{tbl[:columns].size}">)
      ]
      tbl[:columns].each_with_index do |col, i|
        col_name = col.is_a?(Hash) ? col[:name] : col
        col_attrs = %(id="#{i + 1}" name="#{xml_escape(col_name)}")
        if col.is_a?(Hash) && (col[:totals_row_function] || col[:calculated_column_formula])
          col_attrs << %( totalsRowFunction="#{col[:totals_row_function]}") if col[:totals_row_function]
          parts << "<tableColumn #{col_attrs}>"
          if col[:calculated_column_formula]
            parts << "<calculatedColumnFormula>#{xml_escape(col[:calculated_column_formula])}</calculatedColumnFormula>"
          end
          parts << "</tableColumn>"
        else
          parts << "<tableColumn #{col_attrs}/>"
        end
      end
      parts << "</tableColumns>"
      style = tbl[:style] || {}
      style_name = style[:name] || "TableStyleMedium2"
      sfc = style[:show_first_column] ? "1" : "0"
      slc = style[:show_last_column] ? "1" : "0"
      srs = style.key?(:show_row_stripes) ? (style[:show_row_stripes] ? "1" : "0") : "1"
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
      sheet_hyperlinks.each do |(_cell_ref, url)|
        rid += 1
        parts << %(<Relationship Id="rId#{rid}" Type="#{DOC_REL_NS}/hyperlink" Target="#{xml_escape(url)}" TargetMode="External"/>)
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
      @hyperlinks[sheet_name].each do |(_cell_ref, url)|
        rels << { type: "#{DOC_REL_NS}/hyperlink", target: url, external: true }
      end
      sheet_tables.each_with_index do |_tbl, i|
        rels << { type: "#{DOC_REL_NS}/table", target: "../tables/table#{table_start_index + i + 1}.xml" }
      end
      if comment_idx
        rels << { type: "#{DOC_REL_NS}/comments", target: "../comments#{comment_idx}.xml" }
        rels << { type: "#{DOC_REL_NS}/vmlDrawing", target: "../drawings/vmlDrawing#{comment_idx}.vml" }
      end
      if drawing_idx
        rels << { type: "#{DOC_REL_NS}/drawing", target: "../drawings/drawing#{drawing_idx}.xml" }
      end
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

    XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
    C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"

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
          parts << '<xdr:twoCellAnchor editAs="oneCell">'
          parts << anchor_xml("from", img[:from_col], img[:from_row])
          parts << anchor_xml("to", img[:to_col], img[:to_row])
          parts << "<xdr:pic>"
          parts << %(<xdr:nvPicPr><xdr:cNvPr id="#{dp[:rid_index] + 1}" name="#{xml_escape(img[:name])}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>)
          parts << %(<xdr:blipFill><a:blip r:embed="#{rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>)
          parts << '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>'
          parts << "</xdr:pic>"
          parts << "<xdr:clientData/>"
          parts << "</xdr:twoCellAnchor>"
        when :chart
          chart = dp[:chart]
          rid = "rId#{dp[:rid_index]}"
          parts << "<xdr:twoCellAnchor>"
          parts << anchor_xml("from", 0, 0)
          parts << anchor_xml("to", 10, 15)
          parts << %(<xdr:graphicFrame macro="">)
          parts << %(<xdr:nvGraphicFramePr><xdr:cNvPr id="#{dp[:rid_index] + 1}" name="#{xml_escape(chart[:title] || "Chart")}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>)
          parts << '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="5000000" cy="3000000"/></xdr:xfrm>'
          parts << %(<a:graphic><a:graphicData uri="#{C_NS}"><c:chart xmlns:c="#{C_NS}" r:id="#{rid}"/></a:graphicData></a:graphic>)
          parts << "</xdr:graphicFrame>"
          parts << "<xdr:clientData/>"
          parts << "</xdr:twoCellAnchor>"
        when :sp
          shape = dp[:shape]
          parts << "<xdr:twoCellAnchor>"
          parts << anchor_xml("from", shape[:from_col], shape[:from_row])
          parts << anchor_xml("to", shape[:to_col], shape[:to_row])
          parts << "<xdr:sp>"
          parts << %(<xdr:nvSpPr><xdr:cNvPr id="#{dp[:id]}" name="#{xml_escape(shape[:name])}"/><xdr:cNvSpPr/></xdr:nvSpPr>)
          parts << %(<xdr:spPr><a:prstGeom prst="#{xml_escape(shape[:preset])}"><a:avLst/></a:prstGeom></xdr:spPr>)
          if shape[:text]
            parts << '<xdr:txBody><a:bodyPr/><a:lstStyle/>'
            parts << "<a:p><a:r><a:t>#{xml_escape(shape[:text])}</a:t></a:r></a:p>"
            parts << '</xdr:txBody>'
          end
          parts << "</xdr:sp>"
          parts << "<xdr:clientData/>"
          parts << "</xdr:twoCellAnchor>"
        end
      end

      parts << "</xdr:wsDr>"
      parts.join
    end

    def anchor_xml(tag, col, row)
      "<xdr:#{tag}><xdr:col>#{col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>#{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:#{tag}>"
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

    CHART_TYPE_MAP = { bar: "barChart", line: "lineChart", pie: "pieChart" }.freeze

    def generate_chart_xml(chart)
      chart_type = CHART_TYPE_MAP[chart[:type]] || "barChart"
      is_pie = chart[:type] == :pie
      parts = [
        XML_HEADER,
        %(<c:chartSpace xmlns:c="#{C_NS}" xmlns:a="#{A_NS}" xmlns:r="#{DOC_REL_NS}">),
        "<c:chart>"
      ]

      if chart[:title]
        parts << "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>#{xml_escape(chart[:title])}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>"
      end

      parts << "<c:plotArea><c:layout/>"
      parts << "<c:#{chart_type}>"
      parts << '<c:barDir val="col"/><c:grouping val="clustered"/>' if chart_type == "barChart"
      parts << '<c:grouping val="standard"/>' if chart_type == "lineChart"

      all_series = chart[:series] || []
      all_series.each_with_index do |ser, idx|
        parts << "<c:ser><c:idx val=\"#{idx}\"/><c:order val=\"#{idx}\"/>"
        if ser[:name]
          parts << "<c:tx><c:strRef><c:f>#{xml_escape(ser[:name])}</c:f></c:strRef></c:tx>"
        end
        if chart[:data_labels]
          dl = chart[:data_labels]
          parts << "<c:dLbls>"
          parts << "<c:showLegendKey val=\"#{dl[:show_legend_key] ? 1 : 0}\"/>" unless dl[:show_legend_key].nil?
          parts << "<c:showVal val=\"#{dl[:show_val] ? 1 : 0}\"/>" unless dl[:show_val].nil?
          parts << "<c:showCatName val=\"#{dl[:show_cat_name] ? 1 : 0}\"/>" unless dl[:show_cat_name].nil?
          parts << "<c:showSerName val=\"#{dl[:show_ser_name] ? 1 : 0}\"/>" unless dl[:show_ser_name].nil?
          parts << "<c:showPercent val=\"#{dl[:show_percent] ? 1 : 0}\"/>" unless dl[:show_percent].nil?
          parts << "</c:dLbls>"
        end
        if ser[:cat_ref]
          parts << "<c:cat><c:strRef><c:f>#{xml_escape(ser[:cat_ref])}</c:f></c:strRef></c:cat>"
        end
        if ser[:val_ref]
          parts << "<c:val><c:numRef><c:f>#{xml_escape(ser[:val_ref])}</c:f></c:numRef></c:val>"
        end
        parts << "</c:ser>"
      end

      unless is_pie
        parts << '<c:axId val="1"/><c:axId val="2"/>'
      end
      parts << "</c:#{chart_type}>"

      unless is_pie
        parts << '<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>'
        if chart[:cat_axis_title]
          parts << "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>#{xml_escape(chart[:cat_axis_title])}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>"
        end
        parts << '<c:crossAx val="2"/></c:catAx>'
        parts << '<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>'
        if chart[:val_axis_title]
          parts << "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>#{xml_escape(chart[:val_axis_title])}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>"
        end
        parts << '<c:crossAx val="1"/></c:valAx>'
      end

      parts << "</c:plotArea>"
      legend_pos = chart.dig(:legend, :position) || "r"
      parts << %(<c:legend><c:legendPos val="#{legend_pos}"/></c:legend>)
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
        parts << %(<comment ref="#{c[:ref]}" authorId="#{aid}"><text><r><t>#{xml_escape(c[:text])}</t></r></text></comment>)
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
        '</v:shapetype>'
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
        parts << '<x:MoveWithCells/><x:SizeWithCells/>'
        parts << "<x:Anchor>#{col + 1}, 15, #{row}, 10, #{col + 3}, 15, #{row + 4}, 4</x:Anchor>"
        parts << '<x:AutoFill>False</x:AutoFill>'
        parts << "<x:Row>#{row}</x:Row>"
        parts << "<x:Column>#{col}</x:Column>"
        parts << '</x:ClientData></v:shape>'
      end
      parts << '</xml>'
      parts.join
    end

    def cell_to_col_row(cell_ref)
      m = cell_ref.match(/\A([A-Z]+)(\d+)\z/)
      return [0, 0] unless m
      col = column_letter_to_index(m[1]) - 1
      row = m[2].to_i - 1
      [col, row]
    end

    def generate_pivot_table_xml(pt, cache_id)
      data_caption = pt[:data_fields].first ? pt[:data_fields].first[:name] : "Values"
      parts = [
        XML_HEADER,
        %(<pivotTableDefinition xmlns="#{SSML_NS}" name="#{xml_escape(pt[:name])}" cacheId="#{cache_id}" dataCaption="#{xml_escape(data_caption)}" dataOnRows="0" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1">)
      ]

      # Compute field count from source range or explicit field_names.
      field_count = if pt[:field_names]
                      pt[:field_names].size
                    else
                      (pt[:row_fields].size + pt[:col_fields].size + pt[:data_fields].size).clamp(1, 100)
                    end
      parts << %(<location ref="#{pt[:dest_ref]}" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>)
      parts << %(<pivotFields count="#{field_count}">)
      field_count.times do |fi|
        attrs = +""
        if pt[:row_fields].include?(fi)
          attrs << ' axis="axisRow"'
        elsif pt[:col_fields].include?(fi)
          attrs << ' axis="axisCol"'
        end
        attrs << ' dataField="1"' if pt[:data_fields].any? { |df| df[:fld] == fi }
        attrs << ' showAll="0"'

        field_items = pt[:items] && pt[:items][fi]
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

      unless pt[:row_fields].empty?
        parts << %(<rowFields count="#{pt[:row_fields].size}">)
        pt[:row_fields].each { |f| parts << %(<field x="#{f}"/>) }
        parts << "</rowFields>"
      end

      unless pt[:col_fields].empty?
        parts << %(<colFields count="#{pt[:col_fields].size}">)
        pt[:col_fields].each { |f| parts << %(<field x="#{f}"/>) }
        parts << "</colFields>"
      end

      unless pt[:data_fields].empty?
        parts << %(<dataFields count="#{pt[:data_fields].size}">)
        pt[:data_fields].each do |df|
          parts << %(<dataField name="#{xml_escape(df[:name])}" fld="#{df[:fld]}" subtotal="#{df[:subtotal] || "sum"}"/>)
        end
        parts << "</dataFields>"
      end

      parts << "</pivotTableDefinition>"
      parts.join
    end

    def generate_pivot_cache_definition_xml(pt, cache_id)
      parts = [
        XML_HEADER,
        %(<pivotCacheDefinition xmlns="#{SSML_NS}" xmlns:r="#{DOC_REL_NS}" r:id="rId1" refreshOnLoad="1">)
      ]

      # Parse source ref: "Sheet1!A1:C4" => sheet name + range.
      source = pt[:source_ref]
      if source.include?("!")
        sname, srange = source.split("!", 2)
        sname = sname.delete("'")
        parts << %(<cacheSource type="worksheet"><worksheetSource ref="#{srange}" sheet="#{xml_escape(sname)}"/></cacheSource>)
      else
        parts << %(<cacheSource type="worksheet"><worksheetSource ref="#{source}"/></cacheSource>)
      end

      field_count = pt[:field_names] ? pt[:field_names].size : (pt[:row_fields].size + pt[:col_fields].size + pt[:data_fields].size)
      parts << %(<cacheFields count="#{field_count}">)
      field_count.times do |fi|
        fname = if pt[:field_names] && pt[:field_names][fi]
                  pt[:field_names][fi]
                else
                  df = pt[:data_fields].find { |d| d[:fld] == fi }
                  df ? df[:name] : "Field#{fi + 1}"
                end
        field_items = pt[:items] && pt[:items][fi]
        if field_items
          parts << %(<cacheField name="#{xml_escape(fname)}" numFmtId="0">)
          parts << %(<sharedItems count="#{field_items.size}">)
          field_items.each { |v| parts << %(<s v="#{xml_escape(v.to_s)}"/>) }
          parts << "</sharedItems>"
          parts << "</cacheField>"
        else
          parts << %(<cacheField name="#{xml_escape(fname)}" numFmtId="0"><sharedItems/></cacheField>)
        end
      end
      parts << "</cacheFields>"
      parts << "</pivotCacheDefinition>"
      parts.join
    end

    def generate_pivot_cache_records_xml(pt)
      items = pt[:items]
      if items && items.values.any? { |v| v && !v.empty? }
        max_len = items.values.map { |v| v ? v.size : 0 }.max
        parts = [XML_HEADER, %(<pivotCacheRecords xmlns="#{SSML_NS}" count="#{max_len}">)]
        max_len.times do |ri|
          parts << "<r>"
          field_count = pt[:field_names] ? pt[:field_names].size : (pt[:row_fields].size + pt[:col_fields].size + pt[:data_fields].size)
          field_count.times do |fi|
            field_items = items[fi]
            if field_items
              parts << %(<x v="#{ri < field_items.size ? ri : 0}"/>)
            else
              parts << %(<n v="0"/>)
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

    def generate_external_link_xml(el)
      parts = [
        XML_HEADER,
        %(<externalLink xmlns="#{SSML_NS}" xmlns:r="#{DOC_REL_NS}">),
        '<externalBook r:id="rId1">'
      ]
      unless el[:sheet_names].empty?
        parts << "<sheetNames>"
        el[:sheet_names].each { |sn| parts << %(<sheetName val="#{xml_escape(sn)}"/>) }
        parts << "</sheetNames>"
      end
      parts << "</externalBook>"
      parts << "</externalLink>"
      parts.join
    end

    def generate_external_link_rels(el)
      [
        XML_HEADER,
        %(<Relationships xmlns="#{REL_NS}">),
        %(<Relationship Id="rId1" Type="#{DOC_REL_NS}/externalLinkPath" Target="#{xml_escape(el[:target])}" TargetMode="External"/>),
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

    def rich_text_xml(rt)
      rt.runs.map do |run|
        font = run[:font]
        if font && !font.empty?
          rpr = +""
          rpr << "<b/>" if font[:bold]
          rpr << "<i/>" if font[:italic]
          rpr << "<u/>" if font[:underline]
          rpr << %(<sz val="#{font[:sz]}"/>) if font[:sz]
          rpr << %(<color rgb="#{font[:color]}"/>) if font[:color]
          rpr << %(<rFont val="#{xml_escape(font[:name])}"/>) if font[:name]
          "<r><rPr>#{rpr}</rPr><t>#{xml_escape(run[:text])}</t></r>"
        else
          "<r><t>#{xml_escape(run[:text])}</t></r>"
        end
      end.join
    end

    def cell_xml(cell_ref, value, style_idx, sst = nil)
      s_attr = style_idx ? %( s="#{style_idx}") : ""
      case value
      when Formula
        f_attrs = +""
        if value.type == :shared
          f_attrs << %( t="shared" si="#{value.shared_index}")
          f_attrs << %( ref="#{value.ref}") if value.ref
        elsif value.type == :array
          f_attrs << %( t="array" ref="#{value.ref}") if value.ref
        end
        parts = %(<c r="#{cell_ref}"#{s_attr}><f#{f_attrs}>#{xml_escape(value.expression)}</f>)
        parts << "<v>#{xml_escape(value.cached_value.to_s)}</v>" unless value.cached_value.nil?
        parts << "</c>"
        parts
      when RichText
        if sst
          idx = sst[value.object_id]
          %(<c r="#{cell_ref}" t="s"#{s_attr}><v>#{idx}</v></c>)
        else
          %(<c r="#{cell_ref}" t="inlineStr"#{s_attr}><is>#{rich_text_xml(value)}</is></c>)
        end
      when true, false
        %(<c r="#{cell_ref}" t="b"#{s_attr}><v>#{value ? 1 : 0}</v></c>)
      when Date
        serial = Xlsxrb.date_to_serial(value)
        date_style = resolve_style_index(date_num_fmt_id)
        ds_attr = date_style ? %( s="#{date_style}") : ""
        %(<c r="#{cell_ref}"#{ds_attr}><v>#{serial}</v></c>)
      when Numeric
        %(<c r="#{cell_ref}"#{s_attr}><v>#{value}</v></c>)
      else
        if sst
          idx = sst[value.to_s]
          %(<c r="#{cell_ref}" t="s"#{s_attr}><v>#{idx}</v></c>)
        else
          %(<c r="#{cell_ref}" t="inlineStr"#{s_attr}><is><t>#{xml_escape(value)}</t></is></c>)
        end
      end
    end

    # Returns the numFmtId for dates, registering it on first use.
    def date_num_fmt_id
      @date_num_fmt_id ||= add_number_format(DEFAULT_DATE_FORMAT)
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
        attrs = filter[:blank] ? ' blank="1"' : ""
        if filter[:values]&.any?
          parts = ["<filters#{attrs}>"]
          filter[:values].each { |v| parts << %(<filter val="#{xml_escape(v)}"/>) }
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
        %(<dynamicFilter type="#{filter[:dynamic_type]}"/>)
      when :top10
        top_attr = filter[:top] ? ' top="1"' : ""
        pct_attr = filter[:percent] ? ' percent="1"' : ""
        %(<top10#{top_attr}#{pct_attr} val="#{filter[:val]}"/>)
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

      case type
      when :cell_is, :expression
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
          parts << "<cfvo #{cfvo_attrs}/>"
        end
        cs[:colors]&.each { |c| parts << %(<color rgb="#{c}"/>) }
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
          parts << "<cfvo #{cfvo_attrs}/>"
        end
        parts << %(<color rgb="#{db[:color]}"/>) if db[:color]
        parts << "</dataBar>"
        parts << "</cfRule>"
      when :icon_set
        is = rule[:icon_set]
        parts << "<cfRule #{rule_attrs}>"
        is_attrs = +""
        is_attrs << %( iconSet="#{is[:icon_set]}") if is[:icon_set]
        is_attrs << %( reverse="#{is[:reverse] ? 1 : 0}") unless is[:reverse].nil?
        is_attrs << %( showValue="#{is[:show_value] ? 1 : 0}") unless is[:show_value].nil?
        parts << "<iconSet#{is_attrs}>"
        is[:cfvo]&.each do |cfvo|
          cfvo_attrs = %(type="#{cfvo[:type]}")
          cfvo_attrs << %( val="#{cfvo[:val]}") if cfvo[:val]
          parts << "<cfvo #{cfvo_attrs}/>"
        end
        parts << "</iconSet>"
        parts << "</cfRule>"
      else
        parts << "<cfRule #{rule_attrs}/>"
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
      parts << "<dc:creator>#{xml_escape(@core_properties[:creator])}</dc:creator>" if @core_properties[:creator]
      parts << %(<dcterms:created xsi:type="dcterms:W3CDTF">#{xml_escape(@core_properties[:created])}</dcterms:created>) if @core_properties[:created]
      parts << %(<dcterms:modified xsi:type="dcterms:W3CDTF">#{xml_escape(@core_properties[:modified])}</dcterms:modified>) if @core_properties[:modified]
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
        if xf[:alignment]
          alignment_xml = emit_alignment_xml(xf[:alignment])
          parts << %(<xf numFmtId="#{xf[:num_fmt_id]}" fontId="#{xf[:font_id]}" fillId="#{xf[:fill_id]}" borderId="#{xf[:border_id]}" xfId="0"#{apply_attrs.join}>#{alignment_xml}</xf>)
        else
          parts << %(<xf numFmtId="#{xf[:num_fmt_id]}" fontId="#{xf[:font_id]}" fillId="#{xf[:fill_id]}" borderId="#{xf[:border_id]}" xfId="0"#{apply_attrs.join}/>)
        end
      end
      parts << "</cellXfs>"

      # dxfs
      unless @dxfs.empty?
        parts << %(<dxfs count="#{@dxfs.size}">)
        @dxfs.each { |d| parts << emit_dxf_xml(d) }
        parts << "</dxfs>"
      end

      # cellStyles
      parts << %(<cellStyles count="#{@cell_style_names.size}">)
      @cell_style_names.each do |cs|
        cs_attrs = %(name="#{xml_escape(cs[:name])}" xfId="#{cs[:xf_id]}")
        cs_attrs << %( builtinId="#{cs[:builtin_id]}") if cs[:builtin_id]
        parts << "<cellStyle #{cs_attrs}/>"
      end
      parts << "</cellStyles>"

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
      attrs << %(shrinkToFit="1") if alignment[:shrink_to_fit]
      "<alignment #{attrs.join(" ")}/>"
    end

    def emit_font_xml(font)
      parts = ["<font>"]
      parts << "<b/>" if font[:bold]
      parts << "<i/>" if font[:italic]
      parts << "<strike/>" if font[:strike]
      if font[:underline]
        if font[:underline] == true
          parts << "<u/>"
        else
          parts << %(<u val="#{font[:underline]}"/>)
        end
      end
      parts << %(<vertAlign val="#{font[:vert_align]}"/>) if font[:vert_align]
      parts << %(<sz val="#{font[:sz]}"/>) if font[:sz]
      parts << %(<color rgb="#{font[:color]}"/>) if font[:color]
      parts << %(<name val="#{xml_escape(font[:name])}"/>) if font[:name]
      parts << %(<family val="#{font[:family]}"/>) if font[:family]
      parts << %(<scheme val="#{font[:scheme]}"/>) if font[:scheme]
      parts << "</font>"
      parts.join
    end

    def emit_fill_xml(fill)
      if fill[:gradient]
        return emit_gradient_fill_xml(fill[:gradient])
      end
      return "<fill><patternFill patternType=\"#{fill[:pattern]}\"/></fill>" if fill[:pattern] && !fill[:fg_color] && !fill[:bg_color]

      parts = ["<fill>"]
      pt = fill[:pattern] || "solid"
      parts << %(<patternFill patternType="#{pt}">)
      parts << %(<fgColor rgb="#{fill[:fg_color]}"/>) if fill[:fg_color]
      parts << %(<bgColor rgb="#{fill[:bg_color]}"/>) if fill[:bg_color]
      parts << "</patternFill>"
      parts << "</fill>"
      parts.join
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
      parts << "<gradientFill#{attrs.empty? ? "" : " #{attrs.join(" ")}"}"
      if gradient[:stops]&.any?
        parts[-1] = "#{parts[-1]}>"
        gradient[:stops].each do |stop|
          parts << %(<stop position="#{stop[:position]}"><color rgb="#{stop[:color]}"/></stop>)
        end
        parts << "</gradientFill>"
      else
        parts[-1] = "#{parts[-1]}/>"
      end
      parts << "</fill>"
      parts.join
    end

    def emit_border_xml(bdr)
      parts = ["<border>"]
      %i[left right top bottom].each do |side|
        s = bdr[side]
        if s.is_a?(Hash)
          parts << %(<#{side} style="#{s[:style]}">)
          parts << %(<color rgb="#{s[:color]}"/>) if s[:color]
          parts << "</#{side}>"
        else
          parts << "<#{side}/>"
        end
      end
      parts << "<diagonal/>"
      parts << "</border>"
      parts.join
    end

    def emit_dxf_xml(dxf)
      parts = ["<dxf>"]
      parts << emit_font_xml(dxf[:font]) if dxf[:font]
      parts << emit_fill_xml(dxf[:fill]) if dxf[:fill]
      parts << emit_border_xml(dxf[:border]) if dxf[:border]
      if dxf[:num_fmt]
        nf = dxf[:num_fmt]
        parts << %(<numFmt numFmtId="#{nf[:num_fmt_id]}" formatCode="#{xml_escape(nf[:format_code])}"/>)
      end
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
  end
end
