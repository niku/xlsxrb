# frozen_string_literal: true

# rbs_inline: enabled

require_relative "zip_generator"
require_relative "xml_builder"
require_relative "writer/drawing_xml"
require_relative "writer/features_xml"
require_relative "writer/styles_xml"
require_relative "../dsl_helpers"

module Xlsxrb
  module Ooxml
    # Writes spreadsheet data into an XLSX file.
    #
    # @api private
    class Writer
      include DrawingXml
      include FeaturesXml
      include StylesXml

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

      attr_reader :fonts, :fills, :borders, :xf_entries, :num_fmts, :dxfs

      # : (?default_font: untyped?) -> void
      def initialize(default_font: nil)
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
        default_font(default_font) if default_font
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
        @threaded_comments_data = { "Sheet1" => [] }
        @threaded_comment_count = 0
        @persons_data = []
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
      # : (untyped name) -> untyped
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
        @threaded_comments_data[name] = []
        @pivot_tables_data[name] = []
        @sheet_protection[name] = nil
        @protected_ranges[name] = []
        @phonetic_properties[name] = nil
        @sheet_order << name
      end

      # Registers a cell value at the given address (e.g. "A1").
      # : (untyped cell_address, untyped value, ?sheet: untyped?) -> untyped
      def set_cell(cell_address, value, sheet: nil)
        validate_cell_address!(cell_address)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @sheets.key?(sheet_name)

        @sheets[sheet_name][cell_address] = value
      end

      # Returns the registered cells for the first (or given) sheet.
      # : (?sheet: untyped?) -> untyped
      def cells(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        @sheets[sheet_name] || {}
      end

      # Sets the width for a column (e.g. "A", "BC").
      # : (untyped col_letter, untyped width, ?sheet: untyped?) -> untyped
      def set_column_width(col_letter, width, sheet: nil)
        raise ArgumentError, "column must be a String of uppercase letters" unless col_letter.is_a?(String) && col_letter.match?(/\A[A-Z]+\z/)

        col_index = column_letter_to_index(col_letter)
        raise ArgumentError, "column out of range: #{col_letter}" unless col_index.between?(1, MAX_COLUMN_INDEX)

        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @column_widths.key?(sheet_name)

        @column_widths[sheet_name][col_letter] = width
      end

      # Returns column widths for the first (or given) sheet.
      # : (?sheet: untyped?) -> untyped
      def column_widths(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        @column_widths[sheet_name] || {}
      end

      # Sets a column attribute (e.g. :hidden, :best_fit, :outline_level, :collapsed, :style).
      # : (untyped col_letter, untyped name, untyped value, ?sheet: untyped?) -> untyped
      def set_column_attribute(col_letter, name, value, sheet: nil)
        raise ArgumentError, "column must be a String of uppercase letters" unless col_letter.is_a?(String) && col_letter.match?(/\A[A-Z]+\z/)

        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @column_attrs.key?(sheet_name)

        @column_attrs[sheet_name][col_letter] ||= {}
        @column_attrs[sheet_name][col_letter][name] = value
      end

      # Returns column attributes for the first (or given) sheet.
      # : (?sheet: untyped?) -> untyped
      def column_attributes(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        @column_attrs[sheet_name] || {}
      end

      # Sets a row height.
      # : (untyped row_num, untyped height, ?sheet: untyped?) -> untyped
      def set_row_height(row_num, height, sheet: nil)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @row_attrs.key?(sheet_name)
        raise ArgumentError, "row must be a positive Integer" unless row_num.is_a?(Integer) && row_num >= 1

        @row_attrs[sheet_name][row_num] ||= {}
        @row_attrs[sheet_name][row_num][:height] = height
      end

      # Hides a row.
      # : (untyped row_num, ?hidden: bool, ?sheet: untyped?) -> untyped
      def set_row_hidden(row_num, hidden: true, sheet: nil)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @row_attrs.key?(sheet_name)
        raise ArgumentError, "row must be a positive Integer" unless row_num.is_a?(Integer) && row_num >= 1

        @row_attrs[sheet_name][row_num] ||= {}
        @row_attrs[sheet_name][row_num][:hidden] = hidden
      end

      # Sets a row outline level.
      # : (untyped row_num, untyped level, ?sheet: untyped?) -> untyped
      def set_row_outline_level(row_num, level, sheet: nil)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @row_attrs.key?(sheet_name)
        raise ArgumentError, "row must be a positive Integer" unless row_num.is_a?(Integer) && row_num >= 1

        @row_attrs[sheet_name][row_num] ||= {}
        @row_attrs[sheet_name][row_num][:outline_level] = level
      end

      # Sets a row collapsed state.
      # : (untyped row_num, ?collapsed: bool, ?sheet: untyped?) -> untyped
      def set_row_collapsed(row_num, collapsed: true, sheet: nil)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @row_attrs.key?(sheet_name)
        raise ArgumentError, "row must be a positive Integer" unless row_num.is_a?(Integer) && row_num >= 1

        @row_attrs[sheet_name][row_num] ||= {}
        @row_attrs[sheet_name][row_num][:collapsed] = collapsed
      end

      # Sets a row default style index (cellXfs index).
      # : (untyped row_num, untyped style_id, ?sheet: untyped?) -> untyped
      def set_row_style(row_num, style_id, sheet: nil)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @row_attrs.key?(sheet_name)
        raise ArgumentError, "row must be a positive Integer" unless row_num.is_a?(Integer) && row_num >= 1
        raise ArgumentError, "style_id must be a non-negative Integer" unless style_id.is_a?(Integer) && style_id >= 0

        @row_attrs[sheet_name][row_num] ||= {}
        @row_attrs[sheet_name][row_num][:style] = style_id
      end

      # Sets a row thick top border flag.
      # : (untyped row_num, ?thick: bool, ?sheet: untyped?) -> untyped
      def set_row_thick_top(row_num, thick: true, sheet: nil)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @row_attrs.key?(sheet_name)
        raise ArgumentError, "row must be a positive Integer" unless row_num.is_a?(Integer) && row_num >= 1

        @row_attrs[sheet_name][row_num] ||= {}
        @row_attrs[sheet_name][row_num][:thick_top] = thick
      end

      # Sets a row thick bottom border flag.
      # : (untyped row_num, ?thick: bool, ?sheet: untyped?) -> untyped
      def set_row_thick_bot(row_num, thick: true, sheet: nil)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @row_attrs.key?(sheet_name)
        raise ArgumentError, "row must be a positive Integer" unless row_num.is_a?(Integer) && row_num >= 1

        @row_attrs[sheet_name][row_num] ||= {}
        @row_attrs[sheet_name][row_num][:thick_bot] = thick
      end

      # Sets a row phonetic flag.
      # : (untyped row_num, ?phonetic: bool, ?sheet: untyped?) -> untyped
      def set_row_phonetic(row_num, phonetic: true, sheet: nil)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @row_attrs.key?(sheet_name)
        raise ArgumentError, "row must be a positive Integer" unless row_num.is_a?(Integer) && row_num >= 1

        @row_attrs[sheet_name][row_num] ||= {}
        @row_attrs[sheet_name][row_num][:ph] = phonetic
      end

      # Returns row attributes for the first (or given) sheet.
      # : (?sheet: untyped?) -> untyped
      def row_attributes(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        @row_attrs[sheet_name] || {}
      end

      # Merges a range of cells (e.g. "A1:B2").
      # : (untyped range, ?sheet: untyped?) -> untyped
      def merge(range, sheet: nil)
        raise ArgumentError, "range must be a String like 'A1:B2'" unless range.is_a?(String) && range.match?(/\A[A-Z]+\d+:[A-Z]+\d+\z/)

        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @merge_cells.key?(sheet_name)

        @merge_cells[sheet_name] << range
      end

      # Returns merged cell ranges for the first (or given) sheet.
      # : (?sheet: untyped?) -> untyped
      def merged_cells(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        @merge_cells[sheet_name] || []
      end

      # Adds a hyperlink on a cell.
      # : (untyped cell_address, ?untyped? url, ?sheet: untyped?, ?display: untyped?, ?tooltip: untyped?, ?location: untyped?) -> untyped
      def hyperlink(cell_address, url = nil, sheet: nil, display: nil, tooltip: nil, location: nil)
        validate_cell_address!(cell_address)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @hyperlinks.key?(sheet_name)

        if url.is_a?(Elements::Hyperlink)
          link = {}
          link[:url] = url.url if url.url
          link[:display] = url.display || display if url.display || display
          link[:tooltip] = url.tooltip || tooltip if url.tooltip || tooltip
          link[:location] = url.location || location if url.location || location
          raise ArgumentError, "url or location required" if link[:url].nil? && link[:location].nil?

          @hyperlinks[sheet_name][cell_address] = link
          return
        end

        raise ArgumentError, "url or location required" if url.nil? && location.nil?

        link = {}
        link[:url] = url if url
        link[:display] = display if display
        link[:tooltip] = tooltip if tooltip
        link[:location] = location if location
        @hyperlinks[sheet_name][cell_address] = link
      end

      # Returns hyperlinks for the first (or given) sheet.
      # : (?sheet: untyped?) -> untyped
      def hyperlinks(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        @hyperlinks[sheet_name] || {}
      end

      # Sets an autoFilter range (e.g. "A1:B10") for the given sheet.
      # : (untyped range, ?sheet: untyped?) -> untyped
      def set_auto_filter(range, sheet: nil)
        raise ArgumentError, "range must be a String like 'A1:B10'" unless range.is_a?(String) && range.match?(/\A[A-Z]+\d+:[A-Z]+\d+\z/)

        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @auto_filters.key?(sheet_name)

        @auto_filters[sheet_name] = range
      end

      # Returns the autoFilter range for the first (or given) sheet.
      # : (?sheet: untyped?) -> untyped
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
      # : (untyped col_id, untyped filter, ?sheet: untyped?) -> untyped
      def add_filter_column(col_id, filter, sheet: nil)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @filter_columns.key?(sheet_name)

        @filter_columns[sheet_name][col_id] = filter
      end

      # Returns filter columns for the first (or given) sheet.
      # : (?sheet: untyped?) -> untyped
      def filter_columns(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        @filter_columns[sheet_name] || {}
      end

      # Sets a sort state for the sheet. ref: sort range, sort_conditions: array of { ref:, descending: }.
      # Options: column_sort:, case_sensitive:, sort_method:
      # : (untyped ref, untyped sort_conditions, ?sheet: untyped?, **untyped opts) -> untyped
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
      # : (?sheet: untyped?) -> untyped
      def sort_state(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        @sort_state[sheet_name]
      end

      # Registers a custom number format and returns its numFmtId (starting at 164).
      # : (untyped format_code) -> untyped
      def add_number_format(format_code)
        existing = @num_fmts.find { |nf| nf[:format_code] == format_code }
        return existing[:num_fmt_id] if existing

        num_fmt_id = 164 + @num_fmts.size
        @num_fmts << { num_fmt_id: num_fmt_id, format_code: format_code }
        num_fmt_id
      end

      # Sets a number format on a cell. num_fmt_id is from add_number_format or a built-in id.
      # : (untyped cell_address, untyped num_fmt_id, ?sheet: untyped?) -> untyped
      def set_cell_format(cell_address, num_fmt_id, sheet: nil)
        validate_cell_address!(cell_address)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @cell_styles.key?(sheet_name)

        @cell_styles[sheet_name][cell_address] = num_fmt_id
      end

      # Registers a font and returns its font_id. Opts: bold, italic, sz, color, name.
      # : (**untyped opts) -> untyped
      def add_font(**opts)
        existing = @fonts.index(opts)
        return existing if existing

        @fonts << opts
        @fonts.size - 1
      end

      # Registers a fill and returns its fill_id. Opts: pattern, fg_color, bg_color.
      # : (**untyped opts) -> untyped
      def add_fill(**opts)
        existing = @fills.index(opts)
        return existing if existing

        @fills << opts
        @fills.size - 1
      end

      # Registers a border and returns its border_id.
      # Opts: left, right, top, bottom (each a hash with :style and optional :color).
      # : (**untyped opts) -> untyped
      def add_border(**opts)
        existing = @borders.index(opts)
        return existing if existing

        @borders << opts
        @borders.size - 1
      end

      # Registers a cell style (xf entry) and returns its index for use with set_cell_style.
      # Opts: font_id, fill_id, border_id, num_fmt_id, alignment (hash with horizontal, vertical, wrap_text, text_rotation, indent, shrink_to_fit).
      # : (**untyped opts) -> untyped
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
      # : (name: untyped, ?num_fmt_id: ::Integer, ?font_id: ::Integer, ?fill_id: ::Integer, ?border_id: ::Integer, ?builtin_id: untyped?, ?i_level: untyped?, ?hidden: untyped?, ?custom_builtin: untyped?) -> untyped
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
      # : (untyped cell_address, untyped style_id, ?sheet: untyped?) -> untyped
      def set_cell_style(cell_address, style_id, sheet: nil)
        validate_cell_address!(cell_address)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @cell_styles.key?(sheet_name)

        @cell_styles[sheet_name][cell_address] = { xf_index: style_id }
      end

      # Marks a cell as containing phonetic text (ph="1" on the <c> element).
      # : (untyped cell_address, ?sheet: untyped?) -> untyped
      def set_cell_phonetic(cell_address, sheet: nil)
        validate_cell_address!(cell_address)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @cell_phonetic.key?(sheet_name)

        @cell_phonetic[sheet_name][cell_address] = true
      end

      # Registers a differential format (dxf) for conditional formatting. Returns dxf_id.
      # Opts: font (hash), fill (hash), border (hash), num_fmt (hash).
      # : (**untyped opts) -> untyped
      def add_dxf(**opts)
        @dxfs << opts
        @dxfs.size - 1
      end

      # Sets the workbook-wide default font (font 0).
      #
      # @param name [String, Hash, nil] Font name or font options Hash.
      # @param size [Float, Integer, nil] Font size in points.
      # @param opts [Hash] Additional font options (e.g. `bold:`, `italic:`, `family:`, `charset:`, `color:`).
      # @return [void]
      # @api public
      #: (?String | Hash[Symbol, untyped]? name, ?Float | Integer | nil size, **untyped opts) -> void
      def default_font(name = nil, size = nil, **opts)
        if name.is_a?(Hash)
          opts = name.merge(opts)
          name = opts[:name]
        end

        font = @fonts[0] ? @fonts[0].dup : { name: "Calibri", sz: 11, size: 11 }
        font[:name] = (name || opts[:name]).to_s if name || opts[:name]
        sz_val = size || opts[:size] || opts[:sz]
        if sz_val
          font[:sz] = sz_val
          font[:size] = sz_val
        end
        font[:family] = opts[:family] if opts.key?(:family)
        font[:charset] = opts[:charset] if opts.key?(:charset)
        font[:scheme] = opts[:scheme].to_s if opts.key?(:scheme)
        font[:color] = opts[:color] if opts.key?(:color)
        font[:bold] = opts[:bold] if opts.key?(:bold)
        font[:italic] = opts[:italic] if opts.key?(:italic)
        font[:strike] = opts[:strike] if opts.key?(:strike)
        font[:underline] = opts[:underline] if opts.key?(:underline)

        @fonts[0] = font
      end

      # Sets the indexed colors palette (array of ARGB hex strings, e.g. ["FF000000", "FFFFFFFF"]).
      # rubocop:disable Naming/AccessorMethodName
      # : (untyped colors) -> untyped
      def set_indexed_colors(colors)
        @indexed_colors = colors
      end
      # rubocop:enable Naming/AccessorMethodName

      # Returns the indexed colors palette.
      # : () -> untyped
      def indexed_colors
        @indexed_colors.dup
      end

      # Sets the MRU (most recently used) colors (array of color hashes, e.g. [{rgb: "FFFF0000"}]).
      # rubocop:disable Naming/AccessorMethodName
      # : (untyped colors) -> untyped
      def set_mru_colors(colors)
        @mru_colors = colors
      end
      # rubocop:enable Naming/AccessorMethodName

      # Returns the MRU colors.
      # : () -> untyped
      def mru_colors
        @mru_colors.map(&:dup)
      end

      # Sets table styles options (defaultTableStyle, defaultPivotStyle).
      # : (untyped name, untyped value) -> untyped
      def set_table_styles_option(name, value)
        raise ArgumentError, "name must be a Symbol" unless name.is_a?(Symbol)

        @table_styles[name] = value
      end

      # Adds a table style definition. Returns the style name.
      # elements: array of { type:, dxf_id:, size: }
      # : (name: untyped, ?elements: untyped, ?pivot: untyped?, ?table: untyped?) -> untyped
      def add_table_style(name:, elements: [], pivot: nil, table: nil)
        @table_styles[:styles] ||= []
        style = { name: name, elements: elements }
        style[:pivot] = pivot unless pivot.nil?
        style[:table] = table unless table.nil?
        @table_styles[:styles] << style
        name
      end

      # Returns table styles configuration.
      # : () -> untyped
      def table_styles
        deep_copy = {}
        @table_styles.each do |k, v|
          deep_copy[k] = v.is_a?(Array) ? v.map(&:dup) : v
        end
        deep_copy
      end

      # Sets a core property.
      # : (untyped name, untyped value) -> untyped
      def core_property(name, value)
        raise ArgumentError, "name must be a Symbol" unless name.is_a?(Symbol)

        @core_properties[name] = value
      end

      # Returns core properties hash.
      # : () -> untyped
      def core_properties
        @core_properties.dup
      end

      # Sets an app property.
      # : (untyped name, untyped value) -> untyped
      def app_property(name, value)
        raise ArgumentError, "name must be a Symbol" unless name.is_a?(Symbol)

        @app_properties[name] = value
      end

      # Returns app properties hash.
      # : () -> untyped
      def app_properties
        @app_properties.dup
      end

      # Adds a custom document property. type: :string (default), :number, :bool, :date.
      # : (untyped name, untyped value, ?type: ::Symbol) -> untyped
      def custom_property(name, value, type: :string)
        @custom_properties << { name: name, value: value, type: type }
      end

      # Returns custom properties array.
      # : () -> untyped
      def custom_properties
        @custom_properties.map(&:dup)
      end

      # Sets a sheet-level property (e.g. :tab_color, :summary_below, :summary_right).
      # : (untyped name, untyped value, ?sheet: untyped?) -> untyped
      def set_sheet_property(name, value, sheet: nil)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @sheet_properties.key?(sheet_name)

        @sheet_properties[sheet_name][name] = value
      end

      # Returns sheet properties for the first (or given) sheet.
      # : (?sheet: untyped?) -> untyped
      def sheet_properties(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        (@sheet_properties[sheet_name] || {}).dup
      end

      # Sets phonetic properties for a sheet (e.g. :font_id, :type, :alignment).
      # : (untyped props, ?sheet: untyped?) -> untyped
      def set_phonetic_properties(props, sheet: nil)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @phonetic_properties.key?(sheet_name)

        @phonetic_properties[sheet_name] = props
      end

      # Returns phonetic properties for the first (or given) sheet.
      # : (?sheet: untyped?) -> untyped
      def phonetic_properties(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        @phonetic_properties[sheet_name]&.dup
      end

      # Sets a sheet format property (e.g. :default_row_height, :default_col_width, :base_col_width).
      # : (untyped name, untyped value, ?sheet: untyped?) -> untyped
      def set_sheet_format(name, value, sheet: nil)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @sheet_formats.key?(sheet_name)

        @sheet_formats[sheet_name][name] = value
      end

      # Returns sheet format properties for the first (or given) sheet.
      # : (?sheet: untyped?) -> untyped
      def sheet_format(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        (@sheet_formats[sheet_name] || {}).dup
      end

      # Sets a sheet view property (e.g. :show_grid_lines, :show_row_col_headers, :right_to_left, :zoom_scale).
      # : (untyped name, untyped value, ?sheet: untyped?) -> untyped
      def set_sheet_view(name, value, sheet: nil)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @sheet_views.key?(sheet_name)

        @sheet_views[sheet_name][name] = value
      end

      # Returns sheet view properties for the first (or given) sheet.
      # : (?sheet: untyped?) -> untyped
      def sheet_view(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        (@sheet_views[sheet_name] || {}).dup
      end

      # Sets a freeze pane. row: rows to freeze from top, column: columns to freeze from left.
      # : (?row: ::Integer, ?col: ::Integer?, ?column: ::Integer?, ?sheet: untyped?) -> untyped
      def set_freeze_pane(row: 0, col: nil, column: nil, sheet: nil)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @freeze_panes.key?(sheet_name)

        col_val = column || col || 0
        @freeze_panes[sheet_name] = { row: row, column: col_val, state: :frozen }
      end

      # Sets a split pane (non-frozen). x_split/y_split are in 1/20th of a point (twips).
      # top_left_cell: the cell at top-left of the bottom-right pane.
      # : (?x_split: ::Integer, ?y_split: ::Integer, ?top_left_cell: untyped?, ?sheet: untyped?) -> untyped
      def set_split_pane(x_split: 0, y_split: 0, top_left_cell: nil, sheet: nil)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @freeze_panes.key?(sheet_name)

        @freeze_panes[sheet_name] = { x_split: x_split, y_split: y_split, top_left_cell: top_left_cell, state: :split }
      end

      # Returns freeze pane settings for the first (or given) sheet.
      # : (?sheet: untyped?) -> untyped
      def freeze_pane(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        @freeze_panes[sheet_name]
      end

      # Sets the active cell selection.
      # : (untyped active_cell, ?sqref: untyped?, ?pane: untyped?, ?active_cell_id: untyped?, ?sheet: untyped?) -> untyped
      def set_selection(active_cell, sqref: nil, pane: nil, active_cell_id: nil, sheet: nil)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @selections.key?(sheet_name)

        sel = { active_cell: active_cell, sqref: sqref || active_cell }
        sel[:pane] = pane if pane
        sel[:active_cell_id] = active_cell_id if active_cell_id
        @selections[sheet_name] = sel
      end

      # Returns selection for the first (or given) sheet.
      # : (?sheet: untyped?) -> untyped
      def selection(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        @selections[sheet_name]
      end

      # Sets a print option (e.g. :grid_lines, :headings, :horizontal_centered, :vertical_centered).
      # : (untyped name, untyped value, ?sheet: untyped?) -> untyped
      def set_print_option(name, value, sheet: nil)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @print_options.key?(sheet_name)

        @print_options[sheet_name][name] = value
      end

      # Returns print options for the first (or given) sheet.
      # : (?sheet: untyped?) -> untyped
      def print_options(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        (@print_options[sheet_name] || {}).dup
      end

      # Sets page margins (all values in inches).
      # : (?left: untyped?, ?right: untyped?, ?top: untyped?, ?bottom: untyped?, ?header: untyped?, ?footer: untyped?, ?sheet: untyped?) -> untyped
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
      # : (?sheet: untyped?) -> untyped
      def page_margins(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        @page_margins[sheet_name]
      end

      # Sets a page setup property (e.g. :orientation, :paper_size, :scale, :fit_to_width, :fit_to_height).
      # : (untyped name, untyped value, ?sheet: untyped?) -> untyped
      def set_page_setup(name, value, sheet: nil)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @page_setup.key?(sheet_name)

        @page_setup[sheet_name][name] = value
      end

      # Returns page setup for the first (or given) sheet.
      # : (?sheet: untyped?) -> untyped
      def page_setup(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        (@page_setup[sheet_name] || {}).dup
      end

      # Sets header/footer text (:odd_header, :odd_footer, :even_header, :even_footer).
      # : (untyped name, untyped value, ?sheet: untyped?) -> untyped
      def set_header_footer(name, value, sheet: nil)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @header_footer.key?(sheet_name)

        @header_footer[sheet_name][name] = value
      end

      # Returns header/footer for the first (or given) sheet.
      # : (?sheet: untyped?) -> untyped
      def header_footer(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        (@header_footer[sheet_name] || {}).dup
      end

      # Adds a row break (page break before a given row number).
      # : (untyped row_num, ?sheet: untyped?) -> untyped
      def add_row_break(row_num, sheet: nil)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @row_breaks.key?(sheet_name)

        @row_breaks[sheet_name] << row_num
      end

      # Returns row breaks for the first (or given) sheet.
      # : (?sheet: untyped?) -> untyped
      def row_breaks(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        @row_breaks[sheet_name] || []
      end

      # Adds a column break (page break before a given column index, 1-based).
      # : (untyped col_index, ?sheet: untyped?) -> untyped
      def add_col_break(col_index, sheet: nil)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @col_breaks.key?(sheet_name)

        @col_breaks[sheet_name] << col_index
      end

      # Returns column breaks for the first (or given) sheet.
      # : (?sheet: untyped?) -> untyped
      def col_breaks(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        @col_breaks[sheet_name] || []
      end

      # Adds a data validation rule.
      # sqref: cell range (e.g. "A1:A100")
      # Options: type:, operator:, formula1:, formula2:, allow_blank:, show_input_message:,
      #          show_error_message:, error_style:, error_title:, error:, prompt_title:, prompt:
      # : (untyped sqref, ?sheet: untyped?, **untyped opts) -> untyped
      def add_data_validation(sqref, sheet: nil, **opts)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @data_validations.key?(sheet_name)

        @data_validations[sheet_name] << opts.merge(sqref: sqref)
      end

      # Returns data validations for the first (or given) sheet.
      # : (?sheet: untyped?) -> untyped
      def data_validations(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        @data_validations[sheet_name] || []
      end

      # Sets a data validations container option (e.g. disable_prompts, x_window, y_window).
      # : (untyped name, untyped value, ?sheet: untyped?) -> untyped
      def set_data_validations_option(name, value, sheet: nil)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @data_validations_options.key?(sheet_name)

        @data_validations_options[sheet_name][name] = value
      end

      # Adds a conditional formatting rule to the specified range.
      # Options: type (:cell_is, :expression, :color_scale, :data_bar, :icon_set),
      # operator, priority, formula/formulas, format_id, color_scale, data_bar, icon_set.
      # : (?untyped sqref, ?sheet: untyped?, **untyped opts) ?{ (ConditionalFormatBuilder) -> void } -> untyped
      def add_conditional_format(sqref = nil, sheet: nil, **opts, &)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @conditional_formats.key?(sheet_name)

        normalized_list = DslHelpers.normalize_conditional_format(sqref, **opts, &)
        normalized_list.each do |rule|
          if (rule[:dxf] || rule[:font_color] || rule[:fill_color] || rule[:bold] || rule[:italic]) && (rule[:dxf_id] || rule[:format_id]).nil?
            dxf = rule.delete(:dxf) || DslHelpers.extract_dxf_from_rule!(rule)
            did = @dxfs.index(dxf) || add_dxf(**dxf) if dxf
            rule[:dxf_id] = did
            rule[:format_id] = did
          elsif rule[:dxf_id] && rule[:format_id].nil?
            rule[:format_id] = rule[:dxf_id]
          elsif rule[:format_id] && rule[:dxf_id].nil?
            rule[:dxf_id] = rule[:format_id]
          end
          @conditional_formats[sheet_name] << rule
        end
      end

      # Returns conditional formatting rules for the first (or given) sheet.
      # : (?sheet: untyped?) -> untyped
      def conditional_formats(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        @conditional_formats[sheet_name] || []
      end

      # Adds a table definition to a sheet.
      # columns: array of column name strings.
      # : (untyped ref, columns: untyped, ?name: untyped?, ?display_name: untyped?, ?sheet: untyped?, ?totals_row_count: ::Integer, ?style: untyped?, **untyped opts) -> untyped
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
           header_row_cell_style totals_row_cell_style connection_id table_type
           auto_filter show_first_column show_last_column show_row_stripes show_column_stripes].each do |key|
          tbl[key] = opts[key] if opts.key?(key)
        end
        @tables[sheet_name] << tbl
      end

      # Returns table definitions for the first (or given) sheet.
      # : (?sheet: untyped?) -> untyped
      def tables(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        @tables[sheet_name] || []
      end

      # Enables shared string table mode (strings stored in sharedStrings.xml).
      # : () -> untyped
      def use_shared_strings!
        @use_shared_strings = true
      end

      # Sets a sheet's visibility state (:visible, :hidden, :very_hidden).
      # : (untyped sheet_name, untyped state) -> untyped
      def set_sheet_state(sheet_name, state)
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @sheets.key?(sheet_name)
        raise ArgumentError, "state must be :visible, :hidden, or :very_hidden" unless %i[visible hidden very_hidden].include?(state)

        @sheet_states[sheet_name] = state
      end

      # Returns the sheet state for a given sheet.
      # : (untyped sheet_name) -> untyped
      def sheet_state(sheet_name)
        @sheet_states[sheet_name] || :visible
      end

      # Adds a defined name. Options: sheet: (local scope), hidden: true, value: (formula or constant).
      # : (untyped name, untyped value, ?sheet: untyped?, ?hidden: bool, **untyped opts) -> untyped
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
      # : () -> untyped
      def defined_names
        @defined_names.map(&:dup)
      end

      # Sets the print area for a sheet. range should be like "A1:D20".
      # Generates the _xlnm.Print_Area defined name automatically.
      # : (untyped range, ?sheet: untyped?) -> untyped
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
      # : (?rows: untyped?, ?cols: untyped?, ?sheet: untyped?) -> untyped
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
      # : (untyped name, untyped value) -> untyped
      def set_workbook_property(name, value)
        @workbook_properties[name] = value
      end
      alias workbook_property set_workbook_property

      # Returns workbook properties hash.
      # : () -> untyped
      def workbook_properties
        @workbook_properties.dup
      end

      # Sets a file version property (e.g. :app_name, :last_edited, :lowest_edited, :rup_build, :code_name).
      # : (untyped name, untyped value) -> untyped
      def set_file_version(name, value)
        @file_version[name] = value
      end

      # Returns file version hash.
      # : () -> untyped
      def file_version
        @file_version.dup
      end

      # Sets a file sharing property (e.g. :read_only_recommended, :user_name).
      # : (untyped name, untyped value) -> untyped
      def set_file_sharing(name, value)
        @file_sharing[name] = value
      end

      # Returns file sharing hash.
      # : () -> untyped
      def file_sharing
        @file_sharing.dup
      end

      # Sets a workbook view property (e.g. :active_tab, :first_sheet).
      # : (untyped name, untyped value) -> untyped
      def set_workbook_view(name, value)
        @workbook_views[name] = value
      end

      # Returns workbook view properties hash.
      # : () -> untyped
      def workbook_views
        @workbook_views.dup
      end

      # Sets a calc property (e.g. :calc_id, :full_calc_on_load).
      # : (untyped name, untyped value) -> untyped
      def set_calc_property(name, value)
        @calc_properties[name] = value
      end

      # Returns calc properties hash.
      # : () -> untyped
      def calc_properties
        @calc_properties.dup
      end

      # Sets a file recovery property (e.g. :auto_recover, :crash_save).
      # : (untyped name, untyped value) -> untyped
      def set_file_recovery_property(name, value)
        @file_recovery_properties[name] = value
      end

      # Returns file recovery properties hash.
      # : () -> untyped
      def file_recovery_properties
        @file_recovery_properties.dup
      end

      # Sets sheet protection options.
      # Options: :password, :sheet, :objects, :scenarios, :format_cells, :format_columns,
      #   :format_rows, :insert_columns, :insert_rows, :insert_hyperlinks,
      #   :delete_columns, :delete_rows, :select_locked_cells, :sort, :auto_filter,
      #   :pivot_tables, :select_unlocked_cells, :algorithm_name, :hash_value, :salt_value, :spin_count
      # : (?sheet: untyped?, **untyped opts) -> untyped
      def set_sheet_protection(sheet: nil, **opts)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @sheets.key?(sheet_name)

        @sheet_protection[sheet_name] = opts
      end

      # Returns sheet protection settings for the given sheet.
      # : (?sheet: untyped?) -> untyped
      def sheet_protection(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        @sheet_protection[sheet_name]&.dup
      end

      # Adds a protected range to the given sheet.
      # Required: name:, sqref:
      # Optional: algorithm_name:, hash_value:, salt_value:, spin_count:, security_descriptors: []
      # : (?sheet: untyped?, **untyped opts) -> untyped
      def add_protected_range(sheet: nil, **opts)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @sheets.key?(sheet_name)
        raise ArgumentError, "name is required" unless opts[:name]
        raise ArgumentError, "sqref is required" unless opts[:sqref]

        @protected_ranges[sheet_name] << opts
      end

      # Returns protected ranges for the given sheet.
      # : (?sheet: untyped?) -> untyped
      def protected_ranges(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        (@protected_ranges[sheet_name] || []).map(&:dup)
      end

      # Adds a cell watch to the given sheet.
      # : (untyped cell_ref, ?sheet: untyped?) -> untyped
      def add_cell_watch(cell_ref, sheet: nil)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @sheets.key?(sheet_name)

        @cell_watches[sheet_name] << cell_ref
      end

      # Returns cell watches for the given sheet.
      # : (?sheet: untyped?) -> untyped
      def cell_watches(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        (@cell_watches[sheet_name] || []).dup
      end

      # Adds an ignored error entry for the given sheet.
      # Options: sqref:, eval_error:, two_digit_text_year:, number_stored_as_text:, formula:,
      #   formula_range:, unlocked_formula:, empty_cell_reference:, list_data_validation:, calculated_column:
      # : (?sheet: untyped?, **untyped opts) -> untyped
      def add_ignored_error(sheet: nil, **opts)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @sheets.key?(sheet_name)
        raise ArgumentError, "sqref is required" unless opts[:sqref]

        @ignored_errors[sheet_name] << opts
      end

      # Returns ignored errors for the given sheet.
      # : (?sheet: untyped?) -> untyped
      def ignored_errors(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        (@ignored_errors[sheet_name] || []).dup
      end

      # Sets data consolidation options for the given sheet.
      # Options: function:, start_labels:, top_labels:, link:, data_refs: [{ref:, name:, sheet:}]
      # : (?sheet: untyped?, **untyped opts) -> untyped
      def set_data_consolidate(sheet: nil, **opts)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @sheets.key?(sheet_name)

        @data_consolidate[sheet_name] = opts
      end

      # Returns data consolidation settings for the given sheet.
      # : (?sheet: untyped?) -> untyped
      def data_consolidate(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        @data_consolidate[sheet_name]&.dup
      end

      # Sets scenarios for the given sheet.
      # Options: current:, show:, sqref:, scenarios: [{name:, input_cells: [{r:, val:, ...}], ...}]
      # : (?sheet: untyped?, **untyped opts) -> untyped
      def set_scenarios(sheet: nil, **opts)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @sheets.key?(sheet_name)

        @scenarios[sheet_name] = opts
      end

      # Returns scenarios for the given sheet.
      # : (?sheet: untyped?) -> untyped
      def scenarios(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        @scenarios[sheet_name]&.dup
      end

      # Sets workbook protection options.
      # Options: :lock_structure, :lock_windows, :password, :algorithm_name, :hash_value, :salt_value, :spin_count
      # : (**untyped opts) -> untyped
      def set_workbook_protection(**opts)
        @workbook_protection = opts
      end

      # Returns workbook protection settings.
      # : () -> untyped
      def workbook_protection
        @workbook_protection&.dup
      end

      # Returns ordered sheet names.
      # : untyped
      attr_reader :sheet_order

      # Adds a raw ZIP entry to be included in the output (for pass-through retention).
      # : (untyped path, untyped content, ?content_type: untyped?) -> (nil | untyped)
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
      # : (untyped filepath) -> untyped
      def copy_entries_from(filepath)
        reader = Xlsxrb::Ooxml::Reader.new(filepath)
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
      # : (untyped file_data, ?ext: ::String, ?from_col: ::Integer, ?from_row: ::Integer, ?to_col: ::Integer, ?to_row: ::Integer, ?from_col_off: untyped?, ?from_row_off: untyped?, ?to_col_off: untyped?, ?to_row_off: untyped?, ?name: untyped?, ?description: untyped?, ?title: untyped?, ?hidden: untyped?, ?macro: untyped?, ?no_change_aspect: bool, ?no_crop: untyped?, ?line_color: untyped?, ?line_width: untyped?, ?rotation: untyped?, ?edit_as: untyped?, ?published: untyped?, ?locks_with_sheet: untyped?, ?prints_with_sheet: untyped?, ?src_rect: untyped?, ?alpha_mod_fix: untyped?, ?sheet: untyped?) -> untyped
      def insert_image(file_data, ext: "png", from_col: 0, from_row: 0, to_col: 5, to_row: 10, from_col_off: nil, from_row_off: nil, to_col_off: nil, to_row_off: nil, name: nil, description: nil, title: nil, hidden: nil, macro: nil, no_change_aspect: true, no_crop: nil, line_color: nil, line_width: nil, rotation: nil, edit_as: nil, published: nil, locks_with_sheet: nil, prints_with_sheet: nil, src_rect: nil, alpha_mod_fix: nil, sheet: nil)
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
        img[:src_rect] = src_rect if src_rect
        img[:alpha_mod_fix] = alpha_mod_fix if alpha_mod_fix
        @images[sheet_name] << img
      end

      # Returns images for the first (or given) sheet.
      # : (?sheet: untyped?) -> untyped
      def images(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        @images[sheet_name] || []
      end

      # Adds a chart to the given sheet.
      # type: :bar, :line, :pie. title: chart title string.
      # data_ref: e.g. "Sheet1!$A$1:$B$4". cat_ref/val_ref for explicit series.
      # : (?type: ::Symbol, ?title: untyped?, ?title_overlay: untyped?, ?title_font: untyped?, ?title_fill_color: untyped?, ?title_no_fill: untyped?, ?title_line_color: untyped?, ?title_line_width: untyped?, ?title_line_dash: untyped?, ?title_rotation: untyped?, ?auto_title_deleted: untyped?, ?cat_ref: untyped?, ?val_ref: untyped?, ?series: untyped?, ?legend: untyped?, ?data_labels: untyped?, ?cat_axis_title: untyped?, ?val_axis_title: untyped?, ?cat_axis_title_font: untyped?, ?cat_axis_title_fill: untyped?, ?cat_axis_title_no_fill: untyped?, ?cat_axis_title_line_color: untyped?, ?cat_axis_title_line_width: untyped?, ?cat_axis_title_line_dash: untyped?, ?cat_axis_title_rotation: untyped?, ?val_axis_title_font: untyped?, ?val_axis_title_fill: untyped?, ?val_axis_title_no_fill: untyped?, ?val_axis_title_line_color: untyped?, ?val_axis_title_line_width: untyped?, ?val_axis_title_line_dash: untyped?, ?val_axis_title_rotation: untyped?, ?cat_axis_tick_lbl_pos: untyped?, ?val_axis_tick_lbl_pos: untyped?, ?cat_axis_major_gridlines: untyped?, ?val_axis_major_gridlines: untyped?, ?cat_axis_minor_gridlines: untyped?, ?val_axis_minor_gridlines: untyped?, ?cat_axis_delete: untyped?, ?val_axis_delete: untyped?, ?cat_axis_orientation: untyped?, ?val_axis_orientation: untyped?, ?cat_axis_num_fmt: untyped?, ?val_axis_num_fmt: untyped?, ?cat_axis_major_tick_mark: untyped?, ?cat_axis_minor_tick_mark: untyped?, ?val_axis_major_tick_mark: untyped?, ?val_axis_minor_tick_mark: untyped?, ?cat_axis_crosses: untyped?, ?val_axis_crosses: untyped?, ?cat_axis_crosses_at: untyped?, ?val_axis_crosses_at: untyped?, ?cat_axis_tick_lbl_skip: untyped?, ?cat_axis_tick_mark_skip: untyped?, ?cat_axis_lbl_offset: untyped?, ?cat_axis_auto: untyped?, ?cat_axis_lbl_algn: untyped?, ?cat_axis_no_multi_lvl_lbl: untyped?, ?val_axis_cross_between: untyped?, ?val_axis_major_unit: untyped?, ?val_axis_minor_unit: untyped?, ?cat_axis_scaling_max: untyped?, ?cat_axis_scaling_min: untyped?, ?val_axis_scaling_max: untyped?, ?val_axis_scaling_min: untyped?, ?cat_axis_log_base: untyped?, ?val_axis_log_base: untyped?, ?val_axis_disp_units: untyped?, ?gap_width: untyped?, ?gap_depth: untyped?, ?overlap: untyped?, ?first_slice_ang: untyped?, ?hole_size: untyped?, ?smooth: untyped?, ?marker: untyped?, ?scatter_style: untyped?, ?radar_style: untyped?, ?bar_shape: untyped?, ?bubble_3d: untyped?, ?bubble_scale: untyped?, ?show_neg_bubbles: untyped?, ?size_represents: untyped?, ?wireframe: untyped?, ?grouping: untyped?, ?bar_dir: untyped?, ?vary_colors: untyped?, ?style: untyped?, ?rounded_corners: untyped?, ?view_3d: untyped?, ?cat_axis_pos: untyped?, ?val_axis_pos: untyped?, ?name: untyped?, ?description: untyped?, ?frame_title: untyped?, ?frame_hidden: untyped?, ?frame_macro: untyped?, ?frame_no_grp: untyped?, ?from_col: ::Integer, ?from_row: ::Integer, ?to_col: ::Integer, ?to_row: ::Integer, ?from_col_off: untyped?, ?from_row_off: untyped?, ?to_col_off: untyped?, ?to_row_off: untyped?, ?edit_as: untyped?, ?published: untyped?, ?locks_with_sheet: untyped?, ?prints_with_sheet: untyped?, ?plot_vis_only: untyped?, ?disp_blanks_as: untyped?, ?show_d_lbls_over_max: untyped?, ?data_table: untyped?, ?plot_area_fill: untyped?, ?plot_area_no_fill: untyped?, ?plot_area_line_color: untyped?, ?plot_area_line_width: untyped?, ?plot_area_line_dash: untyped?, ?plot_area_layout: untyped?, ?floor: untyped?, ?side_wall: untyped?, ?back_wall: untyped?, ?drop_lines: untyped?, ?hi_low_lines: untyped?, ?ser_lines: untyped?, ?up_down_bars: untyped?, ?band_fmts: untyped?, ?of_pie_type: untyped?, ?split_type: untyped?, ?split_pos: untyped?, ?cust_split: untyped?, ?second_pie_size: untyped?, ?cat_axis_label_rotation: untyped?, ?val_axis_label_rotation: untyped?, ?cat_axis_font: untyped?, ?val_axis_font: untyped?, ?cat_axis_line_color: untyped?, ?cat_axis_line_width: untyped?, ?cat_axis_line_dash: untyped?, ?val_axis_line_color: untyped?, ?val_axis_line_width: untyped?, ?val_axis_line_dash: untyped?, ?cat_axis_fill: untyped?, ?cat_axis_no_fill: untyped?, ?val_axis_fill: untyped?, ?val_axis_no_fill: untyped?, ?cat_axis_type: untyped?, ?cat_axis_base_time_unit: untyped?, ?cat_axis_major_time_unit: untyped?, ?cat_axis_minor_time_unit: untyped?, ?cat_axis_major_unit: untyped?, ?cat_axis_minor_unit: untyped?, ?chart_fill: untyped?, ?chart_no_fill: untyped?, ?chart_line_color: untyped?, ?chart_line_width: untyped?, ?chart_line_dash: untyped?, ?protection: untyped?, ?print_settings: untyped?, ?chart_font: untyped?, ?secondary_value_axis: untyped?, ?secondary_axis: untyped?, ?category_axis: untyped?, ?value_axis: untyped?, ?sheet: untyped?, **untyped extra_options) -> untyped
      def add_chart(type: :bar, title: nil, title_overlay: nil, title_font: nil, title_fill_color: nil, title_no_fill: nil, title_line_color: nil, title_line_width: nil, title_line_dash: nil, title_rotation: nil, auto_title_deleted: nil, cat_ref: nil, val_ref: nil, series: nil, legend: nil, data_labels: nil, cat_axis_title: nil, val_axis_title: nil, cat_axis_title_font: nil, cat_axis_title_fill: nil, cat_axis_title_no_fill: nil, cat_axis_title_line_color: nil, cat_axis_title_line_width: nil, cat_axis_title_line_dash: nil, cat_axis_title_rotation: nil, val_axis_title_font: nil, val_axis_title_fill: nil, val_axis_title_no_fill: nil, val_axis_title_line_color: nil, val_axis_title_line_width: nil, val_axis_title_line_dash: nil, val_axis_title_rotation: nil, cat_axis_tick_lbl_pos: nil, val_axis_tick_lbl_pos: nil, cat_axis_major_gridlines: nil, val_axis_major_gridlines: nil, cat_axis_minor_gridlines: nil, val_axis_minor_gridlines: nil, cat_axis_delete: nil, val_axis_delete: nil, cat_axis_orientation: nil, val_axis_orientation: nil, cat_axis_num_fmt: nil, val_axis_num_fmt: nil, cat_axis_major_tick_mark: nil, cat_axis_minor_tick_mark: nil, val_axis_major_tick_mark: nil, val_axis_minor_tick_mark: nil, cat_axis_crosses: nil, val_axis_crosses: nil, cat_axis_crosses_at: nil, val_axis_crosses_at: nil, cat_axis_tick_lbl_skip: nil, cat_axis_tick_mark_skip: nil, cat_axis_lbl_offset: nil, cat_axis_auto: nil, cat_axis_lbl_algn: nil, cat_axis_no_multi_lvl_lbl: nil, val_axis_cross_between: nil, val_axis_major_unit: nil, val_axis_minor_unit: nil, cat_axis_scaling_max: nil, cat_axis_scaling_min: nil, val_axis_scaling_max: nil, val_axis_scaling_min: nil, cat_axis_log_base: nil, val_axis_log_base: nil, val_axis_disp_units: nil, gap_width: nil, gap_depth: nil, overlap: nil, first_slice_ang: nil, hole_size: nil, smooth: nil, marker: nil, scatter_style: nil, radar_style: nil, bar_shape: nil, bubble_3d: nil, bubble_scale: nil, show_neg_bubbles: nil, size_represents: nil, wireframe: nil, grouping: nil, bar_dir: nil, vary_colors: nil, style: nil, rounded_corners: nil, view_3d: nil, cat_axis_pos: nil, val_axis_pos: nil, name: nil, description: nil, frame_title: nil, frame_hidden: nil, frame_macro: nil, frame_no_grp: nil, from_col: 0, from_row: 0, to_col: 10, to_row: 15, from_col_off: nil, from_row_off: nil, to_col_off: nil, to_row_off: nil, edit_as: nil, published: nil, locks_with_sheet: nil, prints_with_sheet: nil, plot_vis_only: nil, disp_blanks_as: nil, show_d_lbls_over_max: nil, data_table: nil, plot_area_fill: nil, plot_area_no_fill: nil, plot_area_line_color: nil, plot_area_line_width: nil, plot_area_line_dash: nil, plot_area_layout: nil, floor: nil, side_wall: nil, back_wall: nil, drop_lines: nil, hi_low_lines: nil, ser_lines: nil, up_down_bars: nil, band_fmts: nil, of_pie_type: nil, split_type: nil, split_pos: nil, cust_split: nil, second_pie_size: nil, cat_axis_label_rotation: nil, val_axis_label_rotation: nil, cat_axis_font: nil, val_axis_font: nil, cat_axis_line_color: nil, cat_axis_line_width: nil, cat_axis_line_dash: nil, val_axis_line_color: nil, val_axis_line_width: nil, val_axis_line_dash: nil, cat_axis_fill: nil, cat_axis_no_fill: nil, val_axis_fill: nil, val_axis_no_fill: nil, cat_axis_type: nil, cat_axis_base_time_unit: nil, cat_axis_major_time_unit: nil, cat_axis_minor_time_unit: nil, cat_axis_major_unit: nil, cat_axis_minor_unit: nil, chart_fill: nil, chart_no_fill: nil, chart_line_color: nil, chart_line_width: nil, chart_line_dash: nil, protection: nil, print_settings: nil, chart_font: nil, secondary_value_axis: nil, secondary_axis: nil, category_axis: nil, value_axis: nil, sheet: nil, **extra_options)
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
        chart[:title_overlay] = title_overlay unless title_overlay.nil?
        chart[:title_font] = title_font if title_font
        chart[:title_fill_color] = title_fill_color if title_fill_color
        chart[:title_no_fill] = title_no_fill unless title_no_fill.nil?
        chart[:title_line_color] = title_line_color if title_line_color
        chart[:title_line_width] = title_line_width if title_line_width
        chart[:title_line_dash] = title_line_dash if title_line_dash
        chart[:title_rotation] = title_rotation if title_rotation
        chart[:legend] = legend if legend
        chart[:data_labels] = data_labels if data_labels
        chart[:cat_axis_title] = cat_axis_title if cat_axis_title
        chart[:val_axis_title] = val_axis_title if val_axis_title
        chart[:cat_axis_title_font] = cat_axis_title_font if cat_axis_title_font
        chart[:cat_axis_title_fill] = cat_axis_title_fill if cat_axis_title_fill
        chart[:cat_axis_title_no_fill] = cat_axis_title_no_fill unless cat_axis_title_no_fill.nil?
        chart[:cat_axis_title_line_color] = cat_axis_title_line_color if cat_axis_title_line_color
        chart[:cat_axis_title_line_width] = cat_axis_title_line_width if cat_axis_title_line_width
        chart[:cat_axis_title_line_dash] = cat_axis_title_line_dash if cat_axis_title_line_dash
        chart[:cat_axis_title_rotation] = cat_axis_title_rotation if cat_axis_title_rotation
        chart[:val_axis_title_font] = val_axis_title_font if val_axis_title_font
        chart[:val_axis_title_fill] = val_axis_title_fill if val_axis_title_fill
        chart[:val_axis_title_no_fill] = val_axis_title_no_fill unless val_axis_title_no_fill.nil?
        chart[:val_axis_title_line_color] = val_axis_title_line_color if val_axis_title_line_color
        chart[:val_axis_title_line_width] = val_axis_title_line_width if val_axis_title_line_width
        chart[:val_axis_title_line_dash] = val_axis_title_line_dash if val_axis_title_line_dash
        chart[:val_axis_title_rotation] = val_axis_title_rotation if val_axis_title_rotation
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
        chart[:cat_axis_auto] = cat_axis_auto unless cat_axis_auto.nil?
        chart[:cat_axis_lbl_algn] = cat_axis_lbl_algn if cat_axis_lbl_algn
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
        chart[:band_fmts] = band_fmts if band_fmts
        chart[:of_pie_type] = of_pie_type if of_pie_type
        chart[:split_type] = split_type if split_type
        chart[:split_pos] = split_pos if split_pos
        chart[:cust_split] = cust_split if cust_split
        chart[:second_pie_size] = second_pie_size if second_pie_size
        chart[:style] = style if style
        chart[:rounded_corners] = rounded_corners unless rounded_corners.nil?
        chart[:view_3d] = view_3d if view_3d
        chart[:floor] = floor if floor
        chart[:side_wall] = side_wall if side_wall
        chart[:back_wall] = back_wall if back_wall
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
        chart[:plot_area_no_fill] = plot_area_no_fill unless plot_area_no_fill.nil?
        chart[:plot_area_line_color] = plot_area_line_color if plot_area_line_color
        chart[:plot_area_line_width] = plot_area_line_width if plot_area_line_width
        chart[:plot_area_line_dash] = plot_area_line_dash if plot_area_line_dash
        chart[:plot_area_layout] = plot_area_layout if plot_area_layout
        chart[:drop_lines] = drop_lines unless drop_lines.nil?
        chart[:hi_low_lines] = hi_low_lines unless hi_low_lines.nil?
        chart[:ser_lines] = ser_lines unless ser_lines.nil?
        chart[:up_down_bars] = up_down_bars if up_down_bars
        chart[:cat_axis_label_rotation] = cat_axis_label_rotation if cat_axis_label_rotation
        chart[:val_axis_label_rotation] = val_axis_label_rotation if val_axis_label_rotation
        chart[:cat_axis_font] = cat_axis_font if cat_axis_font
        chart[:val_axis_font] = val_axis_font if val_axis_font
        chart[:cat_axis_line_color] = cat_axis_line_color if cat_axis_line_color
        chart[:cat_axis_line_width] = cat_axis_line_width if cat_axis_line_width
        chart[:cat_axis_line_dash] = cat_axis_line_dash if cat_axis_line_dash
        chart[:val_axis_line_color] = val_axis_line_color if val_axis_line_color
        chart[:val_axis_line_width] = val_axis_line_width if val_axis_line_width
        chart[:val_axis_line_dash] = val_axis_line_dash if val_axis_line_dash
        chart[:cat_axis_fill] = cat_axis_fill if cat_axis_fill
        chart[:cat_axis_no_fill] = cat_axis_no_fill unless cat_axis_no_fill.nil?
        chart[:val_axis_fill] = val_axis_fill if val_axis_fill
        chart[:val_axis_no_fill] = val_axis_no_fill unless val_axis_no_fill.nil?
        chart[:cat_axis_type] = cat_axis_type if cat_axis_type
        chart[:cat_axis_base_time_unit] = cat_axis_base_time_unit if cat_axis_base_time_unit
        chart[:cat_axis_major_time_unit] = cat_axis_major_time_unit if cat_axis_major_time_unit
        chart[:cat_axis_minor_time_unit] = cat_axis_minor_time_unit if cat_axis_minor_time_unit
        chart[:cat_axis_major_unit] = cat_axis_major_unit if cat_axis_major_unit
        chart[:cat_axis_minor_unit] = cat_axis_minor_unit if cat_axis_minor_unit
        chart[:chart_fill] = chart_fill if chart_fill
        chart[:chart_no_fill] = chart_no_fill unless chart_no_fill.nil?
        chart[:chart_line_color] = chart_line_color if chart_line_color
        chart[:chart_line_width] = chart_line_width if chart_line_width
        chart[:chart_line_dash] = chart_line_dash if chart_line_dash
        chart[:protection] = protection if protection
        chart[:print_settings] = print_settings if print_settings
        chart[:chart_font] = chart_font if chart_font
        chart[:category_axis] = category_axis if category_axis
        chart[:value_axis] = value_axis if value_axis
        chart[:secondary_value_axis] = secondary_value_axis || secondary_axis if secondary_value_axis || secondary_axis
        chart.merge!(extra_options) if extra_options && !extra_options.empty?

        if chart[:category_axis].is_a?(Hash)
          ca = chart[:category_axis]
          chart[:cat_axis_title] ||= ca[:title]
          chart[:cat_axis_num_fmt] ||= ca[:num_fmt]
          chart[:cat_axis_pos] ||= ca[:pos]
          chart[:cat_axis_major_gridlines] = ca[:major_gridlines] if chart[:cat_axis_major_gridlines].nil? && !ca[:major_gridlines].nil?
          chart[:cat_axis_minor_gridlines] = ca[:minor_gridlines] if chart[:cat_axis_minor_gridlines].nil? && !ca[:minor_gridlines].nil?
          chart[:cat_axis_font] ||= ca[:font]
          chart[:cat_axis_line_color] ||= ca[:line_color]
          chart[:cat_axis_line_width] ||= ca[:line_width]
          chart[:cat_axis_line_dash] ||= ca[:line_dash]
          chart[:cat_axis_scaling_min] ||= ca[:min]
          chart[:cat_axis_scaling_max] ||= ca[:max]
          chart[:cat_axis_label_rotation] ||= ca[:label_rotation] || ca[:rotation]
        end

        if chart[:value_axis].is_a?(Hash)
          va = chart[:value_axis]
          chart[:val_axis_title] ||= va[:title]
          chart[:val_axis_num_fmt] ||= va[:num_fmt]
          chart[:val_axis_pos] ||= va[:pos]
          chart[:val_axis_major_gridlines] = va[:major_gridlines] if chart[:val_axis_major_gridlines].nil? && !va[:major_gridlines].nil?
          chart[:val_axis_minor_gridlines] = va[:minor_gridlines] if chart[:val_axis_minor_gridlines].nil? && !va[:minor_gridlines].nil?
          chart[:val_axis_font] ||= va[:font]
          chart[:val_axis_line_color] ||= va[:line_color]
          chart[:val_axis_line_width] ||= va[:line_width]
          chart[:val_axis_line_dash] ||= va[:line_dash]
          chart[:val_axis_scaling_min] ||= va[:min]
          chart[:val_axis_scaling_max] ||= va[:max]
          chart[:val_axis_major_unit] ||= va[:major_unit]
          chart[:val_axis_minor_unit] ||= va[:minor_unit]
          chart[:val_axis_label_rotation] ||= va[:label_rotation] || va[:rotation]
        end

        if chart[:secondary_value_axis].is_a?(Hash)
          sva = chart[:secondary_value_axis]
          chart[:sec_val_axis_title] ||= sva[:title]
          chart[:sec_val_axis_num_fmt] ||= sva[:num_fmt]
          chart[:sec_val_axis_pos] ||= sva[:pos] || "r"
          chart[:sec_val_axis_major_gridlines] = sva[:major_gridlines] unless sva[:major_gridlines].nil?
          chart[:sec_val_axis_minor_gridlines] = sva[:minor_gridlines] unless sva[:minor_gridlines].nil?
          chart[:sec_val_axis_font] ||= sva[:font]
          chart[:sec_val_axis_line_color] ||= sva[:line_color]
          chart[:sec_val_axis_line_width] ||= sva[:line_width]
          chart[:sec_val_axis_line_dash] ||= sva[:line_dash]
          chart[:sec_val_axis_scaling_min] ||= sva[:min]
          chart[:sec_val_axis_scaling_max] ||= sva[:max]
          chart[:sec_val_axis_major_unit] ||= sva[:major_unit]
          chart[:sec_val_axis_minor_unit] ||= sva[:minor_unit]
          chart[:sec_val_axis_label_rotation] ||= sva[:label_rotation] || sva[:rotation]
          chart[:sec_val_axis_crosses] ||= sva[:crosses] || "max"
          chart[:sec_val_axis_crosses_at] ||= sva[:crosses_at]
          chart[:sec_val_axis_cross_between] ||= sva[:cross_between]
          chart[:sec_val_axis_delete] = sva[:delete] unless sva[:delete].nil?
          chart[:sec_val_axis_orientation] ||= sva[:orientation]
          chart[:sec_val_axis_log_base] ||= sva[:log_base]
          chart[:sec_val_axis_disp_units] ||= sva[:disp_units]
          chart[:sec_val_axis_major_tick_mark] ||= sva[:major_tick_mark]
          chart[:sec_val_axis_minor_tick_mark] ||= sva[:minor_tick_mark]
          chart[:sec_val_axis_tick_lbl_pos] ||= sva[:tick_lbl_pos]
        end

        @charts_data[sheet_name] << chart
      end

      # Returns chart definitions for the first (or given) sheet.
      # : (?sheet: untyped?) -> untyped
      def charts(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        @charts_data[sheet_name] || []
      end

      # Adds a shape to the given sheet.
      # preset: preset geometry name (e.g. "rect", "ellipse", "roundRect").
      # text: optional text body string.
      # from_col/from_row/to_col/to_row: anchor coordinates.
      # : (?preset: ::String, ?text: untyped?, ?name: untyped?, ?description: untyped?, ?title: untyped?, ?hidden: untyped?, ?macro: untyped?, ?textlink: untyped?, ?f_locks_text: untyped?, ?no_grp: untyped?, ?no_rot: untyped?, ?fill_color: untyped?, ?fill_alpha: untyped?, ?fill_color_transforms: untyped?, ?no_fill: untyped?, ?gradient_fill: untyped?, ?pattern_fill: untyped?, ?line_color: untyped?, ?line_alpha: untyped?, ?line_color_transforms: untyped?, ?line_width: untyped?, ?no_line: untyped?, ?line_dash: untyped?, ?line_custom_dash: untyped?, ?line_cap: untyped?, ?line_align: untyped?, ?line_compound: untyped?, ?line_join: untyped?, ?line_miter_limit: untyped?, ?head_end: untyped?, ?tail_end: untyped?, ?rotation: untyped?, ?text_wrap: untyped?, ?text_anchor: untyped?, ?text_vert_overflow: untyped?, ?text_horz_overflow: untyped?, ?text_spc_first_last_para: untyped?, ?text_num_col: untyped?, ?text_spc_col: untyped?, ?text_rtl_col: untyped?, ?text_from_word_art: untyped?, ?text_upright: untyped?, ?text_compat_ln_spc: untyped?, ?text_force_aa: untyped?, ?text_warp: untyped?, ?text_vertical: untyped?, ?text_insets: untyped?, ?text_rot: untyped?, ?adjust_values: untyped?, ?text_font: untyped?, ?text_end_para_rpr: untyped?, ?text_def_rpr: untyped?, ?text_align: untyped?, ?text_font_align: untyped?, ?text_def_tab_sz: untyped?, ?text_indent: untyped?, ?text_anchor_ctr: untyped?, ?text_spacing: untyped?, ?text_rtl: untyped?, ?text_ea_ln_brk: untyped?, ?text_latin_ln_brk: untyped?, ?text_hanging_punct: untyped?, ?text_tab_stops: untyped?, ?text_bullet: untyped?, ?text_level: untyped?, ?text_paragraphs: untyped?, ?autofit: untyped?, ?outer_shadow: untyped?, ?inner_shadow: untyped?, ?glow: untyped?, ?soft_edge: untyped?, ?reflection: untyped?, ?blur: untyped?, ?from_col: ::Integer, ?from_row: ::Integer, ?to_col: ::Integer, ?to_row: ::Integer, ?from_col_off: untyped?, ?from_row_off: untyped?, ?to_col_off: untyped?, ?to_row_off: untyped?, ?edit_as: untyped?, ?published: untyped?, ?locks_with_sheet: untyped?, ?prints_with_sheet: untyped?, ?sheet: untyped?) -> untyped
      def add_shape(preset: "rect", text: nil, name: nil, description: nil, title: nil, hidden: nil, macro: nil, textlink: nil, f_locks_text: nil, no_grp: nil, no_rot: nil, fill_color: nil, fill_alpha: nil, fill_color_transforms: nil, no_fill: nil, gradient_fill: nil, pattern_fill: nil, line_color: nil, line_alpha: nil, line_color_transforms: nil, line_width: nil, no_line: nil, line_dash: nil, line_custom_dash: nil, line_cap: nil, line_align: nil, line_compound: nil, line_join: nil, line_miter_limit: nil, head_end: nil, tail_end: nil, rotation: nil, text_wrap: nil, text_anchor: nil, text_vert_overflow: nil, text_horz_overflow: nil, text_spc_first_last_para: nil, text_num_col: nil, text_spc_col: nil, text_rtl_col: nil, text_from_word_art: nil, text_upright: nil, text_compat_ln_spc: nil, text_force_aa: nil, text_warp: nil, text_vertical: nil, text_insets: nil, text_rot: nil, adjust_values: nil, text_font: nil, text_end_para_rpr: nil, text_def_rpr: nil, text_align: nil, text_font_align: nil, text_def_tab_sz: nil, text_indent: nil, text_anchor_ctr: nil, text_spacing: nil, text_rtl: nil, text_ea_ln_brk: nil, text_latin_ln_brk: nil, text_hanging_punct: nil, text_tab_stops: nil, text_bullet: nil, text_level: nil, text_paragraphs: nil, autofit: nil, outer_shadow: nil, inner_shadow: nil, glow: nil, soft_edge: nil, reflection: nil, blur: nil, from_col: 0, from_row: 0, to_col: 5, to_row: 5, from_col_off: nil, from_row_off: nil, to_col_off: nil, to_row_off: nil, edit_as: nil, published: nil, locks_with_sheet: nil, prints_with_sheet: nil, sheet: nil)
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
        shape[:fill_alpha] = fill_alpha if fill_alpha
        shape[:fill_color_transforms] = fill_color_transforms if fill_color_transforms
        shape[:no_fill] = no_fill unless no_fill.nil?
        shape[:line_color] = line_color if line_color
        shape[:line_alpha] = line_alpha if line_alpha
        shape[:line_color_transforms] = line_color_transforms if line_color_transforms
        shape[:line_width] = line_width if line_width
        shape[:line_dash] = line_dash if line_dash
        shape[:line_custom_dash] = line_custom_dash if line_custom_dash
        shape[:line_cap] = line_cap if line_cap
        shape[:line_align] = line_align if line_align
        shape[:line_compound] = line_compound if line_compound
        shape[:line_join] = line_join if line_join
        shape[:line_miter_limit] = line_miter_limit if line_miter_limit
        shape[:head_end] = head_end if head_end
        shape[:tail_end] = tail_end if tail_end
        shape[:no_line] = no_line unless no_line.nil?
        shape[:text_wrap] = text_wrap if text_wrap
        shape[:text_anchor] = text_anchor if text_anchor
        shape[:text_vert_overflow] = text_vert_overflow if text_vert_overflow
        shape[:text_horz_overflow] = text_horz_overflow if text_horz_overflow
        shape[:text_spc_first_last_para] = text_spc_first_last_para unless text_spc_first_last_para.nil?
        shape[:text_num_col] = text_num_col if text_num_col
        shape[:text_spc_col] = text_spc_col if text_spc_col
        shape[:text_rtl_col] = text_rtl_col unless text_rtl_col.nil?
        shape[:text_from_word_art] = text_from_word_art unless text_from_word_art.nil?
        shape[:text_upright] = text_upright unless text_upright.nil?
        shape[:text_compat_ln_spc] = text_compat_ln_spc unless text_compat_ln_spc.nil?
        shape[:text_force_aa] = text_force_aa unless text_force_aa.nil?
        shape[:text_warp] = text_warp if text_warp
        shape[:text_vertical] = text_vertical if text_vertical
        shape[:text_insets] = text_insets if text_insets
        shape[:text_rot] = text_rot if text_rot
        shape[:rotation] = rotation if rotation
        shape[:adjust_values] = adjust_values if adjust_values
        shape[:text_font] = text_font if text_font
        shape[:text_end_para_rpr] = text_end_para_rpr if text_end_para_rpr
        shape[:text_def_rpr] = text_def_rpr if text_def_rpr
        shape[:text_align] = text_align if text_align
        shape[:text_font_align] = text_font_align if text_font_align
        shape[:text_def_tab_sz] = text_def_tab_sz if text_def_tab_sz
        shape[:text_indent] = text_indent if text_indent
        shape[:text_anchor_ctr] = text_anchor_ctr unless text_anchor_ctr.nil?
        shape[:text_spacing] = text_spacing if text_spacing
        shape[:text_rtl] = text_rtl unless text_rtl.nil?
        shape[:text_ea_ln_brk] = text_ea_ln_brk unless text_ea_ln_brk.nil?
        shape[:text_latin_ln_brk] = text_latin_ln_brk unless text_latin_ln_brk.nil?
        shape[:text_hanging_punct] = text_hanging_punct unless text_hanging_punct.nil?
        shape[:text_tab_stops] = text_tab_stops if text_tab_stops
        shape[:text_bullet] = text_bullet if text_bullet
        shape[:text_level] = text_level if text_level
        shape[:text_paragraphs] = text_paragraphs if text_paragraphs
        shape[:autofit] = autofit if autofit
        shape[:outer_shadow] = outer_shadow if outer_shadow
        shape[:inner_shadow] = inner_shadow if inner_shadow
        shape[:glow] = glow if glow
        shape[:soft_edge] = soft_edge if soft_edge
        shape[:reflection] = reflection if reflection
        shape[:blur] = blur if blur
        shape[:gradient_fill] = gradient_fill if gradient_fill
        shape[:pattern_fill] = pattern_fill if pattern_fill
        shape[:edit_as] = edit_as if edit_as
        shape[:published] = published unless published.nil?
        shape[:locks_with_sheet] = locks_with_sheet unless locks_with_sheet.nil?
        shape[:prints_with_sheet] = prints_with_sheet unless prints_with_sheet.nil?
        @shapes_data[sheet_name] << shape
      end

      # Returns shape definitions for the first (or given) sheet.
      # : (?sheet: untyped?) -> untyped
      def shapes(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        @shapes_data[sheet_name] || []
      end

      # Adds a comment on a cell.
      # : (untyped cell_address, untyped text, ?author: ::String, ?sheet: untyped?, ?guid: untyped?, ?shape_id: untyped?, ?visible: bool, ?anchor: Array[Integer]?) -> untyped
      def add_comment(cell_address, text, author: "Author", sheet: nil, guid: nil, shape_id: nil, visible: false, anchor: nil)
        validate_cell_address!(cell_address)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @comments_data.key?(sheet_name)

        entry = { ref: cell_address, text: text, author: author }
        entry[:guid] = guid if guid
        entry[:shape_id] = shape_id if shape_id
        entry[:visible] = true if visible
        entry[:anchor] = anchor if anchor
        @comments_data[sheet_name] << entry
      end

      # Returns comment definitions for the first (or given) sheet.
      # : (?sheet: untyped?) -> untyped
      def comments(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        @comments_data[sheet_name] || []
      end

      # Adds a threaded comment to a cell in the given sheet.
      # : (String | Integer cell_address, String text, ?author: untyped, ?timestamp: untyped, ?id: String?, ?parent_id: String?, ?done: bool?, ?user_id: String?, ?provider_id: String?, ?sheet: untyped?) -> Hash[Symbol, untyped]
      def add_threaded_comment(cell_address, text, author: "Author", timestamp: nil, id: nil, parent_id: nil, done: false, user_id: nil, provider_id: nil, sheet: nil)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @threaded_comments_data.key?(sheet_name)

        person = if author.is_a?(Hash)
                   add_person(author[:display_name] || author[:name] || "Author", id: author[:id], user_id: author[:user_id], provider_id: author[:provider_id])
                 else
                   add_person(author, user_id: user_id, provider_id: provider_id)
                 end

        dt = case timestamp
             when Time, DateTime
               timestamp.to_time.utc.strftime("%Y-%m-%dT%H:%M:%SZ")
             when String
               timestamp
             else
               Time.now.utc.strftime("%Y-%m-%dT%H:%M:%SZ")
             end

        comment_id = format_guid(id)
        parent_guid = parent_id ? format_guid(parent_id) : nil

        entry = {
          cell: cell_address.to_s.upcase,
          ref: cell_address.to_s.upcase,
          text: text.to_s,
          author: person[:display_name],
          person_id: person[:id],
          id: comment_id,
          parent_id: parent_guid,
          timestamp: dt,
          dt: dt,
          done: done ? true : false
        }.compact
        @threaded_comments_data[sheet_name] << entry
        entry
      end
      alias threaded_comment add_threaded_comment

      # Returns threaded comments for the first (or given) sheet.
      # : (?sheet: untyped?) -> Array[Hash[Symbol, untyped]]
      def threaded_comments(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        @threaded_comments_data[sheet_name] || []
      end

      # Registers a person in the workbook person list.
      # : (String name, ?id: String?, ?user_id: String?, ?provider_id: String?) -> Hash[Symbol, untyped]
      def add_person(name, id: nil, user_id: nil, provider_id: nil)
        person_name = name.to_s
        person_id = id ? format_guid(id) : nil
        existing = @persons_data.find do |p|
          (person_id && p[:id] == person_id) || p[:display_name] == person_name
        end
        return existing if existing

        person = {
          id: person_id || format_guid(nil),
          display_name: person_name,
          user_id: user_id || "AD:#{person_name}",
          provider_id: provider_id || "AD"
        }
        @persons_data << person
        person
      end

      # Returns the workbook person list.
      # : () -> Array[Hash[Symbol, untyped]]
      def persons
        @persons_data
      end
      alias person_list persons

      # Adds a pivot table to the given sheet.
      # source_ref: data source range (e.g. "Sheet1!A1:C4"), or nil with source_name: for a table source.
      # row_fields: array of 0-based field indices, or field names listed in field_names, for row axis.
      # data_fields: array of { fld:, name:, subtotal: } hashes ({ field: } takes a field name or index
      #   in place of fld:), or field names.
      # col_fields: array of 0-based field indices, or field names, for column axis.
      # field_names: array of field name strings (for cache definition).
      # items: hash mapping field index to array of item values.
      # : (untyped source_ref, row_fields: untyped, data_fields: untyped, ?col_fields: untyped, ?column_fields: untyped, ?dest_ref: untyped, ?destination: untyped, ?name: untyped?, ?field_names: untyped?, ?items: untyped?, ?sheet: untyped?, **untyped opts) -> untyped
      def add_pivot_table(source_ref, row_fields:, data_fields:, col_fields: nil, column_fields: nil, dest_ref: nil, destination: nil, name: nil, field_names: nil, items: nil, sheet: nil, **opts)
        sheet_name = sheet || @sheet_order.first
        raise ArgumentError, "unknown sheet: #{sheet_name}" unless @pivot_tables_data.key?(sheet_name)

        pt_name = name || "PivotTable#{@pivot_tables_data.values.flatten.size + 1}"
        row_fields = row_fields.map { |f| resolve_pivot_field(f, field_names) }
        c_fields = (column_fields || col_fields || []).map { |f| resolve_pivot_field(f, field_names) }
        data_fields = data_fields.map { |df| resolve_pivot_data_field(df, field_names) }
        target_dest = destination || dest_ref || "E1"
        @pivot_tables_data[sheet_name] << {
          name: pt_name, source_ref: source_ref,
          row_fields: row_fields, col_fields: c_fields, column_fields: c_fields,
          data_fields: data_fields, dest_ref: target_dest, destination: target_dest,
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
          first_header_row: opts[:first_header_row], first_data_row: opts[:first_data_row], first_data_col: opts[:first_data_col],
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

      # Resolves a pivot table field given as a 0-based index or as a name in field_names.
      # : (untyped field, untyped field_names) -> Integer
      def resolve_pivot_field(field, field_names)
        return field if field.is_a?(Integer)
        raise ArgumentError, "pivot table field #{field.inspect} is a name; pass field_names: to resolve it" unless field_names

        index = field_names.index(field.to_s)
        raise ArgumentError, "pivot table field #{field.inspect} is not in field_names #{field_names.inspect}" unless index

        index
      end
      private :resolve_pivot_field

      # Normalizes a pivot table data field to a Hash with an Integer fld.
      # : (untyped data_field, untyped field_names) -> Hash[Symbol, untyped]
      def resolve_pivot_data_field(data_field, field_names)
        data_field = { field: data_field } unless data_field.is_a?(Hash)
        return data_field if data_field[:fld].is_a?(Integer) && !data_field.key?(:field)

        field = data_field.key?(:field) ? data_field[:field] : data_field[:fld]
        raise ArgumentError, "pivot table data field #{data_field.inspect} needs fld: or field:" if field.nil?

        fld = resolve_pivot_field(field, field_names)
        subtotal = (data_field[:subtotal] || "sum").to_s
        caption = data_field[:name] || "#{subtotal[0].upcase}#{subtotal[1..]} of #{field_names ? field_names[fld] : "Field#{fld + 1}"}"
        data_field.except(:field).merge(fld: fld, name: caption, subtotal: subtotal)
      end
      private :resolve_pivot_data_field

      # Returns pivot table definitions for the first (or given) sheet.
      # : (?sheet: untyped?) -> untyped
      def pivot_tables(sheet: nil)
        sheet_name = sheet || @sheet_order.first
        @pivot_tables_data[sheet_name] || []
      end

      # Adds an external link reference to another workbook.
      # target: path or URI to the external workbook (e.g. "Book2.xlsx").
      # sheet_names: array of sheet name strings in the external workbook.
      # : (target: untyped, ?sheet_names: untyped) -> untyped
      def add_external_link(target:, sheet_names: [])
        @external_links << { target: target, sheet_names: sheet_names }
      end

      # Returns external link definitions.
      # : () -> untyped
      def external_links
        @external_links.dup
      end

      # Enables macro preservation mode. Required when copy_entries_from loads a .xlsm file.
      # : () -> untyped
      def preserve_macros!
        @preserve_macros = true
      end

      # Returns whether macro preservation is enabled.
      # : () -> untyped
      def preserve_macros?
        @preserve_macros
      end

      # Writes the workbook as an XLSX file to the given path.
      # : (untyped filepath) -> untyped
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
          sheet_threaded_comments = @threaded_comments_data[sheet_name] || []
          sheet_pivots = @pivot_tables_data[sheet_name] || []
          has_drawing = sheet_images.any? || sheet_charts.any? || sheet_shapes.any?
          has_comments = sheet_comments.any?
          has_tc = sheet_threaded_comments.any?

          # Pre-increment counters so rels reference correct paths.
          sheet_drawing_idx = has_drawing ? (@drawing_count += 1) : nil
          sheet_comment_idx = has_comments ? (@comment_count += 1) : nil
          sheet_tc_idx = has_tc ? (@threaded_comment_count += 1) : nil
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
            sheet_pivot_start, sheet_pivots.size,
            sheet_tc_idx
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

          # Generate threaded comments XML.
          entries["xl/threadedComments/threadedComment#{sheet_tc_idx}.xml"] = generate_threaded_comments_xml(sheet_threaded_comments) if has_tc

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

        # Generate persons entry if any persons registered.
        entries["xl/persons/person.xml"] = generate_persons_xml(@persons_data) if @persons_data.any?

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
        contains_errors: "containsErrors",
        not_contains_errors: "notContainsErrors",
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
        surface3d: "surface3DChart",
        stock: "stockChart", bubble: "bubbleChart",
        of_pie: "ofPieChart",
        combo: "comboChart", column: "barChart"
      }.freeze

      # The elements of EG_ColorTransform in DrawingML. A transform's type is written as an
      # element name, so only these are accepted.
      COLOR_TRANSFORMS = %w[
        tint shade comp inv gray alpha alphaOff alphaMod hue hueOff hueMod sat satOff satMod
        lum lumOff lumMod red redOff redMod green greenOff greenMod blue blueOff blueMod gamma invGamma
      ].freeze

      NO_AXIS_CHARTS = %w[pieChart doughnutChart pie3DChart ofPieChart].freeze
      THREE_AXIS_CHARTS = %w[surface3DChart].freeze
      GROUPING_CHARTS = %w[barChart lineChart areaChart bar3DChart line3DChart area3DChart].freeze

      private

      # : (?::Hash[untyped, untyped] all_entries) -> ::String
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

        # Threaded comments.
        (1..@threaded_comment_count.to_i).each do |tc|
          overrides["/xl/threadedComments/threadedComment#{tc}.xml"] = "application/vnd.ms-excel.threadedcomments+xml"
        end

        # Persons.
        overrides["/xl/persons/person.xml"] = "application/vnd.ms-excel.person+xml" if @persons_data.any?

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

        XmlBuilder.build(declaration: :plain) do |xml|
          xml.tag("Types", { xmlns: CT_NS }) do
            defaults.each { |ext, ct| xml.empty_tag("Default", { Extension: ext || "", ContentType: ct || "" }) }
            overrides.each { |pn, ct| xml.empty_tag("Override", { PartName: pn || "", ContentType: ct || "" }) }
          end
        end
      end

      # : () -> ::String
      def generate_rels_root
        rid_counter = 1
        XmlBuilder.build(declaration: :plain) do |xml|
          xml.tag("Relationships", { xmlns: REL_NS }) do
            xml.empty_tag("Relationship", { Id: "rId#{rid_counter}", Type: "#{DOC_REL_NS}/officeDocument", Target: "xl/workbook.xml" })
            unless @core_properties.empty?
              rid_counter += 1
              xml.empty_tag("Relationship", { Id: "rId#{rid_counter}", Type: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", Target: "docProps/core.xml" })
            end
            if @app_properties.any?
              rid_counter += 1
              xml.empty_tag("Relationship", { Id: "rId#{rid_counter}", Type: "#{DOC_REL_NS}/extended-properties", Target: "docProps/app.xml" })
            end
            unless @custom_properties.empty?
              rid_counter += 1
              xml.empty_tag("Relationship", { Id: "rId#{rid_counter}", Type: "#{DOC_REL_NS}/custom-properties", Target: "docProps/custom.xml" })
            end
          end
        end
      end

      # : (?::Integer pivot_cache_count) -> ::String
      def generate_workbook_xml(pivot_cache_count = 0)
        XmlBuilder.build(declaration: :plain) do |xml|
          xml.tag("workbook", { xmlns: SSML_NS, "xmlns:r": DOC_REL_NS, conformance: @workbook_properties[:conformance] }) do
            write_workbook_file_version_xml(xml)
            write_workbook_file_sharing_xml(xml)
            write_workbook_pr_xml(xml)
            write_workbook_views_xml(xml)

            xml.tag("sheets") do
              @sheet_order.each_with_index do |name, i|
                state_attr = case @sheet_states[name]
                             when :hidden then "hidden"
                             when :very_hidden then "veryHidden"
                             end
                xml.empty_tag("sheet", { name: name || "", sheetId: i + 1, state: state_attr, "r:id": "rId#{i + 1}" })
              end
            end

            write_workbook_protection_xml(xml) if @workbook_protection

            unless @external_links.empty?
              # rId for external links: after sheets + styles + optional SST + pivot caches
              el_rid_base = @sheet_order.size + 1 + (@use_shared_strings ? 1 : 0) + 1 + pivot_cache_count
              xml.tag("externalReferences") do
                @external_links.each_with_index do |_, idx|
                  xml.empty_tag("externalReference", { "r:id": "rId#{el_rid_base + idx}" })
                end
              end
            end

            unless @defined_names.empty?
              xml.tag("definedNames") do
                @defined_names.each do |dn|
                  attrs = {
                    name: dn[:name] || "",
                    localSheetId: dn[:local_sheet_id],
                    hidden: (1 if dn[:hidden]),
                    comment: dn[:comment],
                    description: dn[:description],
                    function: (1 if dn[:function]),
                    vbProcedure: (1 if dn[:vb_procedure]),
                    xlm: (1 if dn[:xlm]),
                    shortcutKey: dn[:shortcut_key],
                    publishToServer: (1 if dn[:publish_to_server]),
                    workbookParameter: (1 if dn[:workbook_parameter]),
                    functionGroupId: dn[:function_group_id],
                    customMenu: dn[:custom_menu],
                    help: dn[:help],
                    statusBar: dn[:status_bar]
                  }
                  xml.text_tag("definedName", dn[:value] || "", attrs)
                end
              end
            end

            write_workbook_calc_pr_xml(xml)

            if pivot_cache_count.positive?
              # rId layout: sheets(1..N), styles(N+1), optional SST(N+2), then pivot caches
              pivot_rid_base = @sheet_order.size + 1 + (@use_shared_strings ? 1 : 0) + 1
              xml.tag("pivotCaches") do
                pivot_cache_count.times do |ci|
                  xml.empty_tag("pivotCache", { cacheId: ci + 1, "r:id": "rId#{pivot_rid_base + ci}" })
                end
              end
            end

            unless @file_recovery_properties.empty?
              frp = @file_recovery_properties
              frp_attrs = {
                autoRecover: xml_bool_attr(frp[:auto_recover]),
                crashSave: xml_bool_attr(frp[:crash_save]),
                dataExtractLoad: xml_bool_attr(frp[:data_extract_load]),
                repairLoad: xml_bool_attr(frp[:repair_load])
              }.compact
              xml.empty_tag("fileRecoveryPr", frp_attrs) unless frp_attrs.empty?
            end
          end
        end
      end

      # : (?has_calc_chain: bool) -> ::String
      def generate_workbook_rels(has_calc_chain: false)
        XmlBuilder.build(declaration: :plain) do |xml|
          xml.tag("Relationships", { xmlns: REL_NS }) do
            @sheet_order.each_with_index do |_, i|
              xml.empty_tag("Relationship", { Id: "rId#{i + 1}", Type: "#{DOC_REL_NS}/worksheet", Target: "worksheets/sheet#{i + 1}.xml" })
            end
            next_rid = @sheet_order.size + 1
            xml.empty_tag("Relationship", { Id: "rId#{next_rid}", Type: "#{DOC_REL_NS}/styles", Target: "styles.xml" })
            next_rid += 1
            if @use_shared_strings
              xml.empty_tag("Relationship", { Id: "rId#{next_rid}", Type: "#{DOC_REL_NS}/sharedStrings", Target: "sharedStrings.xml" })
              next_rid += 1
            end
            (1..@pivot_cache_count.to_i).each do |c|
              xml.empty_tag("Relationship", { Id: "rId#{next_rid}", Type: "#{DOC_REL_NS}/pivotCacheDefinition", Target: "pivotCache/pivotCacheDefinition#{c}.xml" })
              next_rid += 1
            end
            @external_links.each_with_index do |_, idx|
              xml.empty_tag("Relationship", { Id: "rId#{next_rid}", Type: "#{DOC_REL_NS}/externalLink", Target: "externalLinks/externalLink#{idx + 1}.xml" })
              next_rid += 1
            end
            if has_calc_chain
              xml.empty_tag("Relationship", { Id: "rId#{next_rid}", Type: "#{DOC_REL_NS}/calcChain", Target: "calcChain.xml" })
              next_rid += 1
            end
            if @persons_data.any?
              xml.empty_tag(
                "Relationship",
                {
                  Id: "rId#{next_rid}",
                  Type: "http://schemas.microsoft.com/office/2017/10/relationships/person",
                  Target: "persons/person.xml"
                }
              )
            end
          end
        end
      end

      # : () -> (nil | ::String)
      def generate_calc_chain_xml
        chain_entries = []
        @sheet_order.each_with_index do |sheet_name, i|
          @sheets[sheet_name].each do |address, value|
            chain_entries << { ref: address, sheet_id: i + 1 } if value.is_a?(Xlsxrb::Elements::Formula)
          end
        end
        return nil if chain_entries.empty?

        XmlBuilder.build(declaration: :plain) do |xml|
          xml.tag("calcChain", { xmlns: SSML_NS }) do
            chain_entries.each do |entry|
              xml.empty_tag("c", { r: entry[:ref] || "", i: entry[:sheet_id] || "" })
            end
          end
        end
      end

      # : (untyped sheet_cells, untyped sheet_col_widths, untyped sheet_col_attrs, untyped sheet_row_attrs, untyped sheet_auto_filter, untyped sheet_filter_cols, untyped sheet_sort, untyped sheet_merge_cells, untyped sheet_hyperlinks, untyped sheet_cell_styles, untyped sheet_props, untyped sheet_fmt, untyped sheet_sv, untyped sheet_fp, untyped sheet_sel, untyped sheet_po, untyped sheet_pm, untyped sheet_ps, untyped sheet_hf, untyped sheet_rb, untyped sheet_cb, untyped sheet_dv, untyped sheet_cf, ?untyped? sst, ?untyped sheet_tables, ?::Integer hyperlink_count, ?has_drawing: bool, ?has_comments: bool, ?sheet_prot: untyped?, ?vml_rid: untyped?, ?phonetic_pr: untyped?, ?dv_options: ::Hash[untyped, untyped], ?prot_ranges: untyped, ?cell_watches: untyped, ?ignored_errors: untyped, ?data_consol: untyped?, ?sheet_scenarios: untyped?, ?cell_phonetic: ::Hash[untyped, untyped]) -> ::String
      def generate_worksheet_xml(sheet_cells, sheet_col_widths, sheet_col_attrs, sheet_row_attrs, sheet_auto_filter, sheet_filter_cols, sheet_sort, sheet_merge_cells, sheet_hyperlinks, sheet_cell_styles, sheet_props, sheet_fmt, sheet_sv, sheet_fp, sheet_sel, sheet_po, sheet_pm, sheet_ps, sheet_hf, sheet_rb, sheet_cb, sheet_dv, sheet_cf, sst = nil, sheet_tables = [], hyperlink_count = 0, has_drawing: false, has_comments: false, sheet_prot: nil, vml_rid: nil, phonetic_pr: nil, dv_options: {}, prot_ranges: [], cell_watches: [], ignored_errors: [], data_consol: nil, sheet_scenarios: nil, cell_phonetic: {})
        needs_r_ns = !sheet_hyperlinks.empty? || sheet_tables.any? || has_drawing || has_comments
        XmlBuilder.build(declaration: :plain) do |xml|
          xml.tag("worksheet", { xmlns: SSML_NS, "xmlns:r": (DOC_REL_NS if needs_r_ns) }) do
            write_sheet_pr_xml(xml, sheet_props) unless sheet_props.empty?
            xml.empty_tag("dimension", { ref: compute_dimension(sheet_cells) }) unless sheet_cells.empty?
            write_sheet_views_xml(xml, sheet_sv, sheet_fp, sheet_sel) if !sheet_sv.empty? || sheet_fp || sheet_sel

            unless sheet_fmt.empty?
              xml.empty_tag(
                "sheetFormatPr",
                {
                  defaultRowHeight: sheet_fmt[:default_row_height],
                  defaultColWidth: sheet_fmt[:default_col_width],
                  baseColWidth: sheet_fmt[:base_col_width],
                  outlineLevelRow: sheet_fmt[:outline_level_row],
                  outlineLevelCol: sheet_fmt[:outline_level_col],
                  customHeight: (1 if sheet_fmt[:custom_height]),
                  zeroHeight: (1 if sheet_fmt[:zero_height]),
                  thickTop: (1 if sheet_fmt[:thick_top]),
                  thickBottom: (1 if sheet_fmt[:thick_bottom])
                }
              )
            end

            all_cols = sheet_col_widths.keys | sheet_col_attrs.keys
            unless all_cols.empty?
              xml.tag("cols") do
                all_cols.sort_by { |col| column_letter_to_index(col) }.each do |col_letter|
                  idx = column_letter_to_index(col_letter)
                  width = sheet_col_widths[col_letter]
                  ca = sheet_col_attrs[col_letter] || {}
                  xml.empty_tag(
                    "col",
                    {
                      min: idx,
                      max: idx,
                      width: width,
                      customWidth: (1 if width),
                      hidden: (1 if ca[:hidden]),
                      bestFit: (1 if ca[:best_fit]),
                      outlineLevel: ca[:outline_level],
                      collapsed: (1 if ca[:collapsed]),
                      style: ca[:style],
                      phonetic: (1 if ca[:phonetic])
                    }
                  )
                end
              end
            end

            write_sheet_data_xml(xml, sheet_cells, sheet_row_attrs, sheet_cell_styles, sst, cell_phonetic)
            xml.empty_tag("sheetCalcPr", { fullCalcOnLoad: 1 }) if sheet_props[:full_calc_on_load]
            write_sheet_protection_xml(xml, sheet_prot) if sheet_prot
            write_protected_ranges_xml(xml, prot_ranges) unless prot_ranges.empty?
            write_scenarios_xml(xml, sheet_scenarios) if sheet_scenarios

            if sheet_auto_filter
              if sheet_filter_cols.empty?
                xml.empty_tag("autoFilter", { ref: sheet_auto_filter })
              else
                xml.tag("autoFilter", { ref: sheet_auto_filter }) do
                  sheet_filter_cols.sort.each do |col_id, filter|
                    xml.tag(
                      "filterColumn",
                      {
                        colId: col_id,
                        hiddenButton: (1 if filter[:hidden_button]),
                        showButton: (0 if filter[:show_button] == false)
                      }
                    ) do
                      write_filter_xml(xml, filter)
                    end
                  end
                end
              end
            end

            if sheet_sort
              xml.tag(
                "sortState",
                {
                  ref: sheet_sort[:ref] || "",
                  columnSort: (1 if sheet_sort[:column_sort]),
                  caseSensitive: (1 if sheet_sort[:case_sensitive]),
                  sortMethod: sheet_sort[:sort_method]
                }
              ) do
                sheet_sort[:sort_conditions].each do |sc|
                  xml.empty_tag(
                    "sortCondition",
                    {
                      ref: sc[:ref] || "",
                      descending: (1 if sc[:descending]),
                      sortBy: sc[:sort_by],
                      customList: sc[:custom_list],
                      dxfId: sc[:dxf_id],
                      iconSet: sc[:icon_set],
                      iconId: sc[:icon_id]
                    }
                  )
                end
              end
            end

            write_data_consolidate_xml(xml, data_consol) if data_consol

            unless sheet_merge_cells.empty?
              xml.tag("mergeCells", { count: sheet_merge_cells.size }) do
                sheet_merge_cells.each { |ref| xml.empty_tag("mergeCell", { ref: ref || "" }) }
              end
            end

            if phonetic_pr
              xml.empty_tag(
                "phoneticPr",
                {
                  fontId: phonetic_pr[:font_id],
                  type: phonetic_pr[:type],
                  alignment: phonetic_pr[:alignment]
                }
              )
            end

            unless sheet_hyperlinks.empty?
              xml.tag("hyperlinks") do
                ext_rid = 0
                sheet_hyperlinks.each do |(cell_ref, link)|
                  rid = if link[:url]
                          ext_rid += 1
                          "rId#{ext_rid}"
                        end
                  xml.empty_tag(
                    "hyperlink",
                    {
                      ref: cell_ref || "",
                      "r:id": rid,
                      display: link[:display],
                      tooltip: link[:tooltip],
                      location: link[:location]
                    }
                  )
                end
              end
            end

            unless sheet_cf.empty?
              sheet_cf.group_by { |cf| cf[:sqref] }.each do |sqref, rules|
                xml.tag("conditionalFormatting", { sqref: sqref || "", pivot: (1 if rules.any? { |r| r[:pivot] }) }) do
                  rules.each { |cf| write_cf_rule(xml, cf) }
                end
              end
            end

            write_data_validations_xml(xml, sheet_dv, dv_options) unless sheet_dv.empty?
            write_print_and_page_xml(xml, sheet_po, sheet_pm, sheet_ps, sheet_hf)
            write_page_breaks_xml(xml, "rowBreaks", sheet_rb, default_max: 16_383) unless sheet_rb.empty?
            write_page_breaks_xml(xml, "colBreaks", sheet_cb, default_max: 1_048_575) unless sheet_cb.empty?

            unless cell_watches.empty?
              xml.tag("cellWatches") do
                cell_watches.each { |r| xml.empty_tag("cellWatch", { r: r || "" }) }
              end
            end

            unless ignored_errors.empty?
              xml.tag("ignoredErrors") do
                ignored_errors.each do |ie|
                  xml.empty_tag(
                    "ignoredError",
                    {
                      sqref: ie[:sqref] || "",
                      evalError: (1 if ie[:eval_error]),
                      twoDigitTextYear: (1 if ie[:two_digit_text_year]),
                      numberStoredAsText: (1 if ie[:number_stored_as_text]),
                      formula: (1 if ie[:formula]),
                      formulaRange: (1 if ie[:formula_range]),
                      unlockedFormula: (1 if ie[:unlocked_formula]),
                      emptyCellReference: (1 if ie[:empty_cell_reference]),
                      listDataValidation: (1 if ie[:list_data_validation]),
                      calculatedColumn: (1 if ie[:calculated_column])
                    }
                  )
                end
              end
            end

            unless sheet_tables.empty?
              xml.tag("tableParts", { count: sheet_tables.size }) do
                sheet_tables.each_with_index do |_tbl, i|
                  rid = hyperlink_count + i + 1
                  xml.empty_tag("tablePart", { "r:id": "rId#{rid}" })
                end
              end
            end

            if has_drawing
              # The drawing rId is after hyperlinks + tables + comments + vml
              drawing_rid = hyperlink_count + sheet_tables.size + (has_comments ? 2 : 0) + 1
              xml.empty_tag("drawing", { "r:id": "rId#{drawing_rid}" })
            end

            xml.empty_tag("legacyDrawing", { "r:id": "rId#{vml_rid}" }) if vml_rid
          end
        end
      end

      # : () -> ::Array[untyped]
      def build_shared_string_table
        rt_sst = {}.compare_by_identity # Xlsxrb::Elements::RichText object -> global index
        str_sst = {} # String -> global index
        @sheets.each_value do |sheet_cells|
          sheet_cells.each_value do |value|
            total = rt_sst.size + str_sst.size
            case value
            when Xlsxrb::Elements::RichText
              rt_sst[value] = total unless rt_sst.key?(value)
            when String
              str_sst[value] = total unless str_sst.key?(value)
            else
              next if value.is_a?(Numeric) || value.is_a?(Date) || value == true || value == false || value.is_a?(Xlsxrb::Elements::Formula)

              str = value.to_s
              str_sst[str] = total unless str_sst.key?(str)
            end
          end
        end
        [rt_sst, str_sst]
      end

      # : (untyped sst) -> ::String
      def generate_shared_strings_xml(sst)
        rt_sst, str_sst = sst
        total = rt_sst.size + str_sst.size
        entries = Array.new(total)
        rt_sst.each { |rt, idx| entries[idx] = { type: :rt, value: rt } }
        str_sst.each { |str, idx| entries[idx] = { type: :str, value: str } }

        XmlBuilder.build(declaration: :plain) do |xml|
          xml.tag("sst", { xmlns: SSML_NS, count: total, uniqueCount: total }) do
            entries.each do |entry|
              xml.tag("si") do
                if entry[:type] == :rt
                  write_rich_text_xml(xml, entry[:value])
                else
                  xml.tag("t") { xml.xstring_text(entry[:value]) }
                end
              end
            end
          end
        end
      end

      # : (untyped tbl) -> ::String
      def generate_table_xml(tbl)
        trc = tbl[:totals_row_count].to_i
        hrc = tbl[:header_row_count]
        raw_style = tbl[:style] || tbl[:style_info]
        raw_style = raw_style.to_h if raw_style.respond_to?(:to_h)
        style = if raw_style.is_a?(Hash)
                  raw_style
                elsif raw_style
                  { name: raw_style.to_s }
                else
                  {}
                end
        style_name = style[:name] || "TableStyleMedium2"
        show_first_column = style.key?(:show_first_column) ? style[:show_first_column] : tbl[:show_first_column]
        show_last_column = style.key?(:show_last_column) ? style[:show_last_column] : tbl[:show_last_column]
        show_row_stripes = if style.key?(:show_row_stripes)
                             style[:show_row_stripes]
                           elsif tbl.key?(:show_row_stripes)
                             tbl[:show_row_stripes]
                           else
                             true
                           end
        show_column_stripes = style.key?(:show_column_stripes) ? style[:show_column_stripes] : tbl[:show_column_stripes]

        XmlBuilder.build(declaration: :plain) do |xml|
          xml.tag(
            "table",
            {
              xmlns: SSML_NS,
              id: tbl[:id] || "",
              name: tbl[:name] || "",
              displayName: tbl[:display_name] || "",
              ref: tbl[:ref] || "",
              comment: tbl[:comment],
              headerRowCount: (hrc if hrc && hrc != 1),
              insertRow: (1 if tbl[:insert_row]),
              insertRowShift: (1 if tbl[:insert_row_shift]),
              totalsRowCount: (trc if trc.positive?),
              totalsRowShown: (0 if trc.zero?),
              published: (1 if tbl[:published]),
              headerRowDxfId: tbl[:header_row_dxf_id],
              dataDxfId: tbl[:data_dxf_id],
              totalsRowDxfId: tbl[:totals_row_dxf_id],
              headerRowBorderDxfId: tbl[:header_row_border_dxf_id],
              tableBorderDxfId: tbl[:table_border_dxf_id],
              totalsRowBorderDxfId: tbl[:totals_row_border_dxf_id],
              headerRowCellStyle: tbl[:header_row_cell_style],
              totalsRowCellStyle: tbl[:totals_row_cell_style],
              connectionId: tbl[:connection_id],
              tableType: tbl[:table_type]
            }
          ) do
            xml.empty_tag("autoFilter", { ref: tbl[:ref] || "" }) unless tbl[:auto_filter] == false
            xml.tag("tableColumns", { count: tbl[:columns].size }) do
              tbl[:columns].each_with_index do |col, i|
                col = col.to_h if col.is_a?(Elements::MetadataRecord)
                col_name = if col.is_a?(Hash)
                             col[:name]
                           elsif col.is_a?(Elements::TableColumn)
                             col.name
                           else
                             col
                           end
                if col.is_a?(Hash) || col.is_a?(Elements::TableColumn)
                  calc_formula = col[:calculated_column_formula] || col[:data_formula] || col[:formula]
                  tot_formula = col[:totals_row_formula]
                  trf = col[:totals_row_function] || ("custom" if tot_formula)
                  col_attrs = {
                    id: i + 1,
                    name: col_name || "",
                    totalsRowFunction: trf,
                    totalsRowLabel: col[:totals_row_label],
                    dataDxfId: col[:data_dxf_id],
                    totalsRowDxfId: col[:totals_row_dxf_id],
                    headerRowDxfId: col[:header_row_dxf_id],
                    dataCellStyle: col[:data_cell_style]
                  }
                  if calc_formula || tot_formula
                    xml.tag("tableColumn", col_attrs) do
                      xml.text_tag("calculatedColumnFormula", calc_formula.to_s) if calc_formula
                      xml.text_tag("totalsRowFormula", tot_formula.to_s) if tot_formula
                    end
                  else
                    xml.empty_tag("tableColumn", col_attrs)
                  end
                else
                  xml.empty_tag("tableColumn", { id: i + 1, name: col_name || "" })
                end
              end
            end
            xml.empty_tag(
              "tableStyleInfo",
              {
                name: style_name,
                showFirstColumn: show_first_column ? "1" : "0",
                showLastColumn: show_last_column ? "1" : "0",
                showRowStripes: show_row_stripes ? "1" : "0",
                showColumnStripes: show_column_stripes ? "1" : "0"
              }
            )
          end
        end
      end

      # : (untyped sheet_hyperlinks, ?untyped sheet_tables, ?::Integer table_start_index) -> ::String
      def generate_worksheet_rels(sheet_hyperlinks, sheet_tables = [], table_start_index = 0)
        XmlBuilder.build(declaration: :plain) do |xml|
          xml.tag("Relationships", { xmlns: REL_NS }) do
            rid = 0
            sheet_hyperlinks.each do |(_cell_ref, link)|
              next unless link[:url]

              rid += 1
              xml.empty_tag("Relationship", { Id: "rId#{rid}", Type: "#{DOC_REL_NS}/hyperlink", Target: link[:url], TargetMode: "External" })
            end
            sheet_tables.each_with_index do |_tbl, i|
              rid += 1
              xml.empty_tag("Relationship", { Id: "rId#{rid}", Type: "#{DOC_REL_NS}/table", Target: "../tables/table#{table_start_index + i + 1}.xml" })
            end
          end
        end
      end

      # : (untyped sheet_name, untyped sheet_tables, untyped table_start_index, untyped drawing_idx, untyped comment_idx, untyped pivot_start, untyped pivot_count, ?untyped? threaded_comment_idx) -> untyped
      def build_sheet_rels_parts_v2(sheet_name, sheet_tables, table_start_index, drawing_idx, comment_idx, pivot_start, pivot_count, threaded_comment_idx = nil)
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
        rels << { type: "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment", target: "../threadedComments/threadedComment#{threaded_comment_idx}.xml" } if threaded_comment_idx
        rels << { type: "#{DOC_REL_NS}/drawing", target: "../drawings/drawing#{drawing_idx}.xml" } if drawing_idx
        pivot_count.times do |i|
          pt_idx = pivot_start + i + 1
          rels << { type: "#{DOC_REL_NS}/pivotTable", target: "../pivotTables/pivotTable#{pt_idx}.xml" }
        end
        rels
      end

      # : (untyped rels_data) -> ::String
      def generate_generic_rels(rels_data)
        XmlBuilder.build(declaration: :plain) do |xml|
          xml.tag("Relationships", { xmlns: REL_NS }) do
            rels_data.each_with_index do |rel, i|
              xml.empty_tag(
                "Relationship",
                {
                  Id: "rId#{i + 1}",
                  Type: rel[:type] || "",
                  Target: rel[:target] || "",
                  TargetMode: ("External" if rel[:external])
                }
              )
            end
          end
        end
      end

      # : (XmlBuilder xml) -> void
      def write_workbook_file_version_xml(xml)
        return if @file_version.empty?

        attrs = {
          appName: @file_version[:app_name],
          lastEdited: @file_version[:last_edited],
          lowestEdited: @file_version[:lowest_edited],
          rupBuild: @file_version[:rup_build],
          codeName: @file_version[:code_name]
        }.compact
        xml.empty_tag("fileVersion", attrs) unless attrs.empty?
      end

      # : (XmlBuilder xml) -> void
      def write_workbook_file_sharing_xml(xml)
        return if @file_sharing.empty?

        attrs = {
          readOnlyRecommended: (1 if @file_sharing[:read_only_recommended]),
          userName: @file_sharing[:user_name],
          algorithmName: @file_sharing[:algorithm_name],
          hashValue: @file_sharing[:hash_value],
          saltValue: @file_sharing[:salt_value],
          spinCount: @file_sharing[:spin_count]
        }.compact
        xml.empty_tag("fileSharing", attrs) unless attrs.empty?
      end

      # : (XmlBuilder xml) -> void
      def write_workbook_pr_xml(xml)
        return if @workbook_properties.empty?

        wp = @workbook_properties
        attrs = {
          date1904: xml_bool_attr(wp[:date1904]),
          dateCompatibility: (0 if wp[:date_compatibility] == false),
          defaultThemeVersion: wp[:default_theme_version],
          codeName: wp[:code_name],
          filterPrivacy: (1 if wp[:filter_privacy]),
          autoCompressPictures: (0 if wp[:auto_compress_pictures] == false),
          backupFile: (1 if wp[:backup_file]),
          showObjects: wp[:show_objects],
          updateLinks: wp[:update_links],
          refreshAllConnections: (1 if wp[:refresh_all_connections]),
          checkCompatibility: (1 if wp[:check_compatibility]),
          hidePivotFieldList: (1 if wp[:hide_pivot_field_list]),
          showBorderUnselectedTables: (0 if wp[:show_border_unselected_tables] == false),
          promptedSolutions: (1 if wp[:prompted_solutions]),
          showInkAnnotation: (0 if wp[:show_ink_annotation] == false),
          saveExternalLinkValues: (0 if wp[:save_external_link_values] == false),
          showPivotChartFilter: (1 if wp[:show_pivot_chart_filter]),
          allowRefreshQuery: (1 if wp[:allow_refresh_query]),
          publishItems: (1 if wp[:publish_items])
        }.compact
        xml.empty_tag("workbookPr", attrs) unless attrs.empty?
      end

      # : (XmlBuilder xml) -> void
      def write_workbook_views_xml(xml)
        return if @workbook_views.empty?

        wv = @workbook_views
        attrs = {
          visibility: wv[:visibility],
          minimized: (1 if wv[:minimized]),
          showHorizontalScroll: xml_bool_attr(wv[:show_horizontal_scroll]),
          showVerticalScroll: xml_bool_attr(wv[:show_vertical_scroll]),
          showSheetTabs: xml_bool_attr(wv[:show_sheet_tabs]),
          xWindow: wv[:x_window],
          yWindow: wv[:y_window],
          windowWidth: wv[:window_width],
          windowHeight: wv[:window_height],
          tabRatio: wv[:tab_ratio],
          firstSheet: wv[:first_sheet],
          activeTab: wv[:active_sheet_index] || wv[:active_tab],
          autoFilterDateGrouping: xml_bool_attr(wv[:auto_filter_date_grouping])
        }.compact
        xml.tag("bookViews") do
          xml.empty_tag("workbookView", attrs) unless attrs.empty?
        end
      end

      # : (XmlBuilder xml) -> void
      def write_workbook_protection_xml(xml)
        wp = @workbook_protection
        attrs = {
          lockStructure: (1 if wp[:lock_structure]),
          lockWindows: (1 if wp[:lock_windows]),
          lockRevision: (1 if wp[:lock_revision])
        }
        if wp[:algorithm_name]
          attrs[:workbookAlgorithmName] = wp[:algorithm_name]
          attrs[:workbookHashValue] = wp[:hash_value]
          attrs[:workbookSaltValue] = wp[:salt_value]
          attrs[:workbookSpinCount] = wp[:spin_count]
        elsif wp[:password]
          attrs[:workbookPassword] = wp[:password]
        end
        if wp[:revisions_algorithm_name]
          attrs[:revisionsAlgorithmName] = wp[:revisions_algorithm_name]
          attrs[:revisionsHashValue] = wp[:revisions_hash_value]
          attrs[:revisionsSaltValue] = wp[:revisions_salt_value]
          attrs[:revisionsSpinCount] = wp[:revisions_spin_count]
        elsif wp[:revisions_password]
          attrs[:revisionsPassword] = wp[:revisions_password]
        end
        attrs.compact!
        xml.empty_tag("workbookProtection", attrs) unless attrs.empty?
      end

      # : (XmlBuilder xml) -> void
      def write_workbook_calc_pr_xml(xml)
        return if @calc_properties.empty?

        cp = @calc_properties
        attrs = {
          calcId: cp[:calc_id],
          calcMode: cp[:calc_mode],
          fullCalcOnLoad: xml_bool_attr(cp[:full_calc_on_load]),
          iterate: xml_bool_attr(cp[:iterate]),
          iterateCount: cp[:iterate_count],
          iterateDelta: cp[:iterate_delta],
          refMode: cp[:ref_mode],
          calcCompleted: xml_bool_attr(cp[:calc_completed]),
          calcOnSave: xml_bool_attr(cp[:calc_on_save]),
          fullPrecision: xml_bool_attr(cp[:full_precision]),
          concurrentCalc: xml_bool_attr(cp[:concurrent_calc]),
          concurrentManualCount: cp[:concurrent_manual_count],
          forceFullCalc: xml_bool_attr(cp[:force_full_calc])
        }.compact
        xml.empty_tag("calcPr", attrs) unless attrs.empty?
      end

      # : (XmlBuilder xml, untyped sheet_props) -> void
      def write_sheet_pr_xml(xml, sheet_props)
        sp_attrs = {
          syncHorizontal: (1 if sheet_props[:sync_horizontal]),
          syncVertical: (1 if sheet_props[:sync_vertical]),
          syncRef: sheet_props[:sync_ref],
          transitionEvaluation: (1 if sheet_props[:transition_evaluation]),
          transitionEntry: (1 if sheet_props[:transition_entry]),
          codeName: sheet_props[:code_name],
          filterMode: (1 if sheet_props[:filter_mode]),
          published: (0 if sheet_props[:published] == false),
          enableFormatConditionsCalculation: (0 if sheet_props[:enable_format_conditions_calculation] == false)
        }.compact
        tab_color_attrs = resolve_tab_color_attrs(sheet_props)

        sb = sheet_props[:summary_below]
        sr = sheet_props[:summary_right]
        as = sheet_props[:apply_styles]
        sos = sheet_props[:show_outline_symbols]
        has_outline = !(sb.nil? && sr.nil? && as.nil? && sos.nil?)

        psp = sheet_props[:page_setup_properties] || sheet_props[:page_setup_pr]
        ftp = sheet_props[:fit_to_page]
        ftp = psp[:fit_to_page] if ftp.nil? && psp.is_a?(Hash)
        apb = sheet_props[:auto_page_breaks]
        apb = psp[:auto_page_breaks] if apb.nil? && psp.is_a?(Hash)
        has_psp = !(ftp.nil? && apb.nil?)

        return if sp_attrs.empty? && tab_color_attrs.nil? && !has_outline && !has_psp

        xml.tag("sheetPr", sp_attrs) do
          xml.empty_tag("tabColor", tab_color_attrs) if tab_color_attrs
          if has_outline
            xml.empty_tag(
              "outlinePr",
              {
                applyStyles: xml_bool_attr(as),
                summaryBelow: xml_bool_attr(sb),
                summaryRight: xml_bool_attr(sr),
                showOutlineSymbols: xml_bool_attr(sos)
              }
            )
          end
          if has_psp
            xml.empty_tag(
              "pageSetUpPr",
              {
                fitToPage: xml_bool_attr(ftp),
                autoPageBreaks: xml_bool_attr(apb)
              }
            )
          end
        end
      end

      # : (untyped sheet_props) -> (nil | ::Hash[untyped, untyped])
      def resolve_tab_color_attrs(sheet_props)
        tc = sheet_props[:tab_color]
        if tc.is_a?(Hash)
          attrs = {
            auto: (1 if tc[:auto]),
            indexed: tc[:indexed],
            rgb: tc[:rgb]&.then { |v| v.to_s.delete_prefix("#") },
            theme: tc[:theme],
            tint: tc[:tint]
          }.compact
          attrs.empty? ? nil : attrs
        elsif tc
          { rgb: tc.to_s.delete_prefix("#") }
        elsif sheet_props[:tab_color_theme]
          { theme: sheet_props[:tab_color_theme], tint: sheet_props[:tab_color_tint] }
        elsif sheet_props[:tab_color_indexed]
          { indexed: sheet_props[:tab_color_indexed] }
        elsif sheet_props[:tab_color_auto]
          { auto: 1 }
        end
      end

      # : (XmlBuilder xml, untyped sheet_sv, untyped sheet_fp, untyped sheet_sel) -> void
      def write_sheet_views_xml(xml, sheet_sv, sheet_fp, sheet_sel)
        xml.tag("sheetViews") do
          xml.tag(
            "sheetView",
            {
              windowProtection: xml_bool_attr(sheet_sv[:window_protection]),
              showFormulas: xml_bool_attr(sheet_sv[:show_formulas]),
              showGridLines: xml_bool_attr(sheet_sv[:show_grid_lines]),
              showRowColHeaders: xml_bool_attr(sheet_sv[:show_row_col_headers]),
              showZeros: xml_bool_attr(sheet_sv[:show_zeros]),
              rightToLeft: xml_bool_attr(sheet_sv[:right_to_left]),
              tabSelected: (1 if sheet_sv[:tab_selected]),
              showRuler: xml_bool_attr(sheet_sv[:show_ruler]),
              showOutlineSymbols: xml_bool_attr(sheet_sv[:show_outline_symbols]),
              defaultGridColor: xml_bool_attr(sheet_sv[:default_grid_color]),
              showWhiteSpace: xml_bool_attr(sheet_sv[:show_white_space]),
              view: sheet_sv[:view],
              topLeftCell: sheet_sv[:top_left_cell],
              colorId: sheet_sv[:color_id],
              zoomScale: sheet_sv[:zoom_scale],
              zoomScaleNormal: sheet_sv[:zoom_scale_normal],
              zoomScaleSheetLayoutView: sheet_sv[:zoom_scale_sheet_layout_view],
              zoomScalePageLayoutView: sheet_sv[:zoom_scale_page_layout_view],
              workbookViewId: 0
            }
          ) do
            if sheet_fp && sheet_fp[:state] == :split
              has_x = sheet_fp[:x_split].to_i.positive?
              has_y = sheet_fp[:y_split].to_i.positive?
              active_pane = if has_y && has_x
                              "bottomRight"
                            elsif has_y
                              "bottomLeft"
                            else
                              "topRight"
                            end
              xml.empty_tag(
                "pane",
                {
                  xSplit: (sheet_fp[:x_split] if has_x),
                  ySplit: (sheet_fp[:y_split] if has_y),
                  topLeftCell: sheet_fp[:top_left_cell],
                  activePane: active_pane
                }
              )
            elsif sheet_fp
              fp_col = sheet_fp[:column] || sheet_fp[:col] || 0
              if sheet_fp[:row].to_i.positive? || fp_col.to_i.positive?
                top_left = "#{index_to_column_letter(fp_col.to_i + 1)}#{sheet_fp[:row].to_i + 1}"
                active_pane = if sheet_fp[:row].to_i.positive? && fp_col.to_i.positive?
                                "bottomRight"
                              elsif sheet_fp[:row].to_i.positive?
                                "bottomLeft"
                              else
                                "topRight"
                              end
                xml.empty_tag(
                  "pane",
                  {
                    ySplit: (sheet_fp[:row] if sheet_fp[:row].to_i.positive?),
                    xSplit: (fp_col if fp_col.to_i.positive?),
                    topLeftCell: top_left,
                    activePane: active_pane,
                    state: "frozen"
                  }
                )
              end
            end

            if sheet_sel
              xml.empty_tag(
                "selection",
                {
                  pane: sheet_sel[:pane],
                  activeCell: sheet_sel[:active_cell],
                  activeCellId: sheet_sel[:active_cell_id],
                  sqref: sheet_sel[:sqref]
                }
              )
            end
          end
        end
      end

      # : (XmlBuilder xml, untyped sheet_cells, untyped sheet_row_attrs, untyped sheet_cell_styles, untyped? sst, untyped cell_phonetic) -> void
      def write_sheet_data_xml(xml, sheet_cells, sheet_row_attrs, sheet_cell_styles, sst, cell_phonetic)
        xml.tag("sheetData") do
          cells_by_row = {}
          sheet_cells.each do |address, value|
            row_num = extract_row_number(address)
            col_letter = extract_column_letter(address)
            cells_by_row[row_num] ||= {}
            cells_by_row[row_num][col_letter] = value
          end

          sheet_row_attrs.each_key { |rn| cells_by_row[rn] ||= {} }

          cells_by_row.sort.each do |row_num, row_cells|
            ra = sheet_row_attrs[row_num]
            row_attrs = { r: row_num }
            if ra
              if ra[:height]
                row_attrs[:ht] = ra[:height]
                row_attrs[:customHeight] = 1
              end
              row_attrs[:hidden] = 1 if ra[:hidden]
              row_attrs[:outlineLevel] = ra[:outline_level] if ra[:outline_level]
              row_attrs[:collapsed] = 1 if ra[:collapsed]
              if ra.key?(:style)
                row_attrs[:s] = ra[:style] || ""
                row_attrs[:customFormat] = 1
              end
              row_attrs[:thickTop] = 1 if ra[:thick_top]
              row_attrs[:thickBot] = 1 if ra[:thick_bot]
              row_attrs[:ph] = 1 if ra[:ph]
            end
            xml.tag("row", row_attrs) do
              row_cells.sort_by { |col, _| column_letter_to_index(col) }.each do |col_letter, value|
                cell_ref = "#{col_letter}#{row_num}"
                style_idx = resolve_style_index(sheet_cell_styles[cell_ref])
                write_cell_xml(xml, cell_ref, value, style_idx, sst, ph: cell_phonetic[cell_ref])
              end
            end
          end
        end
      end

      # : (XmlBuilder xml, untyped sheet_prot) -> void
      def write_sheet_protection_xml(xml, sheet_prot)
        attrs = {
          sheet: (1 if sheet_prot[:sheet] != false),
          objects: (1 if sheet_prot[:objects]),
          scenarios: (1 if sheet_prot[:scenarios]),
          formatCells: (0 if sheet_prot[:format_cells] == false),
          formatColumns: (0 if sheet_prot[:format_columns] == false),
          formatRows: (0 if sheet_prot[:format_rows] == false),
          insertColumns: (0 if sheet_prot[:insert_columns] == false),
          insertRows: (0 if sheet_prot[:insert_rows] == false),
          insertHyperlinks: (0 if sheet_prot[:insert_hyperlinks] == false),
          deleteColumns: (0 if sheet_prot[:delete_columns] == false),
          deleteRows: (0 if sheet_prot[:delete_rows] == false),
          selectLockedCells: (1 if sheet_prot[:select_locked_cells]),
          sort: (0 if sheet_prot[:sort] == false),
          autoFilter: (0 if sheet_prot[:auto_filter] == false),
          pivotTables: (0 if sheet_prot[:pivot_tables] == false),
          selectUnlockedCells: (1 if sheet_prot[:select_unlocked_cells])
        }
        if sheet_prot[:algorithm_name]
          attrs[:algorithmName] = sheet_prot[:algorithm_name]
          attrs[:hashValue] = sheet_prot[:hash_value]
          attrs[:saltValue] = sheet_prot[:salt_value]
          attrs[:spinCount] = sheet_prot[:spin_count]
        elsif sheet_prot[:password]
          attrs[:password] = sheet_prot[:password]
        end
        xml.empty_tag("sheetProtection", attrs)
      end

      # : (XmlBuilder xml, untyped prot_ranges) -> void
      def write_protected_ranges_xml(xml, prot_ranges)
        xml.tag("protectedRanges") do
          prot_ranges.each do |pr|
            attrs = { sqref: pr[:sqref] || "", name: pr[:name] || "" }
            if pr[:algorithm_name]
              attrs[:algorithmName] = pr[:algorithm_name]
              attrs[:hashValue] = pr[:hash_value]
              attrs[:saltValue] = pr[:salt_value]
              attrs[:spinCount] = pr[:spin_count]
            end
            sds = pr[:security_descriptors] || []
            if sds.empty?
              xml.empty_tag("protectedRange", attrs)
            else
              xml.tag("protectedRange", attrs) do
                sds.each { |sd| xml.text_tag("securityDescriptor", sd || "") }
              end
            end
          end
        end
      end

      # : (XmlBuilder xml, untyped sheet_scenarios) -> void
      def write_scenarios_xml(xml, sheet_scenarios)
        xml.tag(
          "scenarios",
          {
            current: sheet_scenarios[:current],
            show: sheet_scenarios[:show],
            sqref: sheet_scenarios[:sqref]
          }
        ) do
          (sheet_scenarios[:scenarios] || []).each do |sc|
            xml.tag(
              "scenario",
              {
                name: sc[:name] || "",
                locked: (1 if sc[:locked]),
                hidden: (1 if sc[:hidden]),
                count: (sc[:input_cells].size if sc[:input_cells]&.any?),
                user: sc[:user],
                comment: sc[:comment]
              }
            ) do
              (sc[:input_cells] || []).each do |ic|
                xml.empty_tag(
                  "inputCells",
                  {
                    r: ic[:r] || "",
                    val: ic[:val].to_s,
                    deleted: (1 if ic[:deleted]),
                    undone: (1 if ic[:undone]),
                    numFmtId: ic[:num_fmt_id]
                  }
                )
              end
            end
          end
        end
      end

      # : (XmlBuilder xml, untyped data_consol) -> void
      def write_data_consolidate_xml(xml, data_consol)
        attrs = {
          function: (data_consol[:function] if data_consol[:function] && data_consol[:function] != "sum"),
          startLabels: (1 if data_consol[:start_labels]),
          leftLabels: (1 if data_consol[:left_labels]),
          topLabels: (1 if data_consol[:top_labels]),
          link: (1 if data_consol[:link])
        }
        refs = data_consol[:data_refs] || []
        if refs.empty?
          xml.empty_tag("dataConsolidate", attrs)
        else
          xml.tag("dataConsolidate", attrs) do
            xml.tag("dataRefs", { count: refs.size }) do
              refs.each do |r|
                xml.empty_tag("dataRef", { ref: r[:ref], name: r[:name], sheet: r[:sheet] })
              end
            end
          end
        end
      end

      # : (XmlBuilder xml, untyped sheet_dv, ::Hash[untyped, untyped] dv_options) -> void
      def write_data_validations_xml(xml, sheet_dv, dv_options)
        valid_dv = sheet_dv.reject { |dv| dv[:sqref].nil? || dv[:sqref].to_s.empty? }
        return if valid_dv.empty?

        xml.tag(
          "dataValidations",
          {
            count: valid_dv.size,
            disablePrompts: (1 if dv_options[:disable_prompts]),
            xWindow: dv_options[:x_window],
            yWindow: dv_options[:y_window]
          }
        ) do
          valid_dv.each do |dv|
            attrs = {
              sqref: dv[:sqref] || "",
              type: dv[:type],
              operator: dv[:operator],
              errorStyle: dv[:error_style],
              allowBlank: (1 if dv[:allow_blank]),
              showInputMessage: (1 if dv[:show_input_message]),
              showErrorMessage: (1 if dv[:show_error_message]),
              errorTitle: dv[:error_title],
              error: dv[:error],
              promptTitle: dv[:prompt_title],
              prompt: dv[:prompt],
              showDropDown: (1 if dv[:show_drop_down]),
              imeMode: dv[:ime_mode]
            }
            if dv[:formula1] || dv[:formula2]
              xml.tag("dataValidation", attrs) do
                xml.text_tag("formula1", dv[:formula1]) if dv[:formula1]
                xml.text_tag("formula2", dv[:formula2]) if dv[:formula2]
              end
            else
              xml.empty_tag("dataValidation", attrs)
            end
          end
        end
      end

      # : (XmlBuilder xml, untyped sheet_po, untyped sheet_pm, untyped sheet_ps, untyped sheet_hf) -> void
      def write_print_and_page_xml(xml, sheet_po, sheet_pm, sheet_ps, sheet_hf)
        unless sheet_po.empty?
          po_attrs = {
            gridLines: (1 if sheet_po[:grid_lines]),
            headings: (1 if sheet_po[:headings]),
            horizontalCentered: (1 if sheet_po[:horizontal_centered]),
            verticalCentered: (1 if sheet_po[:vertical_centered]),
            gridLinesSet: (0 if sheet_po[:grid_lines_set] == false)
          }.compact
          xml.empty_tag("printOptions", po_attrs) unless po_attrs.empty?
        end

        if sheet_pm
          pm_attrs = %w[left right top bottom header footer].each_with_object({}) do |k, h|
            v = sheet_pm[k.to_sym]
            h[k] = v if v
          end
          xml.empty_tag("pageMargins", pm_attrs) unless pm_attrs.empty?
        end

        unless sheet_ps.empty?
          ps_attrs = {
            orientation: sheet_ps[:orientation],
            paperSize: sheet_ps[:paper_size],
            scale: sheet_ps[:scale],
            fitToWidth: sheet_ps[:fit_to_width],
            fitToHeight: sheet_ps[:fit_to_height],
            pageOrder: sheet_ps[:page_order],
            blackAndWhite: (1 if sheet_ps[:black_and_white]),
            draft: (1 if sheet_ps[:draft]),
            cellComments: sheet_ps[:cell_comments],
            firstPageNumber: (sheet_ps[:first_page_number] if sheet_ps[:first_page_number] && sheet_ps[:first_page_number].to_i >= 0),
            useFirstPageNumber: (1 if sheet_ps[:use_first_page_number]),
            horizontalDpi: (sheet_ps[:horizontal_dpi] if sheet_ps[:horizontal_dpi].to_i.positive?),
            verticalDpi: (sheet_ps[:vertical_dpi] if sheet_ps[:vertical_dpi].to_i.positive?),
            copies: sheet_ps[:copies],
            paperHeight: sheet_ps[:paper_height],
            paperWidth: sheet_ps[:paper_width],
            errors: sheet_ps[:errors],
            usePrinterDefaults: (0 if sheet_ps[:use_printer_defaults] == false)
          }.compact
          xml.empty_tag("pageSetup", ps_attrs) unless ps_attrs.empty?
        end

        return if sheet_hf.empty?

        xml.tag(
          "headerFooter",
          {
            differentFirst: (1 if sheet_hf[:different_first]),
            differentOddEven: (1 if sheet_hf[:different_odd_even]),
            scaleWithDoc: (0 if sheet_hf[:scale_with_doc] == false),
            alignWithMargins: (0 if sheet_hf[:align_with_margins] == false)
          }
        ) do
          xml.text_tag("oddHeader", sheet_hf[:odd_header]) if sheet_hf[:odd_header]
          xml.text_tag("oddFooter", sheet_hf[:odd_footer]) if sheet_hf[:odd_footer]
          xml.text_tag("evenHeader", sheet_hf[:even_header]) if sheet_hf[:even_header]
          xml.text_tag("evenFooter", sheet_hf[:even_footer]) if sheet_hf[:even_footer]
          xml.text_tag("firstHeader", sheet_hf[:first_header]) if sheet_hf[:first_header]
          xml.text_tag("firstFooter", sheet_hf[:first_footer]) if sheet_hf[:first_footer]
        end
      end

      # : (XmlBuilder xml, ::String tag, untyped breaks, default_max: ::Integer) -> void
      def write_page_breaks_xml(xml, tag, breaks, default_max:)
        manual_count = breaks.count do |brk|
          if brk.is_a?(Elements::PageBreak)
            brk.manual?
          elsif brk.is_a?(Hash)
            brk[:man] != false
          else
            true
          end
        end
        xml.tag(tag, { count: breaks.size, manualBreakCount: manual_count }) do
          breaks.each { |brk| write_brk_xml(xml, brk, default_max: default_max) }
        end
      end

      # Extracts the row number from a cell address, e.g. 1 from "A1".
      # : (untyped cell_address) -> untyped
      def extract_row_number(cell_address)
        cell_address.match(/(\d+)$/)[1].to_i
      end

      # Converts a range like "A1:D20" to absolute "$A$1:$D$20".
      # : (untyped range) -> untyped
      def absolute_range(range)
        DslHelpers.absolute_range(range)
      end

      # Converts "A1" to "$A$1".
      # : (untyped cell_ref) -> ::String
      def absolute_cell(cell_ref)
        m = cell_ref.match(/\A([A-Z]+)(\d+)\z/)
        raise ArgumentError, "invalid cell reference: #{cell_ref}" unless m

        "$#{m[1]}$#{m[2]}"
      end
    end
  end
end
