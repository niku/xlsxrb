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

    # Returns ordered sheet names.
    attr_reader :sheet_order

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

      entries = {
        "[Content_Types].xml" => generate_content_types_xml,
        "_rels/.rels" => generate_rels_root,
        "xl/workbook.xml" => generate_workbook_xml,
        "xl/_rels/workbook.xml.rels" => generate_workbook_rels,
        "xl/styles.xml" => generate_styles_xml
      }

      entries["docProps/core.xml"] = generate_core_properties_xml unless @core_properties.empty?
      entries["docProps/app.xml"] = generate_app_properties_xml unless @app_properties.empty?

      @sheet_order.each_with_index do |sheet_name, i|
        entries["xl/worksheets/sheet#{i + 1}.xml"] = generate_worksheet_xml(
          @sheets[sheet_name], @column_widths[sheet_name], @column_attrs[sheet_name], @row_attrs[sheet_name],
          @auto_filters[sheet_name], @filter_columns[sheet_name], @sort_state[sheet_name],
          @merge_cells[sheet_name], @hyperlinks[sheet_name],
          @cell_styles[sheet_name], @sheet_properties[sheet_name], @sheet_formats[sheet_name],
          @sheet_views[sheet_name], @freeze_panes[sheet_name], @selections[sheet_name],
          @print_options[sheet_name], @page_margins[sheet_name], @page_setup[sheet_name],
          @header_footer[sheet_name], @row_breaks[sheet_name], @col_breaks[sheet_name],
          @data_validations[sheet_name]
        )
        next if @hyperlinks[sheet_name].empty?

        entries["xl/worksheets/_rels/sheet#{i + 1}.xml.rels"] = generate_worksheet_rels(@hyperlinks[sheet_name])
      end

      generator = ZipGenerator.new(filepath)
      entries.each { |path, content| generator.add_entry(path, content) }
      generator.generate
    end

    private

    def generate_content_types_xml
      parts = [
        XML_HEADER,
        %(<Types xmlns="#{CT_NS}">),
        %(<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>),
        %(<Default Extension="xml" ContentType="application/xml"/>),
        %(<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>),
        %(<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>)
      ]
      @sheet_order.each_with_index do |_, i|
        parts << %(<Override PartName="/xl/worksheets/sheet#{i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>)
      end
      parts << %(<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>) unless @core_properties.empty?
      parts << %(<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>) unless @app_properties.empty?
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

    def generate_workbook_xml
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
        attrs << %(fullCalcOnLoad="#{@calc_properties[:full_calc_on_load] ? 1 : 0}") unless @calc_properties[:full_calc_on_load].nil?
        parts << "<calcPr #{attrs.join(" ")}/>" unless attrs.empty?
      end

      parts << "</workbook>"
      parts.join
    end

    def generate_workbook_rels
      parts = [
        XML_HEADER,
        %(<Relationships xmlns="#{REL_NS}">)
      ]
      @sheet_order.each_with_index do |_, i|
        parts << %(<Relationship Id="rId#{i + 1}" Type="#{DOC_REL_NS}/worksheet" Target="worksheets/sheet#{i + 1}.xml"/>)
      end
      styles_rid = @sheet_order.size + 1
      parts << %(<Relationship Id="rId#{styles_rid}" Type="#{DOC_REL_NS}/styles" Target="styles.xml"/>)
      parts << "</Relationships>"
      parts.join
    end

    def generate_worksheet_xml(sheet_cells, sheet_col_widths, sheet_col_attrs, sheet_row_attrs, sheet_auto_filter, sheet_filter_cols, sheet_sort, sheet_merge_cells, sheet_hyperlinks, sheet_cell_styles, sheet_props, sheet_fmt, sheet_sv, sheet_fp, sheet_sel, sheet_po, sheet_pm, sheet_ps, sheet_hf, sheet_rb, sheet_cb, sheet_dv)
      worksheet_attrs = %(xmlns="#{SSML_NS}")
      worksheet_attrs << %( xmlns:r="#{DOC_REL_NS}") unless sheet_hyperlinks.empty?
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
          parts << cell_xml(cell_ref, value, style_idx)
        end
        parts << "</row>"
      end

      parts << "</sheetData>"

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

      parts << "</worksheet>"
      parts.join
    end

    def generate_worksheet_rels(sheet_hyperlinks)
      parts = [
        XML_HEADER,
        %(<Relationships xmlns="#{REL_NS}">)
      ]
      sheet_hyperlinks.each_with_index do |(_cell_ref, url), idx|
        parts << %(<Relationship Id="rId#{idx + 1}" Type="#{DOC_REL_NS}/hyperlink" Target="#{xml_escape(url)}" TargetMode="External"/>)
      end
      parts << "</Relationships>"
      parts.join
    end

    def xml_escape(value)
      value.to_s
           .gsub("&", "&amp;")
           .gsub("<", "&lt;")
           .gsub(">", "&gt;")
           .gsub('"', "&quot;")
           .gsub("'", "&apos;")
    end

    def cell_xml(cell_ref, value, style_idx)
      s_attr = style_idx ? %( s="#{style_idx}") : ""
      case value
      when Formula
        parts = %(<c r="#{cell_ref}"#{s_attr}><f>#{xml_escape(value.expression)}</f>)
        parts << "<v>#{xml_escape(value.cached_value.to_s)}</v>" unless value.cached_value.nil?
        parts << "</c>"
        parts
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
        %(<c r="#{cell_ref}" t="inlineStr"#{s_attr}><is><t>#{xml_escape(value)}</t></is></c>)
      end
    end

    # Returns the numFmtId for dates, registering it on first use.
    def date_num_fmt_id
      @date_num_fmt_id ||= add_number_format(DEFAULT_DATE_FORMAT)
    end

    # Maps a numFmtId to a cellXfs index. Index 0 is the default (no format).
    def resolve_style_index(num_fmt_id)
      return nil if num_fmt_id.nil?

      # Build the xf index mapping on first call.
      @xf_index_map ||= begin
        map = {}
        @num_fmts.each_with_index do |nf, i|
          map[nf[:num_fmt_id]] = i + 1 # 0 is the default xf
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

      # fonts — one default
      parts << %(<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>)

      # fills — two required
      parts << %(<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>)

      # borders — one default
      parts << %(<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>)

      # cellStyleXfs
      parts << %(<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>)

      # cellXfs — default + one per numFmt
      xf_count = 1 + @num_fmts.size
      parts << %(<cellXfs count="#{xf_count}">)
      parts << %(<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>)
      @num_fmts.each do |nf|
        parts << %(<xf numFmtId="#{nf[:num_fmt_id]}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>)
      end
      parts << "</cellXfs>"

      parts << "</styleSheet>"
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
