# frozen_string_literal: true

require "date"
require "openssl"
require "securerandom"
require "tempfile"
require "opentelemetry"
require_relative "xlsxrb/version"
require_relative "xlsxrb/ooxml/zip_generator"
require_relative "xlsxrb/ooxml/writer"
require_relative "xlsxrb/ooxml/reader"
require_relative "xlsxrb/ooxml"
require_relative "xlsxrb/elements"
require_relative "xlsxrb/style_builder"

# Ruby XLSX read/write library.
module Xlsxrb
  class Error < StandardError; end

  TRACER = OpenTelemetry.tracer_provider.tracer("xlsxrb", Xlsxrb::VERSION)

  # Builder for block-style chart definitions.
  class ChartBuilder
    def initialize
      @options = {}
    end

    attr_reader :options

    def type(value) = @options[:type] = value
    def title(value) = @options[:title] = value

    def series(value = nil, &block)
      @options[:series] ||= []
      if block_given?
        sb = SeriesBuilder.new
        block.call(sb)
        @options[:series] << sb.options
      elsif value
        @options[:series] << value
      end
      @options[:series]
    end

    def method_missing(name, *args, **kwargs, &)
      key = name.to_sym
      @options[key] = kwargs.empty? ? args.first : kwargs
    end

    def respond_to_missing?(_name, _include_private = false)
      true
    end

    # Builder for a single series entry in block-style chart definitions.
    class SeriesBuilder
      def initialize
        @options = {}
      end

      attr_reader :options

      def method_missing(name, *args, **kwargs, &)
        key = name.to_sym
        @options[key] = kwargs.empty? ? args.first : kwargs
      end

      def respond_to_missing?(_name, _include_private = false)
        true
      end
    end
  end

  # Generic builder for block-style feature definitions.
  # Supports method_missing for setting arbitrary keys.
  # --- Facade API ---

  # Reads an XLSX file into an Elements::Workbook.
  # source: file path (String) or IO object.
  def self.read(source)
    attributes = source.is_a?(String) ? { "filepath" => source } : {}
    TRACER.in_span("Xlsxrb.read", attributes: attributes) do
      entries = Ooxml::ZipReader.open(source, &:read_all)
      shared_strings = Ooxml::SharedStringsParser.parse(entries["xl/sharedStrings.xml"])
      styles = Ooxml::StylesParser.parse(entries["xl/styles.xml"])
      workbook_sheets = Ooxml::WorkbookParser.parse(entries["xl/workbook.xml"])
      rels = Ooxml::RelationshipsParser.parse(entries["xl/_rels/workbook.xml.rels"])

      sheets = workbook_sheets.map do |sheet_info|
        target = rels[sheet_info[:r_id]]
        next nil unless target

        sheet_path = target.start_with?("/") ? target.delete_prefix("/") : "xl/#{target}"
        sheet_xml = entries[sheet_path]
        build_worksheet(sheet_info[:name], sheet_xml, shared_strings, styles)
      end.compact

      Elements::Workbook.new(sheets: sheets, shared_strings: shared_strings, styles: styles)
    end
  end

  # Writes an Elements::Workbook to an XLSX file.
  # target: file path (String) or IO object.
  def self.write(target, workbook)
    raise Error, "target is required" if target.nil?
    raise Error, "workbook must be an Elements::Workbook" unless workbook.is_a?(Elements::Workbook)

    attributes = target.is_a?(String) ? { "filepath" => target } : {}
    TRACER.in_span("Xlsxrb.write", attributes: attributes) do
      sst = []
      sst_index = {}

      # Collect shared strings and build index
      sheet_data = workbook.sheets.map do |ws|
        rows = ws.rows.map do |row|
          cells = row.cells.map do |cell|
            raw = build_raw_cell(cell, sst, sst_index)
            raw
          end
          { index: row.index, cells: cells, attrs: build_row_attrs(row), unmapped: [] }
        end
        columns = ws.columns.map do |col|
          { index: col.index, width: col.width, hidden: col.hidden, custom_width: col.custom_width }
        end
        sd = { name: ws.name, rows: rows, columns: columns }
        sd[:charts] = ws.charts unless ws.charts.empty?

        # Extract facade metadata from unmapped_data
        facade = ws.unmapped_data[:facade]
        facade&.each { |key, val| sd[key] = val }

        sd
      end

      # Extract workbook-level facade metadata
      wb_facade = workbook.unmapped_data[:facade] || {}
      Ooxml::WorkbookWriter.write(
        target,
        sheets: sheet_data,
        shared_strings: sst,
        styles: workbook.styles,
        defined_names: wb_facade[:defined_names],
        core_properties: wb_facade[:core_properties],
        app_properties: wb_facade[:app_properties],
        custom_properties: wb_facade[:custom_properties],
        workbook_protection: wb_facade[:workbook_protection]
      )
    end
  end

  # Streaming read: yields Elements::Row one at a time.
  # source: file path (String) or IO object.
  # Options:
  #   sheet: sheet index (0-based Integer) or name (String). Defaults to 0.
  def self.foreach(source, sheet: 0, &block)
    return enum_for(:foreach, source, sheet: sheet) unless block

    attributes = source.is_a?(String) ? { "filepath" => source } : {}
    TRACER.in_span("Xlsxrb.foreach", attributes: attributes) do
      entries = Ooxml::ZipReader.open(source, &:read_all)
      shared_strings = Ooxml::SharedStringsParser.parse(entries["xl/sharedStrings.xml"])
      workbook_sheets = Ooxml::WorkbookParser.parse(entries["xl/workbook.xml"])
      rels = Ooxml::RelationshipsParser.parse(entries["xl/_rels/workbook.xml.rels"])

      target_sheet = case sheet
                     when Integer
                       workbook_sheets[sheet]
                     when String
                       workbook_sheets.find { |s| s[:name] == sheet }
                     end
      next unless target_sheet

      target = rels[target_sheet[:r_id]]
      next unless target

      sheet_path = target.start_with?("/") ? target.delete_prefix("/") : "xl/#{target}"
      sheet_xml = entries[sheet_path]
      next if sheet_xml.nil? || sheet_xml.empty?

      Ooxml::WorksheetParser.each_row(sheet_xml, shared_strings: shared_strings) do |raw_row|
        row = build_row_from_raw(raw_row)
        block.call(row)
      end
    end
  end

  # Streaming write: yields a StreamWriter context for building XLSX on-the-fly.
  # target: file path (String) or IO object.
  def self.generate(target, &block)
    raise Error, "target is required" if target.nil?
    raise Error, "block is required" unless block

    attributes = target.is_a?(String) ? { "filepath" => target } : {}
    TRACER.in_span("Xlsxrb.generate", attributes: attributes) do
      stream_writer = StreamWriter.new(target)
      block.call(stream_writer)
      stream_writer.close
    end
  end

  # Builds an Elements::Workbook in memory using a DSL.
  def self.build(&block)
    raise Error, "block is required" unless block

    TRACER.in_span("Xlsxrb.build") do
      builder = WorkbookBuilder.new
      block.call(builder)
      builder.build
    end
  end

  # DSL context for Xlsxrb.build.
  class WorkbookBuilder
    def initialize
      @sheets = []
      @sheet_builders = [] # Keep track of sheet builders for style processing
      @defined_names = []
      @core_properties = {}
      @app_properties = {}
      @custom_properties = []
      @workbook_protection = nil
    end

    # Add a new sheet.
    def add_sheet(name = nil, &block)
      name ||= "Sheet#{@sheets.size + 1}"
      sheet_builder = WorksheetBuilder.new(name)
      block.call(sheet_builder) if block_given?
      @sheet_builders << sheet_builder
      @sheets << sheet_builder.build
    end

    # --- Workbook-Level Methods ---

    # Add a defined name.
    def add_defined_name(name, value, sheet: nil, hidden: false)
      entry = { name: name, value: value, hidden: hidden }
      entry[:local_sheet_name] = sheet if sheet
      @defined_names << entry
    end

    # Set the print area for a sheet.
    def set_print_area(range, sheet: nil)
      sheet_name = sheet || @sheets.last&.name || "Sheet1"
      value = "'#{sheet_name}'!#{absolute_range(range)}"
      @defined_names.reject! { |dn| dn[:name] == "_xlnm.Print_Area" && dn[:local_sheet_name] == sheet_name }
      add_defined_name("_xlnm.Print_Area", value, sheet: sheet_name)
    end

    # Set print titles for a sheet.
    def set_print_titles(rows: nil, cols: nil, sheet: nil)
      sheet_name = sheet || @sheets.last&.name || "Sheet1"
      parts = []
      parts << "'#{sheet_name}'!$#{cols.sub(":", ":$")}" if cols
      parts << "'#{sheet_name}'!$#{rows.sub(":", ":$")}" if rows
      value = parts.join(",")
      @defined_names.reject! { |dn| dn[:name] == "_xlnm.Print_Titles" && dn[:local_sheet_name] == sheet_name }
      add_defined_name("_xlnm.Print_Titles", value, sheet: sheet_name)
    end

    # Set workbook protection.
    def set_workbook_protection(**opts)
      @workbook_protection = opts
    end

    # Set a core document property.
    def set_core_property(name, value)
      @core_properties[name] = value
    end

    # Set an app document property.
    def set_app_property(name, value)
      @app_properties[name] = value
    end

    # Add a custom document property.
    def add_custom_property(name, value, type: :string)
      @custom_properties << { name: name, value: value, type: type }
    end

    def build
      # Process styles from all sheets and collect style definitions
      processed_sheets, styles_definition = process_styles(@sheets)

      # Store workbook-level metadata in unmapped_data
      wb_meta = {}
      wb_meta[:defined_names] = resolve_defined_names(@defined_names, processed_sheets) unless @defined_names.empty?
      wb_meta[:core_properties] = @core_properties unless @core_properties.empty?
      wb_meta[:app_properties] = @app_properties unless @app_properties.empty?
      wb_meta[:custom_properties] = @custom_properties unless @custom_properties.empty?
      wb_meta[:workbook_protection] = @workbook_protection if @workbook_protection

      Elements::Workbook.new(
        sheets: processed_sheets,
        styles: styles_definition,
        unmapped_data: wb_meta.empty? ? {} : { facade: wb_meta }
      )
    end

    private

    def absolute_range(range)
      range.gsub(/([A-Z]+)(\d+)/, '$\1$\2')
    end

    def resolve_defined_names(names, sheets)
      sheet_names = sheets.map(&:name)
      names.map do |dn|
        resolved = dn.dup
        if dn[:local_sheet_name]
          idx = sheet_names.index(dn[:local_sheet_name])
          resolved[:local_sheet_id] = idx if idx
          resolved.delete(:local_sheet_name)
        end
        resolved
      end
    end

    def process_styles(sheets)
      # Collect all unique StyleBuilders from all sheets
      all_style_builders = {}
      @sheet_builders.each do |sheet_builder|
        sheet_builder.styles.each do |style_name, style_builder|
          all_style_builders[style_name] = style_builder
        end
      end

      return [sheets, {}] if all_style_builders.empty?

      # Create a temporary writer to register styles and get numeric IDs
      temp_writer = Ooxml::Writer.new
      style_name_to_id = {}
      all_style_builders.each do |name, builder|
        style_id = builder.register_with(temp_writer)
        style_name_to_id[name] = style_id
      end

      # Capture the style definitions from the temporary writer
      styles_definition = extract_styles_from_writer(temp_writer)

      # Update all cells with their resolved style IDs
      updated_sheets = sheets.map do |sheet|
        new_rows = sheet.rows.map do |row|
          new_cells = row.cells.map do |cell|
            # If style_index is a string (style name), resolve it to a numeric ID
            if cell.style_index.is_a?(String) && style_name_to_id.key?(cell.style_index)
              Elements::Cell.new(
                row_index: cell.row_index,
                column_index: cell.column_index,
                value: cell.value,
                formula: cell.formula,
                style_index: style_name_to_id[cell.style_index],
                unmapped_data: cell.unmapped_data,
                errors: cell.errors
              )
            else
              cell
            end
          end
          Elements::Row.new(
            index: row.index,
            cells: new_cells,
            height: row.height,
            hidden: row.hidden,
            custom_height: row.custom_height,
            outline_level: row.outline_level,
            unmapped_data: row.unmapped_data,
            errors: row.errors
          )
        end
        Elements::Worksheet.new(
          name: sheet.name,
          rows: new_rows,
          columns: sheet.columns,
          charts: sheet.charts,
          unmapped_data: sheet.unmapped_data,
          errors: sheet.errors
        )
      end

      [updated_sheets, styles_definition]
    end

    def extract_styles_from_writer(writer)
      # Extract style definitions from the writer that can be reused
      # This captures the fonts, fills, borders, and xf entries that were created
      {
        fonts: writer.instance_variable_get(:@fonts).dup,
        fills: writer.instance_variable_get(:@fills).dup,
        borders: writer.instance_variable_get(:@borders).dup,
        xf_entries: writer.instance_variable_get(:@xf_entries).dup,
        num_fmts: writer.instance_variable_get(:@num_fmts).dup
      }
    end
  end

  # DSL context for a single worksheet in Xlsxrb.build.
  class WorksheetBuilder
    def initialize(name)
      @name = name
      @rows = []
      @columns = []
      @charts = []
      @styles = {} # { style_name => StyleBuilder }
      @style_index_map = {} # { style_name => xf_index } (populated at build time)
      @hyperlinks = []
      @auto_filter = nil
      @filter_columns = {}
      @sort_state = nil
      @data_validations = []
      @conditional_formats = []
      @tables = []
      @comments = []
      @merge_cells_ranges = []
      @freeze_pane = nil
      @split_pane = nil
      @selection = nil
      @page_margins = nil
      @page_setup = {}
      @header_footer = {}
      @print_options = {}
      @sheet_protection = nil
      @images = []
      @shapes = []
      @sheet_properties = {}
      @sheet_view = {}
      @row_breaks = []
      @col_breaks = []
    end

    # Define a named style that can be applied to cells.
    def add_style(name, **opts, &block)
      style_builder = StyleBuilder.new(name)
      style_builder.apply_options!(**opts) unless opts.empty?
      block.call(style_builder) if block_given?
      @styles[name] = style_builder
      style_builder
    end

    # Add a row of values to the sheet.
    # values:: Array of cell values
    # styles:: Hash mapping column indices to style names, or Array of style names for each column
    def add_row(values, styles: nil, height: nil, hidden: false, custom_height: false, outline_level: nil)
      row_index = @rows.size
      cells = values.each_with_index.map do |val, col_index|
        style_name = styles[col_index] if styles.is_a?(Hash) || styles.is_a?(Array)
        # Style index will be resolved at build time
        Elements::Cell.new(
          row_index: row_index,
          column_index: col_index,
          value: val,
          style_index: style_name # Will store the style name for now
        )
      end

      @rows << Elements::Row.new(
        index: row_index,
        cells: cells,
        height: height,
        hidden: hidden,
        custom_height: custom_height || !height.nil?,
        outline_level: outline_level
      )
    end

    # Set column width for a 0-based column index.
    def set_column(index, width: nil, hidden: false, custom_width: false, outline_level: nil)
      @columns << Elements::Column.new(
        index: index,
        width: width,
        hidden: hidden,
        custom_width: custom_width || !width.nil?,
        outline_level: outline_level
      )
    end

    # Add a chart to the sheet.
    def add_chart(**options, &block)
      if block_given?
        builder = ChartBuilder.new
        block.call(builder)
        options = builder.options.merge(options)
      end
      @charts << options
    end

    # --- Hyperlinks ---

    # Add a hyperlink on a cell.
    def add_hyperlink(cell, url = nil, display: nil, tooltip: nil, location: nil)
      link = { cell: cell }
      link[:url] = url if url
      link[:display] = display if display
      link[:tooltip] = tooltip if tooltip
      link[:location] = location if location
      @hyperlinks << link
    end

    # --- Auto Filter / Sort ---

    # Set an auto filter range (e.g. "A1:D10").
    # rubocop:disable Naming/AccessorMethodName
    def set_auto_filter(range)
      @auto_filter = range
    end
    # rubocop:enable Naming/AccessorMethodName

    # Add a filter column to the auto filter.
    def add_filter_column(col_id, filter)
      @filter_columns[col_id] = filter
    end

    # Set sort state.
    def set_sort_state(ref, sort_conditions, **opts)
      @sort_state = { ref: ref, sort_conditions: sort_conditions }.merge(opts)
    end

    # --- Data Validation ---

    # Add a data validation rule.
    def add_data_validation(sqref, **opts)
      @data_validations << opts.merge(sqref: sqref)
    end

    # --- Conditional Formatting ---

    # Add a conditional formatting rule.
    def add_conditional_format(sqref, **opts)
      @conditional_formats << opts.merge(sqref: sqref)
    end

    # --- Tables ---

    # Add a table to the sheet.
    def add_table(ref, columns:, name: nil, display_name: nil, style: nil, **opts)
      tbl = { ref: ref, columns: columns }
      tbl[:name] = name if name
      tbl[:display_name] = display_name if display_name
      tbl[:style] = style if style
      tbl.merge!(opts)
      @tables << tbl
    end

    # --- Comments ---

    # Add a comment on a cell.
    def add_comment(cell, text, author: "Author")
      @comments << { cell: cell, text: text, author: author }
    end

    # --- Merge Cells ---

    # Merge a range of cells (e.g. "A1:B2").
    def merge_cells(range)
      @merge_cells_ranges << range
    end

    # --- Freeze / Split Panes ---

    # Freeze panes at the given row and column.
    def set_freeze_pane(row: 0, col: 0)
      @freeze_pane = { row: row, col: col }
    end

    # Split panes (non-frozen).
    def set_split_pane(x_split: 0, y_split: 0, top_left_cell: nil)
      @split_pane = { x_split: x_split, y_split: y_split, top_left_cell: top_left_cell }
    end

    # Set active cell selection.
    def set_selection(active_cell, sqref: nil, pane: nil)
      @selection = { active_cell: active_cell, sqref: sqref || active_cell }
      @selection[:pane] = pane if pane
    end

    # --- Page Setup / Margins / Print ---

    # Set page margins (in inches).
    def set_page_margins(left: nil, right: nil, top: nil, bottom: nil, header: nil, footer: nil)
      @page_margins = { left: left, right: right, top: top, bottom: bottom, header: header, footer: footer }.compact
    end

    # Set page setup properties.
    def set_page_setup(**opts)
      @page_setup.merge!(opts)
    end

    # Set header/footer text.
    def set_header_footer(**opts)
      @header_footer.merge!(opts)
    end

    # Set a print option.
    def set_print_option(name, value)
      @print_options[name] = value
    end

    # --- Sheet Protection ---

    # Set sheet protection options.
    def set_sheet_protection(**opts)
      normalized = opts.dup
      plain_password = normalized[:password]
      needs_hash = plain_password.is_a?(String) && !plain_password.empty? &&
                   normalized[:algorithm_name].nil? && normalized[:hash_value].nil? &&
                   normalized[:salt_value].nil? && normalized[:spin_count].nil? &&
                   !plain_password.match?(/\A[0-9A-Fa-f]{4}\z/)
      if needs_hash
        normalized.delete(:password)
        normalized.merge!(Xlsxrb::Ooxml::Utils.hash_password(plain_password))
      end
      @sheet_protection = normalized
    end

    # --- Images ---

    # Insert an image from raw file data.
    def add_image(file_data, ext: "png", from_col: 0, from_row: 0, to_col: 5, to_row: 10, **opts)
      img = { file_data: file_data, ext: ext, from_col: from_col, from_row: from_row, to_col: to_col, to_row: to_row }
      img.merge!(opts)
      @images << img
    end

    # --- Shapes ---

    # Add a shape to the sheet.
    def add_shape(preset: "rect", text: nil, from_col: 0, from_row: 0, to_col: 5, to_row: 5, **opts)
      shape = { preset: preset, text: text, from_col: from_col, from_row: from_row, to_col: to_col, to_row: to_row }
      shape[:name] = opts.delete(:name) || "Shape #{@shapes.size + 1}"
      shape.merge!(opts)
      @shapes << shape
    end

    # --- Sheet Properties ---

    # Set a sheet-level property (e.g. :tab_color).
    def set_sheet_property(name, value)
      @sheet_properties[name] = value
    end

    # Set a sheet view property (e.g. :show_grid_lines, :zoom_scale).
    def set_sheet_view(name, value)
      @sheet_view[name] = value
    end

    # --- Row / Column Breaks ---

    # Add a page break before a row.
    def add_row_break(row_num)
      @row_breaks << row_num
    end

    # Add a page break before a column.
    def add_col_break(col_index)
      @col_breaks << col_index
    end

    def build
      facade_meta = {}
      facade_meta[:hyperlinks] = @hyperlinks unless @hyperlinks.empty?
      facade_meta[:auto_filter] = @auto_filter if @auto_filter
      facade_meta[:filter_columns] = @filter_columns unless @filter_columns.empty?
      facade_meta[:sort_state] = @sort_state if @sort_state
      facade_meta[:data_validations] = @data_validations unless @data_validations.empty?
      facade_meta[:conditional_formats] = @conditional_formats unless @conditional_formats.empty?
      facade_meta[:tables] = @tables unless @tables.empty?
      facade_meta[:comments] = @comments unless @comments.empty?
      facade_meta[:merge_cells] = @merge_cells_ranges unless @merge_cells_ranges.empty?
      facade_meta[:freeze_pane] = @freeze_pane if @freeze_pane
      facade_meta[:split_pane] = @split_pane if @split_pane
      facade_meta[:selection] = @selection if @selection
      facade_meta[:page_margins] = @page_margins if @page_margins
      facade_meta[:page_setup] = @page_setup unless @page_setup.empty?
      facade_meta[:header_footer] = @header_footer unless @header_footer.empty?
      facade_meta[:print_options] = @print_options unless @print_options.empty?
      facade_meta[:sheet_protection] = @sheet_protection if @sheet_protection
      facade_meta[:images] = @images unless @images.empty?
      facade_meta[:shapes] = @shapes unless @shapes.empty?
      facade_meta[:sheet_properties] = @sheet_properties unless @sheet_properties.empty?
      facade_meta[:sheet_view] = @sheet_view unless @sheet_view.empty?
      facade_meta[:row_breaks] = @row_breaks unless @row_breaks.empty?
      facade_meta[:col_breaks] = @col_breaks unless @col_breaks.empty?

      Elements::Worksheet.new(
        name: @name, rows: @rows, columns: @columns, charts: @charts,
        unmapped_data: facade_meta.empty? ? {} : { facade: facade_meta }
      )
    end

    # Internal: returns styles for later processing by WorkbookBuilder
    attr_reader :styles
  end

  # DSL context for Xlsxrb.generate streaming writes.
  class StreamWriter
    def initialize(target)
      @target = target
      @sst = []
      @sst_index = {}
      @sheets = []
      @current_sheet = nil
      @current_row_index = 0
      @tempfiles = []
      @current_tempfile = nil
      @current_row_writer = nil
      @current_columns = []
      @current_charts = []
      @current_hyperlinks = []
      @current_auto_filter = nil
      @current_filter_columns = {}
      @current_sort_state = nil
      @current_data_validations = []
      @current_conditional_formats = []
      @current_tables = []
      @current_comments = []
      @current_merge_cells = []
      @current_freeze_pane = nil
      @current_split_pane = nil
      @current_selection = nil
      @current_page_margins = nil
      @current_page_setup = {}
      @current_header_footer = {}
      @current_print_options = {}
      @current_sheet_protection = nil
      @current_images = []
      @current_shapes = []
      @current_sheet_properties = {}
      @current_sheet_view = {}
      @current_row_breaks = []
      @current_col_breaks = []
      @styles = {} # { style_name => StyleBuilder }
      @style_writer = Ooxml::Writer.new
      @style_name_to_id = {}
      # Workbook-level settings
      @defined_names = []
      @core_properties = {}
      @app_properties = {}
      @custom_properties = []
      @workbook_protection = nil
    end

    # Define a named style that can be applied to cells.
    def add_style(name, **opts, &block)
      style_builder = StyleBuilder.new(name)
      style_builder.apply_options!(**opts) unless opts.empty?
      block.call(style_builder) if block_given?
      @styles[name] = style_builder

      # Register immediately
      @style_name_to_id[name] = style_builder.register_with(@style_writer)

      style_builder
    end

    # Start or switch to a named sheet.
    def add_sheet(name = nil)
      flush_current_sheet
      name ||= "Sheet#{@sheets.size + 1}"
      @current_sheet = name
      @current_row_index = 0
      @current_tempfile = Tempfile.new(["xlsxrb_rows", ".xml"])
      @current_tempfile.binmode
      @current_row_writer = Ooxml::WorksheetWriter.new(@current_tempfile)
      @current_row_writer.instance_variable_set(:@started, true)

      @current_columns = []
      @current_charts = []
      @current_hyperlinks = []
      @current_auto_filter = nil
      @current_filter_columns = {}
      @current_sort_state = nil
      @current_data_validations = []
      @current_conditional_formats = []
      @current_tables = []
      @current_comments = []
      @current_merge_cells = []
      @current_freeze_pane = nil
      @current_split_pane = nil
      @current_selection = nil
      @current_page_margins = nil
      @current_page_setup = {}
      @current_header_footer = {}
      @current_print_options = {}
      @current_sheet_protection = nil
      @current_images = []
      @current_shapes = []
      @current_sheet_properties = {}
      @current_sheet_view = {}
      @current_row_breaks = []
      @current_col_breaks = []

      return unless block_given?

      yield self
      flush_current_sheet
    end

    # Add a row of values. values is an Array.
    # styles:: Hash mapping column indices to style names, or Array of style names for each column
    def add_row(values, styles: nil, height: nil, hidden: false)
      add_sheet if @current_sheet.nil?

      row_index = @current_row_index
      @current_row_index += 1

      attrs = nil
      if height || hidden
        attrs = {}
        attrs[:height] = height if height
        attrs[:hidden] = true if hidden
      end

      @current_row_writer.write_row_values(row_index, values, styles: styles, style_map: @style_name_to_id, sst: @sst, sst_index: @sst_index, attrs: attrs)
    end

    # Set column width for a 0-based column index.
    def set_column(index, width: nil, hidden: false)
      add_sheet if @current_sheet.nil?

      @current_columns << { index: index, width: width, hidden: hidden, custom_width: !width.nil? }
    end

    # Add a chart to the current sheet.
    def add_chart(**options, &block)
      add_sheet if @current_sheet.nil?

      if block_given?
        builder = ChartBuilder.new
        block.call(builder)
        options = builder.options.merge(options)
      end

      @current_charts << options
    end

    # --- Hyperlinks ---

    def add_hyperlink(cell, url = nil, display: nil, tooltip: nil, location: nil)
      add_sheet if @current_sheet.nil?
      link = { cell: cell }
      link[:url] = url if url
      link[:display] = display if display
      link[:tooltip] = tooltip if tooltip
      link[:location] = location if location
      @current_hyperlinks << link
    end

    # --- Auto Filter / Sort ---

    # rubocop:disable Naming/AccessorMethodName
    def set_auto_filter(range)
      add_sheet if @current_sheet.nil?
      @current_auto_filter = range
    end
    # rubocop:enable Naming/AccessorMethodName

    def add_filter_column(col_id, filter)
      add_sheet if @current_sheet.nil?
      @current_filter_columns[col_id] = filter
    end

    def set_sort_state(ref, sort_conditions, **opts)
      add_sheet if @current_sheet.nil?
      @current_sort_state = { ref: ref, sort_conditions: sort_conditions }.merge(opts)
    end

    # --- Data Validation ---

    def add_data_validation(sqref, **opts)
      add_sheet if @current_sheet.nil?
      @current_data_validations << opts.merge(sqref: sqref)
    end

    # --- Conditional Formatting ---

    def add_conditional_format(sqref, **opts)
      add_sheet if @current_sheet.nil?
      @current_conditional_formats << opts.merge(sqref: sqref)
    end

    # --- Tables ---

    def add_table(ref, columns:, name: nil, display_name: nil, style: nil, **opts)
      add_sheet if @current_sheet.nil?
      tbl = { ref: ref, columns: columns }
      tbl[:name] = name if name
      tbl[:display_name] = display_name if display_name
      tbl[:style] = style if style
      tbl.merge!(opts)
      @current_tables << tbl
    end

    # --- Comments ---

    def add_comment(cell, text, author: "Author")
      add_sheet if @current_sheet.nil?
      @current_comments << { cell: cell, text: text, author: author }
    end

    # --- Merge Cells ---

    def merge_cells(range)
      add_sheet if @current_sheet.nil?
      @current_merge_cells << range
    end

    # --- Freeze / Split Panes ---

    def set_freeze_pane(row: 0, col: 0)
      add_sheet if @current_sheet.nil?
      @current_freeze_pane = { row: row, col: col }
    end

    def set_split_pane(x_split: 0, y_split: 0, top_left_cell: nil)
      add_sheet if @current_sheet.nil?
      @current_split_pane = { x_split: x_split, y_split: y_split, top_left_cell: top_left_cell }
    end

    def set_selection(active_cell, sqref: nil, pane: nil)
      add_sheet if @current_sheet.nil?
      @current_selection = { active_cell: active_cell, sqref: sqref || active_cell }
      @current_selection[:pane] = pane if pane
    end

    # --- Page Setup / Margins / Print ---

    def set_page_margins(left: nil, right: nil, top: nil, bottom: nil, header: nil, footer: nil)
      add_sheet if @current_sheet.nil?
      @current_page_margins = { left: left, right: right, top: top, bottom: bottom, header: header, footer: footer }.compact
    end

    def set_page_setup(**opts)
      add_sheet if @current_sheet.nil?
      @current_page_setup.merge!(opts)
    end

    def set_header_footer(**opts)
      add_sheet if @current_sheet.nil?
      @current_header_footer.merge!(opts)
    end

    def set_print_option(name, value)
      add_sheet if @current_sheet.nil?
      @current_print_options[name] = value
    end

    # --- Sheet Protection ---

    def set_sheet_protection(**opts)
      add_sheet if @current_sheet.nil?
      normalized = opts.dup
      plain_password = normalized[:password]
      needs_hash = plain_password.is_a?(String) && !plain_password.empty? &&
                   normalized[:algorithm_name].nil? && normalized[:hash_value].nil? &&
                   normalized[:salt_value].nil? && normalized[:spin_count].nil? &&
                   !plain_password.match?(/\A[0-9A-Fa-f]{4}\z/)
      if needs_hash
        normalized.delete(:password)
        normalized.merge!(Xlsxrb::Ooxml::Utils.hash_password(plain_password))
      end
      @current_sheet_protection = normalized
    end

    # --- Images ---

    def add_image(file_data, ext: "png", from_col: 0, from_row: 0, to_col: 5, to_row: 10, **opts)
      add_sheet if @current_sheet.nil?
      img = { file_data: file_data, ext: ext, from_col: from_col, from_row: from_row, to_col: to_col, to_row: to_row }
      img.merge!(opts)
      @current_images << img
    end

    # --- Shapes ---

    def add_shape(preset: "rect", text: nil, from_col: 0, from_row: 0, to_col: 5, to_row: 5, **opts)
      add_sheet if @current_sheet.nil?
      shape = { preset: preset, text: text, from_col: from_col, from_row: from_row, to_col: to_col, to_row: to_row }
      shape[:name] = opts.delete(:name) || "Shape #{@current_shapes.size + 1}"
      shape.merge!(opts)
      @current_shapes << shape
    end

    # --- Sheet Properties ---

    def set_sheet_property(name, value)
      add_sheet if @current_sheet.nil?
      @current_sheet_properties[name] = value
    end

    def set_sheet_view(name, value)
      add_sheet if @current_sheet.nil?
      @current_sheet_view[name] = value
    end

    # --- Row / Column Breaks ---

    def add_row_break(row_num)
      add_sheet if @current_sheet.nil?
      @current_row_breaks << row_num
    end

    def add_col_break(col_index)
      add_sheet if @current_sheet.nil?
      @current_col_breaks << col_index
    end

    # --- Workbook-Level Methods ---

    # Add a defined name.
    def add_defined_name(name, value, sheet: nil, hidden: false)
      entry = { name: name, value: value, hidden: hidden }
      if sheet
        # local_sheet_id will be resolved at close time
        entry[:local_sheet_name] = sheet
      end
      @defined_names << entry
    end

    # Set the print area for the current or named sheet.
    def set_print_area(range, sheet: nil)
      sheet_name = sheet || @current_sheet || "Sheet1"
      value = "'#{sheet_name}'!#{absolute_range(range)}"
      @defined_names.reject! { |dn| dn[:name] == "_xlnm.Print_Area" && dn[:local_sheet_name] == sheet_name }
      add_defined_name("_xlnm.Print_Area", value, sheet: sheet_name)
    end

    # Set print titles for the current or named sheet.
    def set_print_titles(rows: nil, cols: nil, sheet: nil)
      sheet_name = sheet || @current_sheet || "Sheet1"
      parts = []
      parts << "'#{sheet_name}'!$#{cols.sub(":", ":$")}" if cols
      parts << "'#{sheet_name}'!$#{rows.sub(":", ":$")}" if rows
      value = parts.join(",")
      @defined_names.reject! { |dn| dn[:name] == "_xlnm.Print_Titles" && dn[:local_sheet_name] == sheet_name }
      add_defined_name("_xlnm.Print_Titles", value, sheet: sheet_name)
    end

    # Set workbook protection.
    def set_workbook_protection(**opts)
      @workbook_protection = opts
    end

    # Set a core document property.
    def set_core_property(name, value)
      @core_properties[name] = value
    end

    # Set an app document property.
    def set_app_property(name, value)
      @app_properties[name] = value
    end

    # Add a custom document property.
    def add_custom_property(name, value, type: :string)
      @custom_properties << { name: name, value: value, type: type }
    end

    def close
      TRACER.in_span("StreamWriter#close") do
        flush_current_sheet

        styles_definition = {
          fonts: @style_writer.instance_variable_get(:@fonts).dup,
          fills: @style_writer.instance_variable_get(:@fills).dup,
          borders: @style_writer.instance_variable_get(:@borders).dup,
          xf_entries: @style_writer.instance_variable_get(:@xf_entries).dup,
          num_fmts: @style_writer.instance_variable_get(:@num_fmts).dup
        }

        resolved_names = resolve_defined_names(@defined_names, @sheets)

        Ooxml::WorkbookWriter.write(
          @target,
          sheets: @sheets,
          shared_strings: @sst,
          styles: styles_definition,
          defined_names: resolved_names.empty? ? nil : resolved_names,
          core_properties: @core_properties.empty? ? nil : @core_properties,
          app_properties: @app_properties.empty? ? nil : @app_properties,
          custom_properties: @custom_properties.empty? ? nil : @custom_properties,
          workbook_protection: @workbook_protection
        )
      end
    ensure
      @tempfiles.each do |tmp|
        tmp.close
        tmp.unlink
      end
    end

    private

    def absolute_range(range)
      range.gsub(/([A-Z]+)(\d+)/, '$\1$\2')
    end

    def resolve_defined_names(names, sheets)
      sheet_names = sheets.map { |s| s[:name] }
      names.map do |dn|
        resolved = dn.dup
        if dn[:local_sheet_name]
          idx = sheet_names.index(dn[:local_sheet_name])
          resolved[:local_sheet_id] = idx if idx
          resolved.delete(:local_sheet_name)
        end
        resolved
      end
    end

    def flush_current_sheet
      return unless @current_sheet

      @current_tempfile.close

      sheet_data = { name: @current_sheet, rows_tmp_path: @current_tempfile.path, columns: @current_columns }
      sheet_data[:charts] = @current_charts unless @current_charts.empty?
      sheet_data[:hyperlinks] = @current_hyperlinks unless @current_hyperlinks.empty?
      sheet_data[:auto_filter] = @current_auto_filter if @current_auto_filter
      sheet_data[:filter_columns] = @current_filter_columns unless @current_filter_columns.empty?
      sheet_data[:sort_state] = @current_sort_state if @current_sort_state
      sheet_data[:data_validations] = @current_data_validations unless @current_data_validations.empty?
      sheet_data[:conditional_formats] = @current_conditional_formats unless @current_conditional_formats.empty?
      sheet_data[:tables] = @current_tables unless @current_tables.empty?
      sheet_data[:comments] = @current_comments unless @current_comments.empty?
      sheet_data[:merge_cells] = @current_merge_cells unless @current_merge_cells.empty?
      sheet_data[:freeze_pane] = @current_freeze_pane if @current_freeze_pane
      sheet_data[:split_pane] = @current_split_pane if @current_split_pane
      sheet_data[:selection] = @current_selection if @current_selection
      sheet_data[:page_margins] = @current_page_margins if @current_page_margins
      sheet_data[:page_setup] = @current_page_setup unless @current_page_setup.empty?
      sheet_data[:header_footer] = @current_header_footer unless @current_header_footer.empty?
      sheet_data[:print_options] = @current_print_options unless @current_print_options.empty?
      sheet_data[:sheet_protection] = @current_sheet_protection if @current_sheet_protection
      sheet_data[:images] = @current_images unless @current_images.empty?
      sheet_data[:shapes] = @current_shapes unless @current_shapes.empty?
      sheet_data[:sheet_properties] = @current_sheet_properties unless @current_sheet_properties.empty?
      sheet_data[:sheet_view] = @current_sheet_view unless @current_sheet_view.empty?
      sheet_data[:row_breaks] = @current_row_breaks unless @current_row_breaks.empty?
      sheet_data[:col_breaks] = @current_col_breaks unless @current_col_breaks.empty?
      @sheets << sheet_data

      @tempfiles << @current_tempfile
      @current_sheet = nil
      @current_tempfile = nil
      @current_row_writer = nil
    end
  end

  class << self
    private

    def build_worksheet(name, sheet_xml, shared_strings, _styles)
      return Elements::Worksheet.new(name: name) if sheet_xml.nil? || sheet_xml.empty?

      raw_rows = Ooxml::WorksheetParser.parse(sheet_xml, shared_strings: shared_strings)
      raw_columns = Ooxml::WorksheetParser.parse_columns(sheet_xml)

      rows = raw_rows.map { |rr| build_row_from_raw(rr) }
      columns = raw_columns.map do |rc|
        # Columns from OOXML are 1-based min/max ranges; convert to 0-based
        Elements::Column.new(
          index: (rc[:min] || 1) - 1,
          width: rc[:width],
          hidden: rc[:hidden] || false,
          custom_width: rc[:custom_width] || false,
          outline_level: rc[:outline_level]
        )
      end

      Elements::Worksheet.new(name: name, rows: rows, columns: columns)
    end

    def build_row_from_raw(raw_row)
      cells = raw_row[:cells].map do |rc|
        parsed = Elements::Cell.parse_ref(rc[:ref]) if rc[:ref]
        row_idx = parsed ? parsed[0] : raw_row[:index]
        col_idx = parsed ? parsed[1] : 0

        Elements::Cell.new(
          row_index: row_idx,
          column_index: col_idx,
          value: rc[:value],
          formula: rc[:formula],
          style_index: rc[:style_index]
        )
      end
      attrs = raw_row[:attrs] || {}
      Elements::Row.new(
        index: raw_row[:index],
        cells: cells,
        height: attrs[:height],
        hidden: attrs[:hidden] || false,
        custom_height: attrs[:custom_height] || false,
        outline_level: attrs[:outline_level]
      )
    end

    def build_raw_cell(cell, sst, sst_index)
      ref = cell.ref
      value = cell.value
      result = { ref: ref, style_index: cell.style_index }

      case value
      when String
        idx = sst_index[value] ||= begin
          sst << value
          sst.size - 1
        end
        result[:value] = idx
        result[:type] = "s"
      when true, false
        result[:value] = value
        result[:type] = "b"
      when Integer, Float
        result[:value] = value
      when Date
        result[:value] = Xlsxrb::Ooxml::Utils.date_to_serial(value)
      when Time
        result[:value] = Xlsxrb::Ooxml::Utils.datetime_to_serial(value)
      when NilClass
        # empty cell
      end

      result[:formula] = cell.formula if cell.formula
      result
    end

    def build_row_attrs(row)
      attrs = {}
      attrs[:height] = row.height if row.height
      attrs[:hidden] = true if row.hidden
      attrs[:custom_height] = true if row.custom_height
      attrs
    end
  end

  # Builds a raw cell hash from a value for streaming writes.
  def self.build_raw_cell_from_value(row_index, col_index, value, sst, sst_index)
    ref = "#{Elements::Cell.column_letter(col_index)}#{row_index + 1}"
    result = { ref: ref }

    case value
    when String
      idx = sst_index[value] ||= begin
        sst << value
        sst.size - 1
      end
      result[:value] = idx
      result[:type] = "s"
    when true, false
      result[:value] = value
      result[:type] = "b"
    when Integer, Float
      result[:value] = value
    when Date
      result[:value] = Xlsxrb::Ooxml::Utils.date_to_serial(value)
    when Time
      result[:value] = Xlsxrb::Ooxml::Utils.datetime_to_serial(value)
    when NilClass
      # empty cell
    end

    result
  end
end
