# frozen_string_literal: true

# rbs_inline: enabled

require "date"
require "time"
require "openssl"
require "securerandom"
require "tempfile"
require "bigdecimal"
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
  class Error < StandardError
    DIV0  = Elements::CellError.new(code: "#DIV/0!")
    NA    = Elements::CellError.new(code: "#N/A")
    NAME  = Elements::CellError.new(code: "#NAME?")
    NULL  = Elements::CellError.new(code: "#NULL!")
    NUM   = Elements::CellError.new(code: "#NUM!")
    REF   = Elements::CellError.new(code: "#REF!")
    VALUE = Elements::CellError.new(code: "#VALUE!")
  end

  class ParseError < Error; end
  class ValidationError < Error; end
  class ZipError < Error; end

  TRACER = OpenTelemetry.tracer_provider.tracer("xlsxrb", Xlsxrb::VERSION)

  def self.in_span(name, attributes: nil, &)
    if defined?(Ractor) && Ractor.current != Ractor.main
      yield
    elsif attributes
      TRACER.in_span(name, attributes: attributes, &)
    else
      TRACER.in_span(name, &)
    end
  end

  # Helper to easily create RichText objects.
  # Supports both `Xlsxrb.rich_text({ text: "A" }, { text: "B" })`
  # and `Xlsxrb.rich_text(text: "Hi", bold: true)`
  #
  # @param runs [Array<Hash>] Optional rich text runs.
  # @param text [String, nil] Simple text.
  # @param font_props [Hash] Font styling options (e.g., bold: true).
  # @return [Elements::RichText] The resulting rich text.
  # : (*untyped runs, ?text: untyped?, **untyped font_props) -> untyped
  def self.rich_text(*runs, text: nil, **font_props)
    runs = [{ text: text, font: font_props }] if text
    Elements::RichText.new(runs: runs)
  end

  # Builder for block-style chart definitions.
  class ChartBuilder
    # : () -> void
    def initialize
      @options = {}
    end

    # : untyped
    attr_reader :options

    # : (untyped value) -> untyped
    def type(value) = @options[:type] = value
    # : (untyped value) -> untyped
    def title(value) = @options[:title] = value

    # : (?untyped? value) ?{ (untyped) -> untyped } -> untyped
    def series(value = nil)
      @options[:series] ||= []
      if block_given?
        sb = SeriesBuilder.new
        yield sb
        @options[:series] << sb.options
      elsif value
        @options[:series] << value
      end
      @options[:series]
    end

    # : (untyped name, *untyped args, **untyped kwargs) { (?) -> untyped } -> untyped
    def method_missing(name, *args, **kwargs, &)
      key = name.to_sym
      @options[key] = kwargs.empty? ? args.first : kwargs
    end

    # : (untyped _name, ?bool _include_private) -> true
    def respond_to_missing?(_name, _include_private = false)
      true
    end

    # Builder for a single series entry in block-style chart definitions.
    class SeriesBuilder
      # : () -> void
      def initialize
        @options = {}
      end

      # : untyped
      attr_reader :options

      # : (untyped name, *untyped args, **untyped kwargs) { (?) -> untyped } -> untyped
      def method_missing(name, *args, **kwargs, &)
        key = name.to_sym
        @options[key] = kwargs.empty? ? args.first : kwargs
      end

      # : (untyped _name, ?bool _include_private) -> true
      def respond_to_missing?(_name, _include_private = false)
        true
      end
    end
  end

  # Generic builder for block-style feature definitions.
  # Supports method_missing for setting arbitrary keys.
  # --- Facade API ---

  # Creates a Formula object for use in row values.
  #
  # @param expression [String] The formula text (e.g. "SUM(A1:A10)").
  # @param cached_value [Object, nil] Optional cached result. If nil, Excel will calculate on open.
  # @return [Elements::Formula]
  # : (untyped expression, ?cached_value: untyped?) -> untyped
  def self.formula(expression, cached_value: nil)
    Elements::Formula.new(
      expression: expression,
      cached_value: cached_value,
      calculate_always: cached_value.nil? || nil
    )
  end

  # Reads an XLSX file into an Elements::Workbook.
  #
  # @param source [String, IO] File path or IO object.
  # @return [Elements::Workbook] The parsed workbook.
  # : (untyped source) -> untyped
  def self.read(source)
    attributes = source.is_a?(String) ? { "filepath" => source } : {}
    Xlsxrb.in_span("Xlsxrb.read", attributes: attributes) do
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
  #
  # @param target [String, IO] File path or IO object.
  # @param workbook [Elements::Workbook] The workbook to write.
  # @return [void]
  # : (untyped target, untyped workbook) -> untyped
  def self.write(target, workbook)
    raise Error, "target is required" if target.nil?
    raise Error, "workbook must be an Elements::Workbook" unless workbook.is_a?(Elements::Workbook)

    attributes = target.is_a?(String) ? { "filepath" => target } : {}
    Xlsxrb.in_span("Xlsxrb.write", attributes: attributes) do
      sst = []
      sst_index = {}

      # Collect shared strings and build index without allocating new Hashes
      sheet_data = workbook.sheets.map do |ws|
        ws.rows.each do |row|
          row.cells.each do |cell|
            val = cell.value
            if (val.is_a?(String) || val.is_a?(Elements::RichText)) && !sst_index.key?(val)
              sst << val
              sst_index[val] = sst.size - 1
            end
          end
        end
        columns = ws.columns.map do |col|
          { index: col.index, width: col.width, hidden: col.hidden, custom_width: col.custom_width, outline_level: col.outline_level }
        end
        sd = { name: ws.name, rows: ws.rows, columns: columns }
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
        shared_strings_index: sst_index,
        styles: workbook.styles,
        defined_names: wb_facade[:defined_names],
        core_properties: wb_facade[:core_properties],
        app_properties: wb_facade[:app_properties],
        custom_properties: wb_facade[:custom_properties],
        workbook_protection: wb_facade[:workbook_protection],
        workbook_properties: wb_facade[:workbook_properties]
      )
    end
  end

  # Modifies an existing XLSX file.
  # Reads the workbook, passes it to the block, and writes the result.
  # The block receives an Elements::Workbook and must return a modified one (e.g. via `with`).
  # If no target is given, the source is overwritten.
  #
  # @example
  #   Xlsxrb.modify("template.xlsx", "output.xlsx") do |wb|
  #     sheet = wb.sheet(0)
  #     row0 = sheet.row_at(0)
  #     new_cell = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 1, value: "Updated")
  #     new_row = row0.with(cells: row0.cells.map { |c| c.column_index == 1 ? new_cell : c })
  #     new_sheet = sheet.with(rows: sheet.rows.map { |r| r.index == 0 ? new_row : r })
  #     wb.with(sheets: wb.sheets.map.with_index { |s, i| i == 0 ? new_sheet : s })
  #   end
  #
  # @param source [String, IO] The source file path or IO object.
  # @param target [String, IO, nil] The target file path or IO object. If nil, overwrites source.
  # @yield [workbook] Yields the parsed workbook.
  # @yieldparam workbook [Elements::Workbook] The parsed workbook.
  # @yieldreturn [Elements::Workbook] The modified workbook.
  # @return [void]
  # : (untyped source, ?untyped? target) ?{ (untyped) -> untyped } -> untyped
  def self.modify(source, target = nil)
    raise Error, "source is required" if source.nil?
    raise Error, "block is required" unless block_given?

    workbook = read(source)
    result_workbook = yield workbook
    result_workbook = workbook unless result_workbook.is_a?(Elements::Workbook)

    write_target = target || source
    write(write_target, result_workbook)
  end

  # Represents a sheet being streamed sequentially.
  class StreamSheet
    include Enumerable

    attr_reader :name

    def initialize(name, sheet_xml, shared_strings)
      @name = name
      @sheet_xml = sheet_xml
      @shared_strings = shared_strings
    end

    def each_row
      return enum_for(:each_row) unless block_given?

      Ooxml::WorksheetParser.each_row(@sheet_xml, shared_strings: @shared_strings) do |raw_row|
        yield Xlsxrb.send(:build_row_from_raw, raw_row)
      end
    end
    alias each each_row
  end

  # Streaming read: yields StreamSheet objects one at a time for each sheet.
  #
  # @param source [String, IO] File path or IO object.
  # @yield [sheet] Yields each sheet.
  # @yieldparam sheet [StreamSheet] The streaming sheet object.
  # @return [Enumerator] If no block is given.
  # @return [void]
  # : (untyped source) ?{ (untyped) -> untyped } -> untyped
  def self.foreach(source)
    return enum_for(:foreach, source) unless block_given?

    attributes = source.is_a?(String) ? { "filepath" => source } : {}
    Xlsxrb.in_span("Xlsxrb.foreach", attributes: attributes) do
      entries = Ooxml::ZipReader.open(source, &:read_all)
      shared_strings = Ooxml::SharedStringsParser.parse(entries["xl/sharedStrings.xml"])
      workbook_sheets = Ooxml::WorkbookParser.parse(entries["xl/workbook.xml"])
      rels = Ooxml::RelationshipsParser.parse(entries["xl/_rels/workbook.xml.rels"])

      workbook_sheets.each do |sheet_info|
        target = rels[sheet_info[:r_id]]
        next unless target

        sheet_path = target.start_with?("/") ? target.delete_prefix("/") : "xl/#{target}"
        sheet_xml = entries[sheet_path]
        next if sheet_xml.nil? || sheet_xml.empty?

        yield StreamSheet.new(sheet_info[:name], sheet_xml, shared_strings)
      end
    end
  end

  # Streaming write: yields a StreamWriter context for building XLSX on-the-fly.
  #
  # @param target [String, IO] File path or IO object.
  # @yield [stream_writer]
  # @yieldparam stream_writer [StreamWriter]
  # @return [void]
  # : (untyped target) ?{ (untyped) -> untyped } -> untyped
  def self.generate(target, strict_excel_mode: true)
    raise Error, "target is required" if target.nil?
    raise Error, "block is required" unless block_given?

    attributes = target.is_a?(String) ? { "filepath" => target } : {}
    Xlsxrb.in_span("Xlsxrb.generate", attributes: attributes) do
      stream_writer = StreamWriter.new(target, strict_excel_mode: strict_excel_mode)
      begin
        yield stream_writer
        stream_writer.close
      ensure
        stream_writer.cleanup!
      end
    end
  end

  # Builds an Elements::Workbook in memory using a DSL.
  #
  # @yield [builder]
  # @yieldparam builder [WorkbookBuilder]
  # @return [Elements::Workbook]
  # : () ?{ (untyped) -> untyped } -> untyped
  def self.build(strict_excel_mode: true)
    raise Error, "block is required" unless block_given?

    Xlsxrb.in_span("Xlsxrb.build") do
      builder = WorkbookBuilder.new(strict_excel_mode: strict_excel_mode)
      yield builder
      builder.build
    end
  end

  # DSL context for Xlsxrb.build.
  class WorkbookBuilder
    # : (?strict_excel_mode: bool) -> void
    def initialize(strict_excel_mode: true)
      @strict_excel_mode = strict_excel_mode
      @sheets = []
      @sheet_builders = [] # Keep track of sheet builders for style processing
      @defined_names = []
      @core_properties = {}
      @app_properties = {}
      @custom_properties = []
      @workbook_protection = nil
      @workbook_properties = { update_links: "never" }
    end

    # Set a workbook property.
    #
    # @note **SECURITY WARNING:** If you set `:update_links` to anything other than `"never"`,
    #   you may expose end-users to malicious external reference vulnerabilities (e.g., CSV/DDE Injection)
    #   when they open the generated Excel file. Ensure you fully trust the exported data.
    #
    # @param name [Symbol] The property name (e.g. :update_links).
    # @param value [String, Integer, Boolean] The property value.
    # @return [void]
    # : (untyped name, untyped value) -> untyped
    def workbook_property(name, value)
      @workbook_properties[name] = value
    end

    # Add a new sheet.
    #
    # @param name [String, nil] The name of the sheet.
    # @param opts [Hash] Sheet properties.
    # @yield [sheet_builder]
    # @yieldparam sheet_builder [WorksheetBuilder]
    # @return [void]
    # : (?untyped? name, **untyped opts) ?{ (untyped) -> untyped } -> untyped
    def sheet(name = nil, **opts)
      name ||= "Sheet#{@sheets.size + 1}"
      if @strict_excel_mode && name.length > 31
        raise ArgumentError, "Sheet name '#{name}' must be <= 31 characters (Excel limitation)"
      end
      if name.match?(/[\[\]\*?\/\\]/)
        raise ArgumentError, "Sheet name '#{name}' contains invalid characters (ECMA-376 OOXML specification)"
      end
      if @strict_excel_mode && @sheets.map { |s| s.respond_to?(:name) ? s.name.downcase : s.to_s.downcase }.include?(name.downcase)
        raise ArgumentError, "Sheet name '#{name}' is already used. Excel requires unique sheet names."
      end

      sheet_builder = WorksheetBuilder.new(name, strict_excel_mode: @strict_excel_mode)
      opts.each { |k, v| sheet_builder.sheet_properties(k, v) }
      yield sheet_builder if block_given?
      @sheet_builders << sheet_builder
      @sheets << sheet_builder.build
    end

    # --- Workbook-Level Methods ---

    # Add a defined name.
    #
    # @param name [String] The defined name.
    # @param value [String] The formula or value.
    # @param sheet [String, nil] Local sheet name.
    # @param hidden [Boolean] Whether the defined name is hidden.
    # @return [void]
    # : (untyped name, untyped value, ?sheet: untyped?, ?hidden: bool) -> untyped
    def defined_name(name, value, sheet: nil, hidden: false)
      entry = { name: name, value: value, hidden: hidden }
      entry[:local_sheet_name] = sheet if sheet
      @defined_names << entry
    end

    # Set the print area for a sheet.
    #
    # @param range [String] The range string (e.g. "A1:B10").
    # @param sheet [String, nil] The sheet name.
    # @return [void]
    # : (untyped range, ?sheet: untyped?) -> untyped
    def print_area(range, sheet: nil)
      sheet_name = sheet || @sheets.last&.name || "Sheet1"
      value = "'#{sheet_name}'!#{absolute_range(range)}"
      @defined_names.reject! { |dn| dn[:name] == "_xlnm.Print_Area" && dn[:local_sheet_name] == sheet_name }
      defined_name("_xlnm.Print_Area", value, sheet: sheet_name)
    end

    # Set print titles for a sheet.
    #
    # @param rows [String, nil] Rows to repeat (e.g. "1:2").
    # @param cols [String, nil] Columns to repeat (e.g. "A:B").
    # @param sheet [String, nil] The sheet name.
    # @return [void]
    # : (?rows: untyped?, ?cols: untyped?, ?sheet: untyped?) -> untyped
    def print_titles(rows: nil, cols: nil, sheet: nil)
      sheet_name = sheet || @sheets.last&.name || "Sheet1"
      parts = []
      parts << "'#{sheet_name}'!$#{cols.sub(":", ":$")}" if cols
      parts << "'#{sheet_name}'!$#{rows.sub(":", ":$")}" if rows
      value = parts.join(",")
      @defined_names.reject! { |dn| dn[:name] == "_xlnm.Print_Titles" && dn[:local_sheet_name] == sheet_name }
      defined_name("_xlnm.Print_Titles", value, sheet: sheet_name)
    end

    # Set workbook protection.
    #
    # @param opts [Hash] Protection options.
    # @return [void]
    # : (**untyped opts) -> untyped
    def protect_workbook(**opts)
      @workbook_protection = opts
    end

    # Set a core document property.
    #
    # @param name [Symbol] The property name.
    # @param value [String, Integer, Time] The property value.
    # @return [void]
    # : (untyped name, untyped value) -> untyped
    def core_property(name, value)
      @core_properties[name] = value
    end

    # Set an app document property.
    #
    # @param name [Symbol] The property name.
    # @param value [String, Integer, Time] The property value.
    # @return [void]
    # : (untyped name, untyped value) -> untyped
    def app_property(name, value)
      @app_properties[name] = value
    end

    # Set multiple core and/or app properties.
    #
    # @param core [Hash, nil] Core properties.
    # @param app [Hash, nil] App properties.
    # @return [void]
    # : (?core: untyped?, ?app: untyped?) -> untyped
    def properties(core: nil, app: nil)
      core&.each { |k, v| core_property(k, v) }
      app&.each { |k, v| app_property(k, v) }
    end

    # Add a custom document property.
    #
    # @param name [String] The property name.
    # @param value [String, Integer, Float, Boolean, Time] The property value.
    # @param type [Symbol] The type of property (:string, :number, :bool, :date).
    # @return [void]
    # : (untyped name, untyped value, ?type: ::Symbol) -> untyped
    def custom_property(name, value, type: :string)
      @custom_properties << { name: name, value: value, type: type }
    end

    # : () -> untyped
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
      wb_meta[:workbook_properties] = @workbook_properties unless @workbook_properties.empty?

      Elements::Workbook.new(
        sheets: processed_sheets,
        styles: styles_definition,
        unmapped_data: wb_meta.empty? ? {} : { facade: wb_meta }
      )
    end

    private

    # : (untyped range) -> untyped
    def absolute_range(range)
      range.gsub(/([A-Z]+)(\d+)/, '$\1$\2')
    end

    # : (untyped names, untyped sheets) -> untyped
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

    # : (untyped sheets) -> (::Array[untyped | ::Hash[untyped, untyped]] | ::Array[untyped])
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

    # : (untyped writer) -> { fonts: untyped, fills: untyped, borders: untyped, xf_entries: untyped, num_fmts: untyped }
    def extract_styles_from_writer(writer)
      # Extract style definitions from the writer that can be reused
      # This captures the fonts, fills, borders, and xf entries that were created
      {
        fonts: writer.fonts.dup,
        fills: writer.fills.dup,
        borders: writer.borders.dup,
        xf_entries: writer.xf_entries.dup,
        num_fmts: writer.num_fmts.dup
      }
    end
  end

  # DSL context for a single worksheet in Xlsxrb.build.
  class WorksheetBuilder
    # : (untyped name, ?strict_excel_mode: bool) -> void
    def initialize(name, strict_excel_mode: true)
      @name = name
      @strict_excel_mode = strict_excel_mode
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
      @sparkline_groups = []
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
    #
    # @param name [String] The name of the style.
    # @param opts [Hash] Style options (e.g. bold: true).
    # @yield [style_builder]
    # @yieldparam style_builder [StyleBuilder]
    # @return [StyleBuilder]
    # : (untyped name, **untyped opts) ?{ (untyped) -> untyped } -> untyped
    def style(name, **opts)
      style_builder = StyleBuilder.new(name)
      style_builder.apply_options!(**opts) unless opts.empty?
      yield style_builder if block_given?
      @styles[name] = style_builder
      style_builder
    end

    # Add a row to the sheet.
    #
    # @param values [Array, Hash] The cell values.
    # @param styles [String, Array<String>, nil] Styles to apply to cells.
    # @param height [Float, nil] The row height.
    # @param hidden [Boolean] Whether the row is hidden.
    # @param custom_height [Boolean] Whether it's a custom height.
    # @param outline_level [Integer, nil] The outline level.
    # @return [void]
    # : (untyped values, ?styles: untyped?, ?height: untyped?, ?hidden: bool, ?custom_height: bool, ?outline_level: untyped?) -> untyped
    def row(values, styles: nil, height: nil, hidden: false, custom_height: false, outline_level: nil)
      row_index = @rows.size
      # See: https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
      if @strict_excel_mode && row_index >= 1_048_576
        raise ArgumentError, "Row index #{row_index} exceeds Excel limit of 1,048,576 rows"
      end

      if values.is_a?(Hash)
        max_col = values.keys.map { |k| Elements::Cell.column_index(k) }.max || -1
        cells_array = Array.new(max_col + 1)
        values.each do |k, v|
          idx = Elements::Cell.column_index(k)
          cells_array[idx] = v
        end
        values = cells_array
      end

      if styles.is_a?(Hash)
        max_col_style = styles.keys.map { |k| Elements::Cell.column_index(k) }.max || -1
        styles_array = Array.new(max_col_style + 1)
        styles.each do |k, v|
          idx = Elements::Cell.column_index(k)
          styles_array[idx] = v
        end
        styles = styles_array
      end

      # Auto-detect Date / Time for built-in styles
      values.each_with_index do |val, idx|
        cell_style = styles.is_a?(Array) ? styles[idx] : styles

        if val.is_a?(Date) && cell_style.nil?
          style("__xlsxrb_date", number_format: "yyyy-mm-dd") unless @styles.key?("__xlsxrb_date")
          styles = [] if styles.nil?
          styles = Array.new(values.size, styles) unless styles.is_a?(Array)
          styles[idx] = "__xlsxrb_date"
        elsif val.is_a?(Time) && cell_style.nil?
          style("__xlsxrb_time", number_format: "yyyy-mm-dd hh:mm:ss") unless @styles.key?("__xlsxrb_time")
          styles = [] if styles.nil?
          styles = Array.new(values.size, styles) unless styles.is_a?(Array)
          styles[idx] = "__xlsxrb_time"
        end
      end

      max_len = values.size
      max_len = [max_len, styles.size].max if styles.is_a?(Array)
      # See: https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
      if @strict_excel_mode && max_len > 16_384
        raise ArgumentError, "Row contains #{max_len} columns, exceeding Excel limit of 16_384 columns"
      end

      cells = Array.new(max_len)
      style_lookup = styles.is_a?(Array)

      col_index = 0
      while col_index < max_len
        val = col_index < values.size ? values[col_index] : nil
        unless val.nil? || val.is_a?(String) || val.is_a?(Numeric) || val.is_a?(TrueClass) || val.is_a?(FalseClass) || val.is_a?(Date) || val.is_a?(Time) || val.is_a?(Elements::Formula) || (val.is_a?(Hash) && val.key?(:formula))
          raise ArgumentError, "Invalid cell value type: #{val.class} for value #{val.inspect}"
        end
        # See: https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
        if @strict_excel_mode && val.is_a?(String) && val.length > 32_767
          raise ArgumentError, "Cell text length #{val.length} exceeds Excel limit of 32,767 characters"
        end
        style_name = if style_lookup
                       col_index < styles.size ? styles[col_index] : nil
                     else
                       styles
                     end
        if val.nil? && style_name.nil?
          col_index += 1
          next
        end
        # If value is a Formula object or Hash with :formula, store it as the cell's formula
        cells[col_index] = if val.is_a?(Elements::Formula)
                             Elements::Cell.new(
                               row_index: row_index,
                               column_index: col_index,
                               value: nil,
                               formula: val,
                               style_index: style_name
                             )
                           elsif val.is_a?(Hash) && val.key?(:formula)
                             f_obj = Elements::Formula.new(val[:formula])
                             Elements::Cell.new(
                               row_index: row_index,
                               column_index: col_index,
                               value: val[:value],
                               formula: f_obj,
                               style_index: style_name
                             )
                           else
                             Elements::Cell.new(
                               row_index: row_index,
                               column_index: col_index,
                               value: val,
                               style_index: style_name
                             )
                           end
        col_index += 1
      end

      cells.compact!
      @rows << Elements::Row.new(
        index: row_index,
        cells: cells,
        height: height,
        hidden: hidden,
        custom_height: custom_height || !height.nil?,
        outline_level: outline_level
      )
    end

    # Add a column to the sheet.
    #
    # @param index [Integer, String] The column index (0-based) or letter.
    # @param width [Float, nil] The column width.
    # @param hidden [Boolean] Whether the column is hidden.
    # @param custom_width [Boolean] Whether it's a custom width.
    # @param outline_level [Integer, nil] The outline level.
    # @return [void]
    # : (untyped index, ?width: untyped?, ?hidden: bool, ?custom_width: bool, ?outline_level: untyped?) -> untyped
    def column(index, width: nil, hidden: false, custom_width: false, outline_level: nil)
      index = Elements::Cell.column_index(index)

      @columns << Elements::Column.new(
        index: index,
        width: width,
        hidden: hidden,
        custom_width: custom_width || !width.nil?,
        outline_level: outline_level
      )
    end

    # Add a chart to the sheet.
    #
    # @param options [Hash] Chart options.
    # @yield [builder]
    # @yieldparam builder [ChartBuilder]
    # @return [void]
    # : (**untyped options) ?{ (untyped) -> untyped } -> untyped
    def chart(**options)
      if block_given?
        builder = ChartBuilder.new
        yield builder
        options = builder.options.merge(options)
      end
      @charts << options
    end

    # --- Hyperlinks ---

    # Add a hyperlink on a cell.
    #
    # @param cell [String] The cell reference (e.g. "A1").
    # @param url [String, nil] The URL.
    # @param display [String, nil] The display text.
    # @param tooltip [String, nil] The tooltip.
    # @param location [String, nil] The internal location reference.
    # @return [void]
    # : (untyped cell, ?untyped? url, ?display: untyped?, ?tooltip: untyped?, ?location: untyped?) -> untyped
    def hyperlink(cell, url = nil, display: nil, tooltip: nil, location: nil)
      link = { cell: cell }
      link[:url] = url if url
      link[:display] = display if display
      link[:tooltip] = tooltip if tooltip
      link[:location] = location if location
      @hyperlinks << link
    end

    # --- Auto Filter / Sort ---

    # Set an auto filter range (e.g. "A1:D10").
    #
    # @param range [String] The filter range.
    # @return [void]
    # : (untyped range) -> untyped
    def auto_filter(range)
      @auto_filter = range
    end

    # Add a filter column to the auto filter.
    #
    # @param col_id [Integer] 0-based column index within the filter range.
    # @param filter [Hash] The filter options.
    # @return [void]
    # : (untyped col_id, untyped filter) -> untyped
    def filter_column(col_id, filter)
      @filter_columns[col_id] = filter
    end

    # Set sort state.
    #
    # @param ref [String] The sort range.
    # @param sort_conditions [Array<Hash>] Sort conditions.
    # @param opts [Hash] Additional options.
    # @return [void]
    # : (untyped ref, untyped sort_conditions, **untyped opts) -> untyped
    def sort_state(ref, sort_conditions, **opts)
      @sort_state = { ref: ref, sort_conditions: sort_conditions }.merge(opts)
    end

    # --- Data Validation ---

    # Add a data validation rule.
    #
    # @param sqref [String] The cell range (e.g. "A1:A100").
    # @param opts [Hash] Data validation options.
    # @return [void]
    # : (untyped sqref, **untyped opts) -> untyped
    def validate_data(sqref, **opts)
      @data_validations << opts.merge(sqref: sqref)
    end

    # --- Conditional Formatting ---

    # Add a conditional formatting rule.
    #
    # @param sqref [String] The cell range.
    # @param opts [Hash] Conditional format options.
    # @return [void]
    # : (untyped sqref, **untyped opts) -> untyped
    def conditional_format(sqref, **opts)
      @conditional_formats << opts.merge(sqref: sqref)
    end

    # --- Tables ---

    # Add a table to the sheet.
    #
    # @param ref [String] The table range.
    # @param columns [Array<String>] The column names.
    # @param name [String, nil] The table name.
    # @param display_name [String, nil] The display name.
    # @param style [String, nil] The table style.
    # @param opts [Hash] Additional options.
    # @return [void]
    # : (untyped ref, columns: untyped, ?name: untyped?, ?display_name: untyped?, ?style: untyped?, **untyped opts) -> untyped
    def table(ref, columns:, name: nil, display_name: nil, style: nil, **opts)
      tbl = { ref: ref, columns: columns }
      tbl[:name] = name if name
      tbl[:display_name] = display_name if display_name
      tbl[:style] = style if style
      tbl.merge!(opts)
      @tables << tbl
    end

    # --- Pivot Tables ---

    # Add a pivot table to the sheet.
    #
    # @param source_ref [String] data source range (e.g. "Sheet1!A1:C10").
    # @param row_fields [Array<Integer>] array of 0-based field indices for row axis.
    # @param data_fields [Array<Hash>] array of { fld:, name:, subtotal: } hashes.
    # @param col_fields [Array<Integer>] array of 0-based field indices for column axis.
    # @param dest_ref [String] top-left cell for the pivot table (default "E1").
    # @param name [String, nil] Pivot table name.
    # @param field_names [Array<String>, nil] Override field names.
    # @param items [Array, nil] Items configuration.
    # @return [void]
    # : (untyped source_ref, row_fields: untyped, data_fields: untyped, ?col_fields: untyped, ?dest_ref: ::String, ?name: untyped?, ?field_names: untyped?, ?items: untyped?) -> untyped
    def pivot_table(source_ref, row_fields:, data_fields:, col_fields: [], dest_ref: "E1", name: nil, field_names: nil, items: nil)
      @pivot_tables ||= []
      @pivot_tables << {
        source_ref: source_ref, row_fields: row_fields,
        data_fields: data_fields, col_fields: col_fields,
        dest_ref: dest_ref, name: name,
        field_names: field_names, items: items
      }
    end

    # --- Comments ---

    # Add a comment on a cell.
    #
    # @param cell [String] The cell reference.
    # @param text [String] The comment text.
    # @param author [String] The author name.
    # @return [void]
    # : (untyped cell, untyped text, ?author: ::String) -> untyped
    def comment(cell, text, author: "Author")
      @comments << { cell: cell, text: text, author: author }
    end

    # --- Sparklines ---

    # Add a sparkline group to the sheet.
    #
    # @param sparklines [Array<Hash>] Array of { data_ref:, location_ref: } hashes.
    # @param type [String, nil] "line" (default), "column", or "stacked".
    # @param opts [Hash] Additional options.
    # @return [void]
    # : (sparklines: untyped, ?type: untyped?, **untyped opts) -> untyped
    def sparkline_group(sparklines:, type: nil, **opts)
      group = { sparklines: sparklines }
      group[:type] = type if type
      group.merge!(opts)
      @sparkline_groups << group
    end

    # --- Merge Cells ---

    # Merge a range of cells (e.g. "A1:B2"), or by coordinate indices.
    #
    # @param range [String, nil] The string range.
    # @param row [Integer, nil] Single row index.
    # @param col_start [Integer, nil] Starting column index.
    # @param col_end [Integer, nil] Ending column index.
    # @param row_start [Integer, nil] Starting row index.
    # @param row_end [Integer, nil] Ending row index.
    # @return [void]
    # : (?untyped? range, ?row: untyped?, ?col_start: untyped?, ?col_end: untyped?, ?row_start: untyped?, ?row_end: untyped?) -> untyped
    def merge(range = nil, row: nil, col_start: nil, col_end: nil, row_start: nil, row_end: nil)
      if range
        if @strict_excel_mode && !range.match?(/^[A-Za-z]{1,3}\d+(:[A-Za-z]{1,3}\d+)?$/)
          raise ArgumentError, "Invalid merge range format: '#{range}'. Expected format like 'A1:B2'."
        end
        if @merge_cells_ranges.include?(range)
          return
        end
        @merge_cells_ranges << range
      else
        r_start = row || row_start || 0
        r_end = row || row_end || 0
        c_start = col_start || 0
        c_end = col_end || 0
        start_ref = "#{Xlsxrb::Elements::Cell.column_letter(c_start)}#{r_start + 1}"
        end_ref = "#{Xlsxrb::Elements::Cell.column_letter(c_end)}#{r_end + 1}"
        @merge_cells_ranges << "#{start_ref}:#{end_ref}"
      end
    end

    # --- Freeze / Split Panes ---

    # Freeze panes at the given row and column.
    #
    # @param row [Integer] The row index to freeze at (0-based).
    # @param col [Integer, String] The column index to freeze at (0-based or letter).
    # @return [void]
    # : (?row: ::Integer, ?col: ::Integer) -> untyped
    def freeze_pane(row: 0, col: 0)
      col = Elements::Cell.column_index(col)
      @freeze_pane = { row: row, col: col }
    end

    # Split panes (non-frozen).
    #
    # @param x_split [Integer] X coordinate.
    # @param y_split [Integer] Y coordinate.
    # @param top_left_cell [String, nil] Top left cell reference.
    # @return [void]
    # : (?x_split: ::Integer, ?y_split: ::Integer, ?top_left_cell: untyped?) -> untyped
    def split_pane(x_split: 0, y_split: 0, top_left_cell: nil)
      @split_pane = { x_split: x_split, y_split: y_split, top_left_cell: top_left_cell }
    end

    # Set active cell selection.
    #
    # @param active_cell [String] The active cell reference.
    # @param sqref [String, nil] The selected range.
    # @param pane [String, nil] The pane to select in.
    # @return [void]
    # : (untyped active_cell, ?sqref: untyped?, ?pane: untyped?) -> untyped
    def select_cell(active_cell, sqref: nil, pane: nil)
      @selection = { active_cell: active_cell, sqref: sqref || active_cell }
      @selection[:pane] = pane if pane
    end

    # --- Page Setup / Margins / Print ---

    # Set page margins (in inches).
    #
    # @param left [Float, nil] Left margin.
    # @param right [Float, nil] Right margin.
    # @param top [Float, nil] Top margin.
    # @param bottom [Float, nil] Bottom margin.
    # @param header [Float, nil] Header margin.
    # @param footer [Float, nil] Footer margin.
    # @return [void]
    # : (?left: untyped?, ?right: untyped?, ?top: untyped?, ?bottom: untyped?, ?header: untyped?, ?footer: untyped?) -> untyped
    def page_margins(left: nil, right: nil, top: nil, bottom: nil, header: nil, footer: nil)
      @page_margins = { left: left, right: right, top: top, bottom: bottom, header: header, footer: footer }.compact
    end

    # Set page setup properties.
    #
    # @param opts [Hash] Page setup options.
    # @return [void]
    # : (**untyped opts) -> untyped
    def page_setup(**opts)
      @page_setup.merge!(opts)
    end

    # Set header/footer text.
    #
    # @param opts [Hash] Header and footer options.
    # @return [void]
    # : (**untyped opts) -> untyped
    def header_footer(**opts)
      @header_footer.merge!(opts)
    end

    # Set a print option.
    #
    # @param name [Symbol] Option name.
    # @param value [Object] Option value.
    # @return [void]
    # : (untyped name, untyped value) -> untyped
    def print_options(name, value)
      @print_options[name] = value
    end

    # --- Sheet Protection ---

    # Set sheet protection options.
    #
    # @param opts [Hash] Sheet protection options.
    # @return [void]
    # : (**untyped opts) -> untyped
    def protect_sheet(**opts)
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
    # : (untyped file_data, ?ext: ::String, ?from_col: ::Integer, ?from_row: ::Integer, ?to_col: ::Integer, ?to_row: ::Integer, **untyped opts) -> untyped
    def image(file_data, ext: "png", from_col: 0, from_row: 0, to_col: 5, to_row: 10, **opts)
      img = { file_data: file_data, ext: ext, from_col: from_col, from_row: from_row, to_col: to_col, to_row: to_row }
      img.merge!(opts)
      @images << img
    end

    # --- Shapes ---

    # Add a shape to the sheet.
    # : (?preset: ::String, ?text: untyped?, ?from_col: ::Integer, ?from_row: ::Integer, ?to_col: ::Integer, ?to_row: ::Integer, **untyped opts) -> untyped
    def shape(preset: "rect", text: nil, from_col: 0, from_row: 0, to_col: 5, to_row: 5, **opts)
      shape = { preset: preset, text: text, from_col: from_col, from_row: from_row, to_col: to_col, to_row: to_row }
      shape[:name] = opts.delete(:name) || "Shape #{@shapes.size + 1}"
      shape.merge!(opts)
      @shapes << shape
    end

    # --- Sheet Properties ---

    # Set a sheet-level property (e.g. :tab_color).
    # : (untyped name, untyped value) -> untyped
    def sheet_properties(name, value)
      @sheet_properties[name] = value
    end

    # Set a sheet view property (e.g. :show_grid_lines, :zoom_scale).
    # : (untyped name, untyped value) -> untyped
    def sheet_view(name, value)
      @sheet_view[name] = value
    end

    # --- Row / Column Breaks ---

    # Add a page break before a row.
    # : (untyped row_num) -> untyped
    def page_break_row(row_num)
      @row_breaks << row_num
    end

    # Add a page break before a column.
    # : (untyped col_index) -> untyped
    def page_break_col(col_index)
      col_index = Elements::Cell.column_index(col_index)
      @col_breaks << col_index
    end

    # : () -> untyped
    def build
      facade_meta = {}
      facade_meta[:hyperlinks] = @hyperlinks unless @hyperlinks.empty?
      facade_meta[:auto_filter] = @auto_filter if @auto_filter
      facade_meta[:filter_columns] = @filter_columns unless @filter_columns.empty?
      facade_meta[:sort_state] = @sort_state if @sort_state
      facade_meta[:data_validations] = @data_validations unless @data_validations.empty?
      facade_meta[:conditional_formats] = @conditional_formats unless @conditional_formats.empty?
      facade_meta[:tables] = @tables unless @tables.empty?
      facade_meta[:pivot_tables] = @pivot_tables unless (@pivot_tables || []).empty?
      facade_meta[:comments] = @comments unless @comments.empty?
      facade_meta[:sparkline_groups] = @sparkline_groups unless @sparkline_groups.empty?
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
    # : untyped
    attr_reader :styles
  end

  # DSL context for Xlsxrb.generate streaming writes.
  class StreamWriter
    attr_reader :current_sheet

    # : (untyped target, ?strict_excel_mode: bool) -> void
    def initialize(target, strict_excel_mode: true)
      @target = target
      @strict_excel_mode = strict_excel_mode
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
      @workbook_properties = { update_links: "never" }
    end

    # Set a workbook property.
    #
    # @note **SECURITY WARNING:** If you set `:update_links` to anything other than `"never"`,
    #   you may expose end-users to malicious external reference vulnerabilities (e.g., CSV/DDE Injection)
    #   when they open the generated Excel file. Ensure you fully trust the exported data.
    #
    # @param name [Symbol] The property name (e.g. :update_links).
    # @param value [String, Integer, Boolean] The property value.
    # @return [void]
    # : (untyped name, untyped value) -> untyped
    def workbook_property(name, value)
      @workbook_properties[name] = value
    end

    # Define a named style that can be applied to cells.
    #
    # @param name [String] The name of the style.
    # @param opts [Hash] Style options (e.g. bold: true).
    # @yield [style_builder]
    # @yieldparam style_builder [StyleBuilder]
    # @return [StyleBuilder]
    # : (untyped name, **untyped opts) ?{ (untyped) -> untyped } -> untyped
    def style(name, **opts)
      style_builder = StyleBuilder.new(name)
      style_builder.apply_options!(**opts) unless opts.empty?
      yield style_builder if block_given?
      @styles[name] = style_builder

      # Register immediately
      @style_name_to_id[name] = style_builder.register_with(@style_writer)

      style_builder
    end

    # Proxy object yielded by the `sheet` method to prevent writing to inactive sheets.
    class WorksheetProxy
      def initialize(writer, sheet_name)
        @writer = writer
        @sheet_name = sheet_name
      end

      def respond_to_missing?(method_name, include_private = false)
        @writer.respond_to?(method_name, include_private)
      end

      def method_missing(method_name, ...)
        raise Error, "Sheet '#{@sheet_name}' is no longer active. In streaming mode, you cannot write to a previous sheet." if @writer.current_sheet != @sheet_name

        @writer.public_send(method_name, ...)
      end
    end

    # Add a new sheet.
    #
    # @param name [String, nil] The name of the sheet.
    # @param opts [Hash] Sheet properties.
    # @yield [sheet_builder]
    # @yieldparam sheet_builder [WorksheetBuilder]
    # @return [void]
    # : (?untyped? name, **untyped opts) ?{ (untyped) -> untyped } -> untyped
    def sheet(name = nil, **opts)
      name ||= "Sheet#{@sheets.size + 1}"
      if @strict_excel_mode && name.length > 31
        raise ArgumentError, "Sheet name '#{name}' must be <= 31 characters (Excel limitation)"
      end
      if name.match?(/[\[\]\*?\/\\]/)
        raise ArgumentError, "Sheet name '#{name}' contains invalid characters (ECMA-376 OOXML specification)"
      end
      if @strict_excel_mode && @sheets.map { |s| s.respond_to?(:name) ? s.name.downcase : s.to_s.downcase }.include?(name.downcase)
        raise ArgumentError, "Sheet name '#{name}' is already used. Excel requires unique sheet names."
      end

      internal_sheet_setup(name)
      opts.each { |k, v| set_sheet_property(k, v) }

      yield WorksheetProxy.new(self, @current_sheet) if block_given?
      @current_sheet
    end

    # Internal: Start or switch to a named sheet (internal helper).
    # : (?untyped? name) ?{ (untyped) -> untyped } -> (nil | untyped)
    def internal_sheet_setup(name = nil)
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
      @current_pivot_tables = []
      @current_sparkline_groups = []
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
      @current_cells = {}

      return unless block_given?

      yield self
      flush_current_sheet
    end

    # Add a row of values. values is an Array.
    # styles:: Hash mapping column indices to style names, or Array of style names for each column
    # Add a row to the sheet.
    #
    # @param values [Array, Hash] The cell values.
    # @param styles [String, Array<String>, nil] Styles to apply to cells.
    # @param height [Float, nil] The row height.
    # @param hidden [Boolean] Whether the row is hidden.
    # @param custom_height [Boolean] Whether it's a custom height.
    # @param outline_level [Integer, nil] The outline level.
    # @return [void]
    # : (untyped values, ?styles: untyped?, ?height: untyped?, ?hidden: bool, ?custom_height: bool, ?outline_level: untyped?) -> untyped
    def row(values, styles: nil, height: nil, hidden: false, custom_height: false, outline_level: nil)
      sheet if @current_sheet.nil?

      row_index = @current_row_index
      # See: https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
      if @strict_excel_mode && row_index >= 1_048_576
        raise ArgumentError, "Row index #{row_index} exceeds Excel limit of 1,048,576 rows"
      end
      @current_row_index += 1

      if values.is_a?(Hash)
        max_col = values.keys.map { |k| Elements::Cell.column_index(k) }.max || -1
        cells_array = Array.new(max_col + 1)
        values.each do |k, v|
          idx = Elements::Cell.column_index(k)
          cells_array[idx] = v
        end
        values = cells_array
      end

      if styles.is_a?(Hash)
        max_col_style = styles.keys.map { |k| Elements::Cell.column_index(k) }.max || -1
        styles_array = Array.new(max_col_style + 1)
        styles.each do |k, v|
          idx = Elements::Cell.column_index(k)
          styles_array[idx] = v
        end
        styles = styles_array
      end

      # Auto-detect Date / Time for built-in styles
      values.each_with_index do |val, idx|
        cell_style = styles.is_a?(Array) ? styles[idx] : styles

        if val.is_a?(Date) && cell_style.nil?
          style("__xlsxrb_date", number_format: "yyyy-mm-dd") unless @styles.key?("__xlsxrb_date")
          styles = [] if styles.nil?
          styles = Array.new(values.size, styles) unless styles.is_a?(Array)
          styles[idx] = "__xlsxrb_date"
        elsif val.is_a?(Time) && cell_style.nil?
          style("__xlsxrb_time", number_format: "yyyy-mm-dd hh:mm:ss") unless @styles.key?("__xlsxrb_time")
          styles = [] if styles.nil?
          styles = Array.new(values.size, styles) unless styles.is_a?(Array)
          styles[idx] = "__xlsxrb_time"
        end
      end

      @current_cells ||= {}
      row_num = row_index + 1

      max_len = values.size
      max_len = [max_len, styles.size].max if styles.is_a?(Array)
      # See: https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
      if @strict_excel_mode && max_len > 16_384
        raise ArgumentError, "Row contains #{max_len} columns, exceeding Excel limit of 16_384 columns"
      end

      max_len.times do |col_idx|
        val = col_idx < values.size ? values[col_idx] : nil
        next if val.nil?

        # See: https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
        if @strict_excel_mode && val.is_a?(String) && val.length > 32_767
          raise ArgumentError, "Cell text length #{val.length} exceeds Excel limit of 32,767 characters"
        end

        addr = "#{Elements::Cell.column_letter(col_idx)}#{row_num}"
        @current_cells[addr] = val
      end

      attrs = nil
      if height || hidden || outline_level
        attrs = {}
        attrs[:height] = height if height
        attrs[:hidden] = true if hidden
        attrs[:custom_height] = custom_height || !height.nil?
        attrs[:outline_level] = outline_level if outline_level
      end

      @current_row_writer.write_row_values(row_index, values, styles: styles, style_map: @style_name_to_id, sst: @sst, sst_index: @sst_index, attrs: attrs)
    end

    # Set column width for a 0-based column index.
    # Add a column to the sheet.
    #
    # @param index [Integer, String] The column index (0-based) or letter.
    # @param width [Float, nil] The column width.
    # @param hidden [Boolean] Whether the column is hidden.
    # @param custom_width [Boolean] Whether it's a custom width.
    # @param outline_level [Integer, nil] The outline level.
    # @return [void]
    # : (untyped index, ?width: untyped?, ?hidden: bool, ?custom_width: bool, ?outline_level: untyped?) -> untyped
    def column(index, width: nil, hidden: false, custom_width: false, outline_level: nil)
      index = Elements::Cell.column_index(index)
      sheet if @current_sheet.nil?

      @current_columns << { index: index, width: width, hidden: hidden, custom_width: custom_width || !width.nil?, outline_level: outline_level }
    end

    # Add a chart to the current sheet.
    # : (**untyped options) ?{ (untyped) -> untyped } -> untyped
    def chart(**options)
      sheet if @current_sheet.nil?

      if block_given?
        builder = ChartBuilder.new
        yield builder
        options = builder.options.merge(options)
      end

      @current_charts << options
    end

    # --- Hyperlinks ---

    # : (untyped cell, ?untyped? url, ?display: untyped?, ?tooltip: untyped?, ?location: untyped?) -> untyped
    def hyperlink(cell, url = nil, display: nil, tooltip: nil, location: nil)
      sheet if @current_sheet.nil?
      link = { cell: cell }
      link[:url] = url if url
      link[:display] = display if display
      link[:tooltip] = tooltip if tooltip
      link[:location] = location if location
      @current_hyperlinks << link
    end

    # --- Auto Filter / Sort ---

    # : (untyped range) -> untyped
    def auto_filter(range)
      sheet if @current_sheet.nil?
      @current_auto_filter = range
    end

    # : (untyped col_id, untyped filter) -> untyped
    def filter_column(col_id, filter)
      sheet if @current_sheet.nil?
      @current_filter_columns[col_id] = filter
    end

    # : (untyped ref, untyped sort_conditions, **untyped opts) -> untyped
    def sort_state(ref, sort_conditions, **opts)
      sheet if @current_sheet.nil?
      @current_sort_state = { ref: ref, sort_conditions: sort_conditions }.merge(opts)
    end

    # --- Data Validation ---

    # : (untyped sqref, **untyped opts) -> untyped
    def validate_data(sqref, **opts)
      sheet if @current_sheet.nil?
      @current_data_validations << opts.merge(sqref: sqref)
    end

    # --- Conditional Formatting ---

    # : (untyped sqref, **untyped opts) -> untyped
    def conditional_format(sqref, **opts)
      sheet if @current_sheet.nil?
      @current_conditional_formats << opts.merge(sqref: sqref)
    end

    # --- Tables ---

    # : (untyped ref, columns: untyped, ?name: untyped?, ?display_name: untyped?, ?style: untyped?, **untyped opts) -> untyped
    def table(ref, columns:, name: nil, display_name: nil, style: nil, **opts)
      sheet if @current_sheet.nil?
      tbl = { ref: ref, columns: columns }
      tbl[:name] = name if name
      tbl[:display_name] = display_name if display_name
      tbl[:style] = style if style
      tbl.merge!(opts)
      @current_tables << tbl
    end

    # --- Pivot Tables ---

    # : (untyped source_ref, row_fields: untyped, data_fields: untyped, ?col_fields: untyped, ?dest_ref: ::String, ?name: untyped?, ?field_names: untyped?, ?items: untyped?) -> untyped
    def pivot_table(source_ref, row_fields:, data_fields:, col_fields: [], dest_ref: "E1", name: nil, field_names: nil, items: nil)
      sheet if @current_sheet.nil?
      @current_pivot_tables ||= []
      @current_pivot_tables << {
        source_ref: source_ref, row_fields: row_fields,
        data_fields: data_fields, col_fields: col_fields,
        dest_ref: dest_ref, name: name,
        field_names: field_names, items: items
      }
    end

    # --- Comments ---

    # : (untyped cell, untyped text, ?author: ::String) -> untyped
    def comment(cell, text, author: "Author")
      sheet if @current_sheet.nil?
      @current_comments << { cell: cell, text: text, author: author }
    end

    # --- Sparklines ---

    # : (sparklines: untyped, ?type: untyped?, **untyped opts) -> untyped
    def sparkline_group(sparklines:, type: nil, **opts)
      sheet if @current_sheet.nil?
      group = { sparklines: sparklines }
      group[:type] = type if type
      group.merge!(opts)
      @current_sparkline_groups << group
    end

    # Merge a range of cells (e.g. "A1:B2"), or by coordinate indices.
    #
    # @param range [String, nil] The string range.
    # @param row [Integer, nil] Single row index.
    # @param col_start [Integer, nil] Starting column index.
    # @param col_end [Integer, nil] Ending column index.
    # @param row_start [Integer, nil] Starting row index.
    # @param row_end [Integer, nil] Ending row index.
    # @return [void]
    # : (?untyped? range, ?row: untyped?, ?col_start: untyped?, ?col_end: untyped?, ?row_start: untyped?, ?row_end: untyped?) -> untyped
    def merge(range = nil, row: nil, col_start: nil, col_end: nil, row_start: nil, row_end: nil)
      sheet if @current_sheet.nil?
      if range
        if @strict_excel_mode && !range.match?(/^[A-Za-z]{1,3}\d+(:[A-Za-z]{1,3}\d+)?$/)
          raise ArgumentError, "Invalid merge range format: '#{range}'. Expected format like 'A1:B2'."
        end
        @current_merge_cells << range
      else
        r_start = row || row_start || 0
        r_end = row || row_end || 0
        c_start = col_start || 0
        c_end = col_end || 0
        start_ref = "#{Xlsxrb::Elements::Cell.column_letter(c_start)}#{r_start + 1}"
        end_ref = "#{Xlsxrb::Elements::Cell.column_letter(c_end)}#{r_end + 1}"
        @current_merge_cells << "#{start_ref}:#{end_ref}"
      end
    end

    # --- Freeze / Split Panes ---

    # Freeze panes at the given row and column.
    #
    # @param row [Integer] The row index to freeze at (0-based).
    # @param col [Integer, String] The column index to freeze at (0-based or letter).
    # @return [void]
    # : (?row: ::Integer, ?col: untyped) -> untyped
    def freeze_pane(row: 0, col: 0)
      col = Elements::Cell.column_index(col)
      sheet if @current_sheet.nil?
      @current_freeze_pane = { row: row, col: col }
    end

    # : (?x_split: ::Integer, ?y_split: ::Integer, ?top_left_cell: untyped?) -> untyped
    def split_pane(x_split: 0, y_split: 0, top_left_cell: nil)
      sheet if @current_sheet.nil?
      @current_split_pane = { x_split: x_split, y_split: y_split, top_left_cell: top_left_cell }
    end

    # : (untyped active_cell, ?sqref: untyped?, ?pane: untyped?) -> untyped
    def select_cell(active_cell, sqref: nil, pane: nil)
      sheet if @current_sheet.nil?
      @current_selection = { active_cell: active_cell, sqref: sqref || active_cell }
      @current_selection[:pane] = pane if pane
    end

    # --- Page Setup / Margins / Print ---

    # : (?left: untyped?, ?right: untyped?, ?top: untyped?, ?bottom: untyped?, ?header: untyped?, ?footer: untyped?) -> untyped
    def page_margins(left: nil, right: nil, top: nil, bottom: nil, header: nil, footer: nil)
      sheet if @current_sheet.nil?
      @current_page_margins = { left: left, right: right, top: top, bottom: bottom, header: header, footer: footer }.compact
    end

    # : (**untyped opts) -> untyped
    def page_setup(**opts)
      sheet if @current_sheet.nil?
      @current_page_setup.merge!(opts)
    end

    # : (**untyped opts) -> untyped
    def header_footer(**opts)
      sheet if @current_sheet.nil?
      @current_header_footer.merge!(opts)
    end

    # : (untyped name, untyped value) -> untyped
    def print_options(name, value)
      sheet if @current_sheet.nil?
      @current_print_options[name] = value
    end

    # --- Sheet Protection ---

    # : (**untyped opts) -> untyped
    def protect_sheet(**opts)
      sheet if @current_sheet.nil?
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

    # : (untyped file_data, ?ext: ::String, ?from_col: ::Integer, ?from_row: ::Integer, ?to_col: ::Integer, ?to_row: ::Integer, **untyped opts) -> untyped
    def image(file_data, ext: "png", from_col: 0, from_row: 0, to_col: 5, to_row: 10, **opts)
      sheet if @current_sheet.nil?
      img = { file_data: file_data, ext: ext, from_col: from_col, from_row: from_row, to_col: to_col, to_row: to_row }
      img.merge!(opts)
      @current_images << img
    end

    # --- Shapes ---

    # : (?preset: ::String, ?text: untyped?, ?from_col: ::Integer, ?from_row: ::Integer, ?to_col: ::Integer, ?to_row: ::Integer, **untyped opts) -> untyped
    def shape(preset: "rect", text: nil, from_col: 0, from_row: 0, to_col: 5, to_row: 5, **opts)
      sheet if @current_sheet.nil?
      shape = { preset: preset, text: text, from_col: from_col, from_row: from_row, to_col: to_col, to_row: to_row }
      shape[:name] = opts.delete(:name) || "Shape #{@current_shapes.size + 1}"
      shape.merge!(opts)
      @current_shapes << shape
    end

    # --- Sheet Properties ---

    # : (untyped name, untyped value) -> untyped
    def sheet_properties(name, value)
      sheet if @current_sheet.nil?
      @current_sheet_properties[name] = value
    end

    # : (untyped name, untyped value) -> untyped
    def sheet_view(name, value)
      sheet if @current_sheet.nil?
      @current_sheet_view[name] = value
    end

    # --- Row / Column Breaks ---

    # : (untyped row_num) -> untyped
    def page_break_row(row_num)
      sheet if @current_sheet.nil?
      @current_row_breaks << row_num
    end

    # : (untyped col_index) -> untyped
    def page_break_col(col_index)
      col_index = Elements::Cell.column_index(col_index)
      sheet if @current_sheet.nil?
      @current_col_breaks << col_index
    end

    # --- Workbook-Level Methods ---

    # Add a defined name.
    #
    # @param name [String] The defined name.
    # @param value [String] The formula or value.
    # @param sheet [String, nil] Local sheet name.
    # @param hidden [Boolean] Whether the defined name is hidden.
    # @return [void]
    # : (untyped name, untyped value, ?sheet: untyped?, ?hidden: bool) -> untyped
    def defined_name(name, value, sheet: nil, hidden: false)
      entry = { name: name, value: value, hidden: hidden }
      if sheet
        # local_sheet_id will be resolved at close time
        entry[:local_sheet_name] = sheet
      end
      @defined_names << entry
    end

    # Set the print area for the current or named sheet.
    # : (untyped range, ?sheet: untyped?) -> untyped
    def print_area(range, sheet: nil)
      sheet_name = sheet || @current_sheet || "Sheet1"
      value = "'#{sheet_name}'!#{absolute_range(range)}"
      @defined_names.reject! { |dn| dn[:name] == "_xlnm.Print_Area" && dn[:local_sheet_name] == sheet_name }
      defined_name("_xlnm.Print_Area", value, sheet: sheet_name)
    end

    # Set print titles for the current or named sheet.
    # : (?rows: untyped?, ?cols: untyped?, ?sheet: untyped?) -> untyped
    def print_titles(rows: nil, cols: nil, sheet: nil)
      sheet_name = sheet || @current_sheet || "Sheet1"
      parts = []
      parts << "'#{sheet_name}'!$#{cols.sub(":", ":$")}" if cols
      parts << "'#{sheet_name}'!$#{rows.sub(":", ":$")}" if rows
      value = parts.join(",")
      @defined_names.reject! { |dn| dn[:name] == "_xlnm.Print_Titles" && dn[:local_sheet_name] == sheet_name }
      defined_name("_xlnm.Print_Titles", value, sheet: sheet_name)
    end

    # Set workbook protection.
    #
    # @param opts [Hash] Protection options.
    # @return [void]
    # : (**untyped opts) -> untyped
    def protect_workbook(**opts)
      @workbook_protection = opts
    end

    # Set a core document property.
    #
    # @param name [Symbol] The property name.
    # @param value [String, Integer, Time] The property value.
    # @return [void]
    # : (untyped name, untyped value) -> untyped
    def core_property(name, value)
      @core_properties[name] = value
    end

    # Set an app document property.
    #
    # @param name [Symbol] The property name.
    # @param value [String, Integer, Time] The property value.
    # @return [void]
    # : (untyped name, untyped value) -> untyped
    def app_property(name, value)
      @app_properties[name] = value
    end

    # Set multiple core and/or app properties.
    #
    # @param core [Hash, nil] Core properties.
    # @param app [Hash, nil] App properties.
    # @return [void]
    # : (?core: untyped?, ?app: untyped?) -> untyped
    def properties(core: nil, app: nil)
      core&.each { |k, v| core_property(k, v) }
      app&.each { |k, v| app_property(k, v) }
    end

    # Add a custom document property.
    #
    # @param name [String] The property name.
    # @param value [String, Integer, Float, Boolean, Time] The property value.
    # @param type [Symbol] The type of property (:string, :number, :bool, :date).
    # @return [void]
    # : (untyped name, untyped value, ?type: ::Symbol) -> untyped
    def custom_property(name, value, type: :string)
      @custom_properties << { name: name, value: value, type: type }
    end

    # : () -> untyped
    def close
      Xlsxrb.in_span("StreamWriter#close") do
        flush_current_sheet

        styles_definition = {
          fonts: @style_writer.fonts.dup,
          fills: @style_writer.fills.dup,
          borders: @style_writer.borders.dup,
          xf_entries: @style_writer.xf_entries.dup,
          num_fmts: @style_writer.num_fmts.dup
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
      cleanup!
    end

    # Explicitly remove any remaining tempfiles. Called via ensure block.
    # : () -> void
    def cleanup!
      @tempfiles.each do |tmp|
        tmp.close
        tmp.unlink
      end
      @tempfiles.clear
    end

    private

    # : (untyped range) -> untyped
    def absolute_range(range)
      range.gsub(/([A-Z]+)(\d+)/, '$\1$\2')
    end

    # : (untyped names, untyped sheets) -> untyped
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

    # : () -> (nil | untyped)
    def flush_current_sheet
      return unless @current_sheet

      @current_tempfile.close

      sheet_data = { name: @current_sheet, rows_tmp_path: @current_tempfile.path, columns: @current_columns }
      sheet_data[:cells] = @current_cells if @current_cells && !@current_cells.empty?
      @current_cells = nil
      sheet_data[:charts] = @current_charts unless @current_charts.empty?
      sheet_data[:hyperlinks] = @current_hyperlinks unless @current_hyperlinks.empty?
      sheet_data[:auto_filter] = @current_auto_filter if @current_auto_filter
      sheet_data[:filter_columns] = @current_filter_columns unless @current_filter_columns.empty?
      sheet_data[:sort_state] = @current_sort_state if @current_sort_state
      sheet_data[:data_validations] = @current_data_validations unless @current_data_validations.empty?
      sheet_data[:conditional_formats] = @current_conditional_formats unless @current_conditional_formats.empty?
      sheet_data[:tables] = @current_tables unless @current_tables.empty?
      sheet_data[:pivot_tables] = @current_pivot_tables unless @current_pivot_tables.empty?
      sheet_data[:sparkline_groups] = @current_sparkline_groups unless @current_sparkline_groups.empty?
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

        cell_errors = Elements::Cell.validate(row_idx, col_idx, rc[:value])
        if !cell_errors.empty? && rc[:source]
          cell_errors = cell_errors.map do |err|
            "#{err} (at #{rc[:source][:part]} row #{rc[:source][:row] + 1} cell #{rc[:ref] || "unknown"})"
          end
        end

        Elements::Cell.new(
          row_index: row_idx,
          column_index: col_idx,
          value: rc[:value],
          formula: rc[:formula],
          style_index: rc[:style_index],
          errors: cell_errors.empty? ? nil : cell_errors
        )
      end
      attrs = raw_row[:attrs] || {}
      row_errors = Elements::Row.validate(raw_row[:index], cells)
      if !row_errors.empty? && raw_row[:source]
        row_errors = row_errors.map do |err|
          "#{err} (at #{raw_row[:source][:part]} row #{raw_row[:source][:row] + 1})"
        end
      end
      Elements::Row.new(
        index: raw_row[:index],
        cells: cells,
        height: attrs[:height],
        hidden: attrs[:hidden] || false,
        custom_height: attrs[:custom_height] || false,
        outline_level: attrs[:outline_level],
        errors: row_errors.empty? ? nil : row_errors
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

      if cell.formula
        f = cell.formula
        if f.is_a?(Elements::Formula)
          result[:formula] = f.expression
          result[:formula_ca] = true if f.calculate_always
          if f.cached_value
            # Cached value is written as-is (not through SST)
            result[:value] = f.cached_value
            result.delete(:type) # Ensure no type is set; cached values are plain text in <v>
          end
        else
          result[:formula] = f
        end
      end
      result
    end

    def build_row_attrs(row)
      attrs = {}
      attrs[:height] = row.height if row.height
      attrs[:hidden] = true if row.hidden
      attrs[:custom_height] = true if row.custom_height
      attrs[:outline_level] = row.outline_level if row.outline_level
      attrs
    end
  end

  # Builds a raw cell hash from a value for streaming writes.
  # : (untyped row_index, untyped col_index, untyped value, untyped sst, untyped sst_index) -> untyped
  def self.build_raw_cell_from_value(row_index, col_index, value, sst, sst_index)
    ref = "#{Elements::Cell.column_letter(col_index)}#{row_index + 1}"
    result = { ref: ref }

    case value
    when Elements::Formula
      result[:formula] = value.expression
      result[:formula_ca] = true if value.calculate_always
      result[:value] = value.cached_value if value.cached_value
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
