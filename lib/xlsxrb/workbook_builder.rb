# frozen_string_literal: true

# rbs_inline: enabled

require_relative "worksheet_builder"

module Xlsxrb
  # DSL context for building an in-memory {Elements::Workbook} in {Xlsxrb.build}.
  #
  # @example Build an in-memory workbook with multiple sheets
  #   workbook = Xlsxrb.build do |builder|
  #     builder.sheet("Sales") do |sheet|
  #       sheet.row(["Product", "Revenue"])
  #       sheet.row(["Widget", 1000])
  #     end
  #     builder.core_property(:creator, "Reporting System")
  #   end
  #
  # @api public
  class WorkbookBuilder
    # @param strict_excel_mode [Boolean] Whether to enforce Microsoft Excel specifications.
    #: (?strict_excel_mode: bool) -> void
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

    # Sets a workbook-level property.
    #
    # @note **SECURITY WARNING:** If you set `:update_links` to anything other than `"never"`,
    #   you may expose end-users to malicious external reference vulnerabilities (e.g., CSV/DDE Injection)
    #   when they open the generated Excel file. Ensure you fully trust the exported data.
    #
    # @param name [Symbol] Property name (e.g. :update_links).
    # @param value [String, Integer, Boolean] Property value.
    # @return [void]
    # @api public
    #: (Symbol name, String | Integer | bool value) -> (String | Integer | bool)
    def workbook_property(name, value)
      @workbook_properties[name] = value
    end

    # Adds a new worksheet to the workbook.
    #
    # @param name [String, nil] Sheet name (max 31 chars, no forbidden chars `[ ] * ? / \`).
    # @param opts [Hash] Sheet properties.
    # @yield [sheet_builder]
    # @yieldparam sheet_builder [Xlsxrb::WorksheetBuilder]
    # @return [Elements::Worksheet]
    # @raise [ArgumentError] If sheet name violates Excel limits in strict mode.
    # @api public
    #: (?String? name, **untyped opts) ?{ (WorksheetBuilder) -> void } -> untyped
    def sheet(name = nil, **opts)
      name ||= "Sheet#{@sheets.size + 1}"
      raise ArgumentError, "Sheet name '#{name}' must be <= 31 characters (Excel limitation)" if @strict_excel_mode && name.length > 31
      raise ArgumentError, "Sheet name '#{name}' contains invalid characters (ECMA-376 OOXML specification)" if name.match?(%r{[\[\]*?/\\]})
      raise ArgumentError, "Sheet name '#{name}' is already used. Excel requires unique sheet names." if @strict_excel_mode && @sheets.map { |s| s.respond_to?(:name) ? s.name.downcase : s.to_s.downcase }.include?(name.downcase)

      sheet_builder = WorksheetBuilder.new(name, strict_excel_mode: @strict_excel_mode)
      opts.each { |k, v| sheet_builder.sheet_properties(k, v) }
      yield sheet_builder if block_given?
      @sheet_builders << sheet_builder
      @sheets << sheet_builder.build
    end
    alias [] sheet

    # --- Workbook-Level Methods ---

    # Adds a defined name (named range or formula constant) to the workbook.
    #
    # @param name [String] Defined name (e.g. "TaxRate").
    # @param value [String] Formula expression or range reference (e.g. "Sheet1!$B$2").
    # @param sheet [String, nil] Optional sheet name for sheet-scoped defined names.
    # @param hidden [Boolean] Whether the defined name is hidden from Excel's UI.
    # @return [void]
    # @api public
    #: (String name, String value, ?sheet: String?, ?hidden: bool) -> void
    def defined_name(name, value, sheet: nil, hidden: false)
      entry = { name: name, value: value, hidden: hidden }
      entry[:local_sheet_name] = sheet if sheet
      @defined_names << entry
    end

    # Sets the print area for a sheet.
    #
    # @param range [String] Cell range (e.g. "A1:G50").
    # @param sheet [String, nil] Target sheet name (defaults to latest or "Sheet1").
    # @return [void]
    # @api public
    #: (String range, ?sheet: String?) -> void
    def print_area(range, sheet: nil)
      sheet_name = sheet || @sheets.last&.name || "Sheet1"
      value = "'#{sheet_name}'!#{absolute_range(range)}"
      @defined_names.reject! { |dn| dn[:name] == "_xlnm.Print_Area" && dn[:local_sheet_name] == sheet_name }
      defined_name("_xlnm.Print_Area", value, sheet: sheet_name)
    end

    # Sets repeating print titles (rows and/or columns) for pagination.
    #
    # @param rows [String, nil] Repeating row range (e.g. "1:2").
    # @param cols [String, nil] Repeating column range (e.g. "A:B").
    # @param sheet [String, nil] Target sheet name.
    # @return [void]
    # @api public
    #: (?rows: String?, ?cols: String?, ?sheet: String?) -> void
    def print_titles(rows: nil, cols: nil, sheet: nil)
      sheet_name = sheet || @sheets.last&.name || "Sheet1"
      parts = []
      parts << "'#{sheet_name}'!$#{cols.sub(":", ":$")}" if cols
      parts << "'#{sheet_name}'!$#{rows.sub(":", ":$")}" if rows
      value = parts.join(",")
      @defined_names.reject! { |dn| dn[:name] == "_xlnm.Print_Titles" && dn[:local_sheet_name] == sheet_name }
      defined_name("_xlnm.Print_Titles", value, sheet: sheet_name)
    end

    # Sets workbook structure and window protection.
    #
    # @param opts [Hash] Protection options (e.g. lock_structure: true, password: "secret").
    # @return [void]
    # @api public
    #: (**String | Integer | bool | nil opts) -> void
    def protect_workbook(**opts)
      @workbook_protection = opts
    end

    # Sets a Dublin Core document metadata property.
    #
    # @param name [Symbol] Core property name (:creator, :title, :subject, :description, etc.).
    # @param value [String, Integer, Time] Property value.
    # @return [void]
    # @api public
    #: (Symbol name, String | Integer | Time value) -> void
    def core_property(name, value)
      @core_properties[name] = value
    end

    # Sets an extended application metadata property.
    #
    # @param name [Symbol] Application property name (:company, :manager, :app_version, etc.).
    # @param value [String, Integer, Time] Property value.
    # @return [void]
    # @api public
    #: (Symbol name, String | Integer | Time value) -> void
    def app_property(name, value)
      @app_properties[name] = value
    end

    # Sets multiple core, app, and/or custom properties simultaneously.
    #
    # @param core [Hash{Symbol => Object}, nil] Core properties map.
    # @param app [Hash{Symbol => Object}, nil] App properties map.
    # @param custom [Hash{String, Symbol => Object}, nil] Custom properties map.
    # @return [void]
    # @api public
    #: (?core: Hash[Symbol, String | Integer | Time]?, ?app: Hash[Symbol, String | Integer | Time]?, ?custom: Hash[String | Symbol, untyped]?) -> void
    def properties(core: nil, app: nil, custom: nil)
      core&.each { |k, v| core_property(k, v) }
      app&.each { |k, v| app_property(k, v) }
      custom&.each { |k, v| custom_property(k.to_s, v) }
    end

    # Adds a custom document property.
    #
    # @param name [String] Property name.
    # @param value [String, Integer, Float, Boolean, Time] Property value.
    # @param type [Symbol] Value type (:string, :number, :bool, :date).
    # @return [void]
    # @api public
    #: (String name, String | Integer | Float | bool | Time value, ?type: ::Symbol) -> void
    def custom_property(name, value, type: :string)
      @custom_properties << { name: name, value: value, type: type }
    end

    # Builds and returns the compiled {Elements::Workbook}.
    #
    # @return [Elements::Workbook]
    # @raise [ArgumentError] If workbook contains zero sheets in strict mode.
    # @api public
    #: () -> Elements::Workbook
    def build
      raise ArgumentError, "Workbook must contain at least one sheet (Excel limitation)" if @strict_excel_mode && @sheets.empty?

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

    #: (untyped range) -> untyped
    def absolute_range(range)
      range.gsub(/([A-Z]+)(\d+)/, '$\1$\2')
    end

    #: (untyped names, untyped sheets) -> untyped
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

    #: (untyped sheets) -> (::Array[untyped | ::Hash[untyped, untyped]] | ::Array[untyped])
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

    #: (untyped writer) -> { fonts: untyped, fills: untyped, borders: untyped, xf_entries: untyped, num_fmts: untyped }
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
end
