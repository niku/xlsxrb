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
    OMITTED = Object.new.freeze
    private_constant :OMITTED

    # @param strict_excel_mode [Boolean] Whether to enforce Microsoft Excel specifications.
    # @param escape_formulas [Boolean] Whether to treat strings starting with formula prefixes as plain text.
    # @param default_font [Hash, nil] Optional default font options (e.g. `{ name: "Arial", size: 12 }`).
    #: (?strict_excel_mode: bool, ?escape_formulas: bool, ?default_font: untyped?) -> void
    def initialize(strict_excel_mode: true, escape_formulas: false, default_font: nil)
      @strict_excel_mode = strict_excel_mode
      @escape_formulas = escape_formulas ? true : false
      @default_font = default_font
      @sheets = []
      @sheet_builders = [] # Keep track of sheet builders for style processing
      @defined_names = []
      @styles = {}
      @core_properties = {}
      @app_properties = {}
      @custom_properties = []
      @workbook_protection = nil
      @workbook_properties = { update_links: "never" }
      @active_sheet = nil
      @first_sheet = nil
      @show_sheet_tabs = nil
      @tab_ratio = nil
    end

    # @return [Boolean] Whether to treat strings starting with formula prefixes as plain text.
    #: bool
    attr_accessor :escape_formulas

    # Returns whether formula escaping is enabled.
    #
    # @return [Boolean]
    # @api public
    #: () -> bool
    def escape_formulas?
      @escape_formulas ? true : false
    end

    # Sets or configures the workbook-wide default font.
    #
    # @param name [String, Hash, nil] Font name or font options Hash.
    # @param size [Float, Integer, nil] Font size in points.
    # @param opts [Hash] Additional font options (e.g. `bold:`, `italic:`, `family:`, `charset:`, `color:`).
    # @return [void]
    # @api public
    #: (?String | Hash[Symbol, untyped]? name, ?Float | Integer | nil size, **untyped opts) -> void
    def default_font(name = nil, size = nil, **opts)
      existing = @default_font.is_a?(Hash) ? @default_font.dup : {}
      if name.is_a?(Hash)
        opts = name.merge(opts)
        name = opts[:name]
      end
      existing[:name] = name if name
      existing[:size] = size if size
      @default_font = existing.merge(opts)
    end

    # @return [Hash{String => StyleBuilder}]
    #: Hash[String, StyleBuilder]
    attr_reader :styles

    # Defines or configures a named cell style at the workbook level.
    #
    # @example
    #   workbook.style(:header, bold: true, fill_color: "4F81BD", font_color: "FFFFFF")
    #
    # @param name [String, Symbol] The name of the style.
    # @param opts [Hash] Style options (e.g. bold: true, fill_color: "FF0000").
    # @yield [style_builder]
    # @yieldparam style_builder [Xlsxrb::StyleBuilder]
    # @return [StyleBuilder]
    # @api public
    #: (String | Symbol name, **untyped opts) ?{ (StyleBuilder) -> void } -> StyleBuilder
    def style(name, **opts)
      style_name = name.to_s
      style_builder = StyleBuilder.new(style_name)
      style_builder.apply_options!(**opts) unless opts.empty?
      yield style_builder if block_given?
      @styles[style_name] = style_builder
    end

    # Creates an {Elements::Formula} object.
    #
    # @param expression [String] The formula expression without leading '='.
    # @param cached_value [Object, nil] Optional precomputed value.
    # @param type [Symbol, String, nil] Formula type (:shared, :array, etc.).
    # @param ref [String, nil] Target cell or range reference.
    # @param shared_index [Integer, nil] Shared formula group index (0-based).
    # @return [Elements::Formula]
    # @api public
    #: (String expression, ?cached_value: (String | Numeric | bool | Elements::CellError | nil), ?type: (Symbol | String)?, ?ref: String?, ?shared_index: Integer?) -> Elements::Formula
    def formula(...)
      Xlsxrb.formula(...)
    end

    # Creates an array or dynamic array {Elements::Formula} object.
    #
    # @param expression [String] The formula expression without leading '='.
    # @param ref [String, nil] Target cell or range reference (keyword).
    # @param cached_value [Object, nil] Optional precomputed value.
    # @return [Elements::Formula]
    # @api public
    #: (String expression, ?ref: String?, ?cached_value: (String | Numeric | bool | Elements::CellError | nil), **untyped attrs) -> Elements::Formula
    def array_formula(...)
      Xlsxrb.array_formula(...)
    end

    # Creates a shared formula {Elements::Formula} object.
    #
    # @param expression [String, nil] The formula expression.
    # @param shared_index [Integer, nil] Shared formula group index (0-based).
    # @param ref [String, nil] Target range for master cell.
    # @param cached_value [Object, nil] Optional precomputed value.
    # @return [Elements::Formula]
    # @api public
    #: (?String? expression, ?shared_index: Integer?, ?ref: String?, ?cached_value: (String | Numeric | bool | Elements::CellError | nil)) -> Elements::Formula
    def shared_formula(...)
      Xlsxrb.shared_formula(...)
    end

    # Creates an {Elements::RichText} object.
    #
    # @param runs [Array] Text runs or styling options.
    # @return [Elements::RichText]
    # @api public
    #: (*(Hash[Symbol, untyped] | Elements::RichTextRun | String | Array[untyped] | untyped) runs, ?text: String?, ?base_font: untyped, **untyped font_props) -> Elements::RichText
    def rich_text(...)
      Xlsxrb.rich_text(...)
    end

    # Creates or returns an {Elements::CellError} object.
    #
    # @param code [String, Symbol, Elements::CellError] Error code or symbol.
    # @return [Elements::CellError]
    # @api public
    #: (String | Symbol | Elements::CellError code) -> Elements::CellError
    def cell_error(...)
      Xlsxrb.cell_error(...)
    end

    # Gets or sets the active worksheet tab in the workbook window.
    #
    # @return [String, Symbol, Integer, nil]
    # @api public
    #: (String | Symbol | Integer | nil)
    attr_accessor :active_sheet

    # Gets or sets the first sheet tab visible in the sheet tab scroll area.
    #
    # @return [String, Symbol, Integer, nil]
    # @api public
    #: (String | Symbol | Integer | nil)
    attr_accessor :first_sheet

    # Gets or sets whether sheet tabs are visible in the workbook window.
    #
    # @return [Boolean, nil]
    # @api public
    #: bool?
    attr_accessor :show_sheet_tabs

    # Gets or sets the ratio between the sheet tabs bar and the horizontal scroll bar (0-1000).
    #
    # @return [Integer, nil]
    # @api public
    #: Integer?
    attr_accessor :tab_ratio

    # Gets the internal code name of the workbook.
    #
    # @return [String, nil]
    # @api public
    #: () -> String?
    def code_name
      @workbook_properties[:code_name]
    end

    # Sets the internal code name of the workbook.
    #
    # @param name [String, Symbol, nil] The code name.
    # @return [String, nil]
    # @api public
    #: (String | Symbol | nil name) -> String?
    def code_name=(name)
      @workbook_properties[:code_name] = name&.to_s
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
    #: (Symbol name, String | Integer | bool value) -> void
    def workbook_property(name, value)
      @workbook_properties[name] = value
      nil
    end

    # Returns workbook-level properties.
    #
    # @return [Elements::WorkbookProperties]
    # @api public
    #: () -> Elements::WorkbookProperties
    def workbook_properties
      Elements::WorkbookProperties.new(**@workbook_properties)
    end

    # Sets workbook-level properties.
    #
    # @param props [Hash, Elements::MetadataRecord]
    # @return [void]
    # @api public
    #: (untyped props) -> void
    def workbook_properties=(props)
      @workbook_properties = props.is_a?(Elements::MetadataRecord) ? props.to_h : props
    end

    # Returns the list of worksheet names added to this workbook so far.
    #
    # @return [Array<String>] Worksheet names.
    # @api public
    #: () -> Array[String]
    def sheet_names
      (@sheets.map { |s| s.respond_to?(:name) ? s.name : s.to_s } + @sheet_builders.map(&:name)).uniq
    end

    # Adds a new worksheet to the workbook.
    #
    # @param name [String, nil] Sheet name (max 31 chars, no forbidden chars `[ ] * ? / \ :`).
    # @param state [Symbol, String, nil] Sheet visibility state (:visible, :hidden, :very_hidden).
    # @param tab_color [String, Hash, Symbol, nil] Tab color hex string (e.g. "FF0000") or hash options.
    # @param code_name [String, Symbol, nil] Internal VBA code name of the sheet.
    # @param opts [Hash] Additional sheet properties.
    # @yield [sheet_builder]
    # @yieldparam sheet_builder [Xlsxrb::WorksheetBuilder]
    # @return [Elements::Worksheet]
    # @raise [ArgumentError] If sheet name violates Excel limits in strict mode.
    # @api public
    #: (?String? name, ?state: Symbol | String | nil, ?tab_color: String | Hash[Symbol, untyped] | Symbol | nil, ?code_name: String | Symbol | nil, **untyped opts) ?{ (WorksheetBuilder) -> void } -> Elements::Worksheet
    def sheet(name = nil, state: nil, tab_color: nil, code_name: nil, **opts)
      sheet_count = [@sheets.size, @sheet_builders.size].max
      name ||= "Sheet#{sheet_count + 1}"
      raise ArgumentError, "Sheet name '#{name}' must be <= 31 characters (Excel limitation)" if @strict_excel_mode && name.length > 31
      raise ArgumentError, "Sheet name '#{name}' contains invalid characters (Excel limitation)" if name.match?(Elements::Worksheet.forbidden_name_chars)
      raise ArgumentError, "Sheet name '#{name}' cannot start or end with an apostrophe (Excel limitation)" if @strict_excel_mode && (name.start_with?("'") || name.end_with?("'"))
      raise ArgumentError, "Sheet name '#{name}' is reserved by Excel (Excel limitation)" if @strict_excel_mode && name.casecmp?("History")

      existing_names = @sheets.map { |s| s.respond_to?(:name) ? s.name.downcase : s.to_s.downcase } + @sheet_builders.map { |sb| sb.name.downcase }
      raise ArgumentError, "Sheet name '#{name}' is already used. Excel requires unique sheet names." if @strict_excel_mode && existing_names.include?(name.downcase)

      active = opts.delete(:active)
      state = opts.delete(:state) if opts.key?(:state)
      hidden = opts.delete(:hidden)
      code_name = opts.delete(:code_name) if opts.key?(:code_name)
      tab_color = opts.delete(:tab_color) if opts.key?(:tab_color)
      state ||= :hidden if hidden
      state ||= :visible
      state = state.to_sym
      raise ArgumentError, "Invalid sheet state ':#{state}'. Allowed states are :visible, :hidden, :very_hidden" unless VALID_SHEET_STATES.include?(state)

      @active_sheet = name if active

      sheet_builder = WorksheetBuilder.new(name, strict_excel_mode: @strict_excel_mode, state: state, escape_formulas: @escape_formulas)
      sheet_builder.code_name = code_name if code_name
      sheet_builder.tab_color = tab_color if tab_color
      opts.each { |k, v| sheet_builder.sheet_properties(k, v) }
      @sheet_builders << sheet_builder
      yield sheet_builder if block_given?
      @defined_names.concat(sheet_builder.defined_names)
      sheet_builder.defined_names.clear
      built = sheet_builder.build
      @sheets << built
      built
    end

    # Accesses or defines a worksheet by name.
    #
    # @param name_or_index [String, Symbol, Integer, nil]
    # @yield [sheet_builder]
    # @return [Elements::Worksheet]
    # @api public
    #: (?String | Symbol | Integer | nil name_or_index) ?{ (WorksheetBuilder) -> void } -> Elements::Worksheet
    def [](name_or_index = nil, &)
      sheet(name_or_index&.to_s, &)
    end

    # --- Workbook-Level Methods ---

    # Adds a defined name (named range or formula) to the workbook.
    # Bare cell references and coordinate keywords are automatically formatted with sheet quotes and absolute $ references.
    #
    # @param name [String] Defined name identifier.
    # @param value [Object, nil] Value or formula expression.
    # @param sheet [String, nil] Reference sheet name for coordinates/formula.
    # @param scope [String, Symbol, nil] Optional sheet name or :workbook for scope (defaults to workbook).
    # @param hidden [Boolean] Whether the defined name is hidden from Excel's UI.
    # @param columns [Integer, String, nil] Column index or letter for coordinate-based reference.
    # @param rows [Integer, Range, nil] Row or row range (0-based) for coordinate-based reference.
    # @param range [String, nil] Cell range reference.
    # @return [void]
    # @api public
    #: (String name, ?untyped value, ?sheet: String?, ?scope: (String | Symbol)?, ?hidden: bool, ?columns: untyped, ?rows: untyped, ?range: untyped) -> void
    def defined_name(name, value = nil, sheet: nil, scope: nil, hidden: false, columns: nil, rows: nil, range: nil)
      normalized_val = DslHelpers.normalize_defined_name_value(value, sheet: sheet, columns: columns, rows: rows, range: range)
      entry = { name: name, value: normalized_val, hidden: hidden }
      local_sheet = if scope
                      scope == :workbook ? nil : scope.to_s
                    end
      entry[:local_sheet_name] = local_sheet if local_sheet
      @defined_names << entry
    end

    # Returns workbook-level defined names.
    #
    # @return [Array<Elements::DefinedName>]
    # @api public
    #: () -> Array[Elements::DefinedName]
    def defined_names
      @defined_names.map { |dn| dn.is_a?(Elements::DefinedName) ? dn : Elements::DefinedName.new(**dn) }
    end

    # Sets or gets the print area for a sheet.
    #
    # @param range [String, Object, nil] Cell range (e.g. "A1:G50").
    # @param sheet [String, nil] Target sheet name (defaults to latest or "Sheet1").
    # @return [String, nil, void]
    # @api public
    #: (?untyped range, ?sheet: String?) -> (String | nil | void)
    def print_area(range = OMITTED, sheet: nil)
      sheet_name = sheet || @sheets.last&.name || "Sheet1"
      if range.equal?(OMITTED)
        dn = @defined_names.find { |d| d[:name] == "_xlnm.Print_Area" && d[:local_sheet_name] == sheet_name }
        return dn ? dn[:value].split("!", 2).last.tr("$", "") : nil
      end

      @defined_names.reject! { |dn| dn[:name] == "_xlnm.Print_Area" && dn[:local_sheet_name] == sheet_name }
      if range
        value = "'#{sheet_name}'!#{absolute_range(range)}"
        defined_name("_xlnm.Print_Area", value, sheet: sheet_name, scope: sheet_name)
      end
      nil
    end

    # Sets the print area for the worksheet.
    #
    # @param range [String, nil] Cell range (e.g. "A1:G50").
    # @return [String, nil]
    # @api public
    #: (String? range) -> String?
    def print_area=(range)
      print_area(range) if range
      range
    end

    # Sets repeating print titles (rows and/or columns) for pagination.
    #
    # @param rows [String, nil] Repeating row range (e.g. "1:2").
    # @param columns [String, nil] Repeating column range (e.g. "A:B").
    # @param sheet [String, nil] Target sheet name.
    # @return [void]
    # @api public
    #: (?rows: (String | Integer | Range[Integer])?, ?columns: (String | Integer | Symbol | Range[untyped])?, ?sheet: String?) -> void
    def print_titles(rows: nil, columns: nil, sheet: nil)
      c = DslHelpers.normalize_print_titles_columns(columns)
      r = DslHelpers.normalize_print_titles_rows(rows)
      sheet_name = sheet || @sheets.last&.name || "Sheet1"
      existing = @defined_names.find { |dn| dn[:name] == "_xlnm.Print_Titles" && dn[:local_sheet_name] == sheet_name }
      if existing && (r.nil? || c.nil?)
        c = Regexp.last_match(1).delete("$") if c.nil? && existing[:value] =~ /'[^']*'!\$([A-Za-z]+:\$[A-Za-z]+)/
        r = Regexp.last_match(1).delete("$") if r.nil? && existing[:value] =~ /'[^']*'!\$([0-9]+:\$[0-9]+)/
      end
      parts = []
      parts << "'#{sheet_name}'!$#{c.sub(":", ":$")}" if c
      parts << "'#{sheet_name}'!$#{r.sub(":", ":$")}" if r
      value = parts.join(",")
      @defined_names.reject! { |dn| dn[:name] == "_xlnm.Print_Titles" && dn[:local_sheet_name] == sheet_name }
      defined_name("_xlnm.Print_Titles", value, sheet: sheet_name, scope: sheet_name)
      nil
    end

    # Sets repeating print titles.
    #
    # @param opts [Elements::PrintTitles, Hash] Options with :rows and/or :columns.
    # @return [Elements::PrintTitles, Hash]
    # @api public
    #: (Elements::PrintTitles | Hash[Symbol, untyped] opts) -> (Elements::PrintTitles | Hash[Symbol, untyped])
    def print_titles=(opts)
      h = opts.respond_to?(:to_h) ? opts.to_h : opts
      print_titles(**h)
      opts
    end

    # Sets repeating title rows for pagination on the target or active sheet.
    #
    # @example Repeat first row
    #   wb.repeat_rows(0)
    #
    # @param first [Integer, String] First row index (0-based) or range string (e.g. "1:2").
    # @param last [Integer, String, nil] Optional last row index (0-based).
    # @param sheet [String, nil] Target sheet name.
    # @return [void]
    # @api public
    #: (Integer | String first, ?(Integer | String)? last, ?sheet: String?) -> void
    def repeat_rows(first, last = nil, sheet: nil)
      rows_str = if first.is_a?(String)
                   first.include?(":") ? first : "#{first}:#{last || first}"
                 else
                   last ||= first
                   "#{first + 1}:#{last + 1}"
                 end
      print_titles(rows: rows_str, sheet: sheet)
    end

    # Sets repeating title columns for pagination on the target or active sheet.
    #
    # @example Repeat columns A through C
    #   wb.repeat_columns("A", "C")
    #
    # @param first [Integer, String] First column index (0-based) or column name/range.
    # @param last [Integer, String, nil] Optional last column index/name.
    # @param sheet [String, nil] Target sheet name.
    # @return [void]
    # @api public
    #: (Integer | String first, ?(Integer | String)? last, ?sheet: String?) -> void
    def repeat_columns(first, last = nil, sheet: nil)
      cols_str = if first.is_a?(String)
                   first.include?(":") ? first : "#{first}:#{last || first}"
                 else
                   first_name = Elements::Cell._column_letter(first)
                   last_name = if last.is_a?(Integer)
                                 Elements::Cell._column_letter(last)
                               elsif last
                                 last.to_s
                               else
                                 first_name
                               end
                   "#{first_name}:#{last_name}"
                 end
      print_titles(columns: cols_str, sheet: sheet)
    end

    # Sets or returns workbook structure and window protection.
    #
    # @param opts [Hash] Protection options (e.g. lock_structure: true, password: "secret").
    # @return [Elements::WorkbookProtection, nil, void]
    # @api public
    #: (**String | Integer | bool | nil opts) -> (Elements::WorkbookProtection | nil)
    def workbook_protection(**opts)
      if opts.empty?
        @workbook_protection ? Elements::WorkbookProtection.new(**@workbook_protection) : nil
      else
        @workbook_protection = opts
        nil
      end
    end

    # Sets workbook protection settings.
    #
    # @param opts [Hash, Elements::MetadataRecord, nil]
    # @return [void]
    # @api public
    #: (untyped opts) -> void
    def workbook_protection=(opts)
      @workbook_protection = opts.is_a?(Elements::MetadataRecord) ? opts.to_h : opts
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

    # Returns core document properties.
    #
    # @return [Elements::CoreProperties]
    # @api public
    #: () -> Elements::CoreProperties
    def core_properties
      Elements::CoreProperties.new(**@core_properties)
    end

    # Sets an extended application metadata property.
    #
    # @param name [Symbol] Application property name (:company, :manager, :app_version, etc.).
    # @param value [String, Integer, Time] Property value.
    # @return [void]
    # @api public
    #: (Symbol name, String | Integer | Time value) -> void
    def app_property(name, value)
      norm_name = name.to_sym == :application ? :app_name : name.to_sym
      @app_properties[norm_name] = value
    end

    # Returns application document properties.
    #
    # @return [Elements::AppProperties]
    # @api public
    #: () -> Elements::AppProperties
    def app_properties
      Elements::AppProperties.new(**@app_properties)
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

    # Returns custom document properties.
    #
    # @return [Hash<String, Elements::CustomProperty>]
    # @api public
    #: () -> Hash[String, Elements::CustomProperty]
    def custom_properties
      @custom_properties.to_h do |cp|
        prop = cp.is_a?(Elements::CustomProperty) ? cp : Elements::CustomProperty.new(name: (cp[:name] || cp["name"]).to_s, value: cp[:value] || cp["value"], type: cp[:type] || cp["type"] || :string)
        [prop.name.to_s, prop]
      end.freeze
    end

    # Builds and returns the compiled {Elements::Workbook}.
    #
    # @return [Elements::Workbook]
    # @raise [ArgumentError] If workbook contains zero sheets in strict mode.
    # @api public
    #: () -> Elements::Workbook
    def build
      if @strict_excel_mode
        raise ArgumentError, "Workbook must contain at least one sheet (Excel limitation)" if @sheets.empty?
        raise ArgumentError, "Workbook must contain at least one visible sheet (Excel limitation)" if @sheets.none? { |s| s.state == :visible }
      end

      # Process styles from all sheets and collect style definitions
      processed_sheets, styles_definition = process_styles(@sheets)

      # Auto-register print_area and print_titles defined names from sheet builders if not explicitly defined
      @sheet_builders.each do |sb|
        print_area(sb.print_area, sheet: sb.name) if sb.print_area && @defined_names.none? { |dn| dn[:name] == "_xlnm.Print_Area" && dn[:local_sheet_name] == sb.name }
        print_titles(rows: sb.print_titles.rows, columns: sb.print_titles.columns, sheet: sb.name) if sb.print_titles && @defined_names.none? { |dn| dn[:name] == "_xlnm.Print_Titles" && dn[:local_sheet_name] == sb.name }
        sb.defined_names&.each do |dn|
          @defined_names << dn
        end
      end

      active_tab_idx = if @active_sheet.is_a?(Integer)
                         @active_sheet
                       elsif @active_sheet
                         idx = processed_sheets.find_index { |s| s.name == @active_sheet.to_s }
                         raise ArgumentError, "Active sheet '#{@active_sheet}' not found" if idx.nil? && @strict_excel_mode

                         idx || @active_sheet
                       end

      if active_tab_idx.nil?
        first_vis_idx = processed_sheets.find_index { |s| s.state == :visible }
        active_tab_idx = first_vis_idx if first_vis_idx&.positive?
      end

      if active_tab_idx.is_a?(Integer) && @strict_excel_mode
        target = processed_sheets[active_tab_idx]
        raise ArgumentError, "Active sheet index #{active_tab_idx} out of bounds" if target.nil?
        raise ArgumentError, "Cannot set hidden sheet '#{target.name}' as active sheet" if target.state != :visible
      end

      first_sheet_idx = if @first_sheet.is_a?(Integer)
                          @first_sheet
                        elsif @first_sheet
                          idx = processed_sheets.find_index { |s| s.name == @first_sheet.to_s }
                          raise ArgumentError, "First sheet '#{@first_sheet}' not found" if idx.nil? && @strict_excel_mode

                          idx || @first_sheet
                        end

      @workbook_properties[:active_sheet] = active_tab_idx unless active_tab_idx.nil?
      @workbook_properties[:active_tab] = active_tab_idx unless active_tab_idx.nil?
      @workbook_properties[:first_sheet] = first_sheet_idx unless first_sheet_idx.nil?
      @workbook_properties[:show_sheet_tabs] = @show_sheet_tabs unless @show_sheet_tabs.nil?
      @workbook_properties[:tab_ratio] = @tab_ratio unless @tab_ratio.nil?

      # Store workbook-level metadata in unmapped_data
      wb_meta = {}
      dns = resolve_defined_names(@defined_names, processed_sheets)
      wb_meta[:defined_names] = dns unless dns.empty?
      wb_meta[:core_properties] = @core_properties unless @core_properties.empty?
      wb_meta[:app_properties] = @app_properties unless @app_properties.empty?
      wb_meta[:custom_properties] = @custom_properties unless @custom_properties.empty?
      wb_meta[:workbook_protection] = @workbook_protection if @workbook_protection
      wb_meta[:workbook_properties] = @workbook_properties unless @workbook_properties.empty?

      Elements::Workbook.new(
        sheets: processed_sheets,
        styles: styles_definition,
        unmapped_data: wb_meta.empty? ? {} : { facade: wb_meta },
        defined_names: dns
      )
    end

    private

    #: (untyped range) -> untyped
    def absolute_range(range)
      DslHelpers.absolute_range(range)
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
      # Collect all unique StyleBuilders from workbook and sheets
      all_style_builders = @styles.dup
      @sheet_builders.each do |sheet_builder|
        sheet_builder.styles.each do |style_name, style_builder|
          all_style_builders[style_name.to_s] = style_builder
        end
      end

      # Also collect inline styles from columns and rows
      sheets.each do |sheet|
        sheet.columns.each do |col|
          raw_col_style = col.style_index
          next unless raw_col_style.is_a?(Hash)

          inline_name = "__inline_col_#{raw_col_style.hash}"
          next if all_style_builders.key?(inline_name)

          sb = StyleBuilder.new(inline_name)
          sb.apply_options!(**raw_col_style)
          all_style_builders[inline_name] = sb
        end

        sheet.rows.each do |row|
          raw_row_style = row.unmapped_data&.[](:style_index)
          next unless raw_row_style.is_a?(Hash)

          inline_name = "__inline_row_#{raw_row_style.hash}"
          next if all_style_builders.key?(inline_name)

          sb = StyleBuilder.new(inline_name)
          sb.apply_options!(**raw_row_style)
          all_style_builders[inline_name] = sb
        end
      end

      if all_style_builders.empty?
        if @default_font
          temp_writer = Ooxml::Writer.new(default_font: @default_font)
          return [sheets, extract_styles_from_writer(temp_writer)]
        end
        return [sheets, {}]
      end

      # Create a temporary writer to register styles and get numeric IDs
      temp_writer = Ooxml::Writer.new(default_font: @default_font)
      style_name_to_id = {}
      all_style_builders.each do |name, builder|
        style_id = builder.register_with(temp_writer)
        style_name_to_id[name.to_s] = style_id
      end

      # Capture the style definitions from the temporary writer
      styles_definition = extract_styles_from_writer(temp_writer)

      # Update all sheets with resolved style IDs and style inheritance
      updated_sheets = sheets.map do |sheet|
        col_style_map = {}
        col_fmt_map = {}

        new_columns = sheet.columns.map do |col|
          raw_col_style = col.style_index
          resolved_col_style = if raw_col_style.is_a?(Hash)
                                 inline_name = "__inline_col_#{raw_col_style.hash}"
                                 sb = all_style_builders[inline_name]
                                 if sb&.number_format
                                   col_fmt_map[col.index] = sb.number_format.is_a?(String) ? sb.number_format : Ooxml::Utils::BUILTIN_NUM_FMT_CODES[sb.number_format]
                                 end
                                 style_name_to_id[inline_name]
                               elsif (raw_col_style.is_a?(String) || raw_col_style.is_a?(Symbol)) && style_name_to_id.key?(raw_col_style.to_s)
                                 sb = all_style_builders[raw_col_style.to_s]
                                 if sb&.number_format
                                   col_fmt_map[col.index] = sb.number_format.is_a?(String) ? sb.number_format : Ooxml::Utils::BUILTIN_NUM_FMT_CODES[sb.number_format]
                                 end
                                 style_name_to_id[raw_col_style.to_s]
                               else
                                 raw_col_style
                               end

          col_style_map[col.index] = resolved_col_style if resolved_col_style
          if resolved_col_style == col.style_index
            col
          else
            unmapped = (col.unmapped_data || {}).dup
            unmapped[:style_name] = raw_col_style
            unmapped[:style_index] = resolved_col_style if resolved_col_style
            Elements::Column.new(
              index: col.index,
              width: col.width,
              hidden: col.hidden?,
              custom_width: col.custom_width?,
              best_fit: col.best_fit?,
              collapsed: col.collapsed?,
              outline_level: col.outline_level,
              style_index: resolved_col_style,
              unmapped_data: unmapped,
              errors: col.errors
            )
          end
        end

        new_rows = sheet.rows.map do |row|
          raw_row_style = row.unmapped_data&.[](:style_index)
          resolved_row_style = if raw_row_style.is_a?(Hash)
                                 inline_name = "__inline_row_#{raw_row_style.hash}"
                                 style_name_to_id[inline_name]
                               elsif (raw_row_style.is_a?(String) || raw_row_style.is_a?(Symbol)) && style_name_to_id.key?(raw_row_style.to_s)
                                 style_name_to_id[raw_row_style.to_s]
                               elsif raw_row_style.is_a?(Integer)
                                 raw_row_style
                               end
          row_unmapped = row.unmapped_data
          if resolved_row_style && resolved_row_style != raw_row_style
            row_unmapped = (row.unmapped_data || {}).dup
            row_unmapped[:style_index] = resolved_row_style
          end

          new_cells = row.cells.map do |cell|
            raw_cell_style = cell.style || cell.style_index
            resolved_cell_style = if (raw_cell_style.is_a?(String) || raw_cell_style.is_a?(Symbol)) && style_name_to_id.key?(raw_cell_style.to_s)
                                    style_name_to_id[raw_cell_style.to_s]
                                  elsif raw_cell_style.nil?
                                    resolved_row_style || col_style_map[cell.column_index]
                                  else
                                    raw_cell_style
                                  end

            target_fmt = cell.format_code || col_fmt_map[cell.column_index]

            if resolved_cell_style == (cell.style || cell.style_index) && target_fmt == cell.format_code
              cell
            else
              Elements::Cell.new(
                row_index: cell.row_index,
                column_index: cell.column_index,
                value: cell.value,
                formula: cell.formula,
                style: (resolved_cell_style unless resolved_cell_style.is_a?(Integer)),
                style_index: (resolved_cell_style if resolved_cell_style.is_a?(Integer)),
                hyperlink: cell.hyperlink,
                comment: cell.comment,
                raw_value: cell.raw_value,
                format_code: target_fmt,
                unmapped_data: cell.unmapped_data,
                errors: cell.errors
              )
            end
          end

          Elements::Row.new(
            index: row.index,
            cells: new_cells,
            height: row.height,
            hidden: row.hidden?,
            custom_height: row.custom_height?,
            collapsed: row.collapsed?,
            outline_level: row.outline_level,
            style_index: resolved_row_style,
            unmapped_data: row_unmapped,
            errors: row.errors
          )
        end

        sheet.with(columns: new_columns, rows: new_rows)
      end

      [updated_sheets, styles_definition]
    end

    #: (untyped writer) -> Hash[Symbol, untyped]
    def extract_styles_from_writer(writer)
      # Extract style definitions from the writer that can be reused
      # This captures the fonts, fills, borders, and xf entries that were created
      {
        fonts: writer.fonts.dup,
        fills: writer.fills.dup,
        borders: writer.borders.dup,
        cell_xfs: writer.xf_entries.dup,
        xf_entries: writer.xf_entries.dup,
        num_fmts: writer.num_fmts.dup
      }
    end
  end
end
