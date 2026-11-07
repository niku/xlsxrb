# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Elements
    # Represents an entire in-memory XLSX workbook.
    #
    # @example Access sheets
    #   workbook = Xlsxrb.read("report.xlsx")
    #   sheet = workbook.sheet(0) # or workbook["Sheet1"]
    #   workbook.each { |s| puts s.name }
    #
    # @api public
    # rubocop:disable Style/DataInheritance -- Required as class syntax for mutant subject matcher
    class Workbook < Data.define(:sheets, :shared_strings, :styles, :unmapped_data, :errors, :defined_names, :workbook_properties, :core_properties, :app_properties, :workbook_protection, :custom_properties)
      # rubocop:enable Style/DataInheritance
      [Enumerable].each { |m| include m }

      # @param sheets [Array<Elements::Worksheet, StreamSheet>] Worksheets in the workbook.
      # @param shared_strings [Array<String>] Shared strings table.
      # @param styles [Elements::Styles, Hash] Styles definition.
      # @param unmapped_data [Hash] Additional metadata for round-tripping.
      # @param errors [Array<String>, nil] Validation errors.
      # @param defined_names [Array<Hash>, nil] Defined names list.
      # @param workbook_properties [Elements::WorkbookProperties, Hash, nil] Workbook properties.
      # @param core_properties [Elements::CoreProperties, Hash, nil] Core properties.
      # @param app_properties [Elements::AppProperties, Hash, nil] App properties.
      # @param workbook_protection [Elements::WorkbookProtection, Hash, nil] Workbook protection.
      # @param custom_properties [Hash, Array, nil] Custom properties.
      # @param date1904 [Boolean] Whether the workbook uses the 1904 date system.
      #: (?sheets: Array[Elements::Worksheet | StreamSheet], ?shared_strings: Array[String], ?styles: Elements::Styles | Hash[untyped, untyped], ?unmapped_data: Hash[untyped, untyped], ?errors: Array[String]?, ?defined_names: Array[untyped]?, ?workbook_properties: (Elements::WorkbookProperties | Hash[Symbol, untyped])?, ?core_properties: (Elements::CoreProperties | Hash[Symbol, untyped])?, ?app_properties: (Elements::AppProperties | Hash[Symbol, untyped])?, ?workbook_protection: (Elements::WorkbookProtection | Hash[Symbol, untyped])?, ?custom_properties: (Hash[String, Elements::CustomProperty] | Hash[untyped, untyped] | Array[untyped])?, ?date1904: bool) -> void
      def initialize(sheets: [], shared_strings: [], styles: {}, unmapped_data: {}, errors: nil, defined_names: nil,
                     workbook_properties: nil, core_properties: nil, app_properties: nil, workbook_protection: nil,
                     custom_properties: nil, date1904: false)
        dns = (defined_names || unmapped_data[:defined_names] || unmapped_data.dig(:facade, :defined_names) || []).map do |dn|
          dn.is_a?(Elements::DefinedName) ? dn : Elements::DefinedName.new(**dn)
        end
        computed_errors = errors || self.class.validate(sheets)
        unmapped = unmapped_data

        raw_wb_props = workbook_properties || unmapped[:workbook_properties] || unmapped.dig(:facade, :workbook_properties) || {}
        norm_wb_props = raw_wb_props.is_a?(Elements::WorkbookProperties) ? raw_wb_props : Elements::WorkbookProperties.new(**raw_wb_props)
        if date1904 && !norm_wb_props.date1904?
          norm_wb_props = norm_wb_props.with(date1904: true)
          unmapped = unmapped.dup
          unmapped[:workbook_properties] = (unmapped[:workbook_properties] || {}).merge(date1904: true)
        end

        raw_core = core_properties || unmapped[:core_properties] || unmapped.dig(:facade, :core_properties) || {}
        norm_core = raw_core.is_a?(Elements::CoreProperties) ? raw_core : Elements::CoreProperties.new(**raw_core)

        raw_app = app_properties || unmapped[:app_properties] || unmapped.dig(:facade, :app_properties) || {}
        norm_app = raw_app.is_a?(Elements::AppProperties) ? raw_app : Elements::AppProperties.new(**raw_app)

        raw_prot = workbook_protection || unmapped[:workbook_protection] || unmapped.dig(:facade, :workbook_protection)
        norm_prot = if raw_prot.nil?
                      nil
                    elsif raw_prot.is_a?(Elements::WorkbookProtection)
                      raw_prot
                    else
                      Elements::WorkbookProtection.new(**raw_prot)
                    end

        raw_custom = custom_properties || unmapped[:custom_properties] || unmapped.dig(:facade, :custom_properties)
        norm_custom = if raw_custom.nil?
                        nil
                      elsif raw_custom.is_a?(Hash)
                        raw_custom.map do |k, v|
                          if v.is_a?(Elements::CustomProperty)
                            v
                          elsif v.is_a?(Hash)
                            Elements::CustomProperty.new(name: (v[:name] || v["name"] || k).to_s, value: v[:value] || v["value"], type: v[:type] || v["type"])
                          else
                            Elements::CustomProperty.new(name: k.to_s, value: v)
                          end
                        end
                      elsif raw_custom.is_a?(Array)
                        raw_custom.map do |v|
                          if v.is_a?(Hash)
                            Elements::CustomProperty.new(name: (v[:name] || v["name"]).to_s, value: v[:value] || v["value"], type: v[:type] || v["type"])
                          else
                            v
                          end
                        end
                      else
                        []
                      end.to_h { |cp| [cp.name.to_s, cp] }.freeze

        norm_styles = styles.is_a?(Elements::Styles) ? styles : Elements::Styles.new(styles || {})

        super(sheets: sheets.freeze, shared_strings: shared_strings.freeze, styles: norm_styles,
              unmapped_data: unmapped, errors: computed_errors.freeze, defined_names: dns.freeze,
              workbook_properties: norm_wb_props, core_properties: norm_core, app_properties: norm_app,
              workbook_protection: norm_prot, custom_properties: norm_custom)
      end

      # Iterate over worksheets.
      #
      # @example
      #   workbook.each do |sheet|
      #     puts sheet.name
      #   end
      #
      # @yield [sheet]
      # @yieldparam sheet [Elements::Worksheet, StreamSheet]
      # @return [Enumerator, void]
      #: () { (Elements::Worksheet | StreamSheet) -> void } -> void
      #: () -> Enumerator[Elements::Worksheet | StreamSheet, void]
      def each(&)
        sheets.each(&)
      end

      # @deprecated Use {#each} instead. Will be removed in 0.2.0.
      #: () { (Elements::Worksheet | StreamSheet) -> void } -> void
      #: () -> Enumerator[Elements::Worksheet | StreamSheet, void]
      def each_sheet(&)
        warn "[DEPRECATION] #each_sheet is deprecated; use #each instead.", uplevel: 1
        each(&)
      end

      # Returns whether the workbook is valid according to ECMA-376 rules.
      #
      # @return [Boolean]
      #: () -> bool
      def valid?
        errors.empty?
      end

      # Returns whether the workbook uses the 1904 date system.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def date1904?
        workbook_properties.date1904?
      end

      # Returns the internal VBA or application code name of the workbook, or nil if not specified.
      #
      # @return [String, nil]
      # @api public
      #: () -> String?
      def code_name
        workbook_properties.code_name
      end

      # Returns the 0-based active worksheet index from workbookView, or nil if not set.
      #
      # @return [Integer, nil]
      # @api public
      #: () -> Integer?
      def active_sheet_index
        workbook_properties.active_sheet
      end

      # @deprecated Use {#active_sheet_index} or {#active_sheet} instead.
      # @return [Integer, nil]
      # @api public
      #: () -> Integer?
      def active_tab
        warn "[Xlsxrb DEPRECATION] Workbook#active_tab is deprecated and will be removed in v0.2.0. Use #active_sheet_index or #active_sheet instead.", uplevel: 1
        active_sheet_index
      end

      # Returns the active worksheet in the workbook, or the first worksheet if not set.
      #
      # @return [Elements::Worksheet, StreamSheet, nil]
      # @api public
      #: () -> (Elements::Worksheet | StreamSheet)?
      def active_sheet
        idx = active_sheet_index
        idx ? sheets[idx] : sheets.first
      end

      # Returns the index of the first sheet tab visible in the sheet tab scroll area.
      #
      # @return [Integer, nil]
      # @api public
      #: () -> Integer?
      def first_sheet
        workbook_properties.first_sheet
      end

      # Returns whether sheet tabs are displayed in the workbook window.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def show_sheet_tabs?
        workbook_properties.show_sheet_tabs != false
      end

      # @deprecated Use {#show_sheet_tabs?} instead. Will be removed in 0.2.0.
      #: () -> bool
      def show_sheet_tabs
        warn "[DEPRECATION] #show_sheet_tabs is deprecated; use #show_sheet_tabs? instead.", uplevel: 1
        show_sheet_tabs?
      end

      # Returns the ratio between the sheet tabs bar and the horizontal scroll bar (0-1000).
      #
      # @return [Integer, nil]
      # @api public
      #: () -> Integer?
      def tab_ratio
        workbook_properties.tab_ratio
      end

      # Returns whether Excel is instructed to perform a full recalculation when loading the workbook.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def full_calc_on_load?
        workbook_properties.full_calc_on_load ? true : false
      end

      # @deprecated Use {#full_calc_on_load?} instead. Will be removed in 0.2.0.
      #: () -> bool
      def full_calc_on_load
        warn "[DEPRECATION] #full_calc_on_load is deprecated; use #full_calc_on_load? instead.", uplevel: 1
        full_calc_on_load?
      end

      # Returns the defined name matching name (and optionally sheet), or nil.
      #
      # @example
      #   wb.defined_name("SalesTotal")
      #   wb.defined_name("Print_Area", sheet: "Sheet1")
      #
      # @param name [String] The defined name.
      # @param sheet [String, Symbol, Integer, nil] Optional sheet name or 0-based index.
      # @return [Elements::DefinedName, nil]
      # @api public
      #: (String name, ?sheet: (String | Symbol | Integer)?) -> Elements::DefinedName?
      def defined_name(name, sheet: nil)
        sheet_idx = case sheet
                    when Integer then sheet
                    when String, Symbol
                      target = sheet.to_s
                      sheets.find_index { |s| s.name == target }
                    end

        dn_name = ->(dn) { dn.respond_to?(:name) ? dn.name : dn[:name] }
        dn_local = ->(dn) { dn.respond_to?(:local_sheet_id) ? dn.local_sheet_id : dn[:local_sheet_id] }

        if sheet_idx
          defined_names.find { |dn| dn_name.call(dn) == name && dn_local.call(dn) == sheet_idx }
        else
          defined_names.find { |dn| dn_name.call(dn) == name && dn_local.call(dn).nil? } ||
            defined_names.find { |dn| dn_name.call(dn) == name }
        end
      end

      # Returns the worksheet at the given 0-based index or by name.
      #
      # @example
      #   wb.sheet(0)
      #   wb.sheet("Sales")
      #   wb.sheet(:Sales)
      #
      # @param identifier [Integer, String, Symbol, untyped] 0-based index or sheet name.
      # @return [Elements::Worksheet, StreamSheet, nil]
      # @api public
      #: (?Integer | String | Symbol | untyped identifier) -> (Elements::Worksheet | StreamSheet)?
      def sheet(identifier = 0)
        case identifier
        when Integer
          sheets[identifier]
        when String, Symbol
          target = identifier.to_s
          sheets.find { |s| s.name == target }
        end
      end
      alias [] sheet

      # Returns the formatted string representation of a cell's value in the specified sheet.
      #
      # @param sheet_identifier [Integer, String, untyped] 0-based index or sheet name.
      # @param ref_or_row [String, Symbol, Integer] Cell reference (e.g. "A1") or 0-based row index.
      # @param col [Integer, nil] Optional 0-based column index.
      # @return [String, nil]
      # @api public
      #: (Integer | String | untyped sheet_identifier, String | Symbol | Integer ref_or_row, ?Integer? col) -> String?
      def formatted_value(sheet_identifier, ref_or_row, col = nil)
        target_sheet = sheet(sheet_identifier)
        target_sheet&.formatted_value(ref_or_row, col)
      end

      # Returns all embedded images across all worksheets.
      #
      # @return [Array<Elements::Image>]
      # @api public
      #: () -> Array[Elements::Image]
      def images
        sheets.flat_map(&:images)
      end

      # Returns all images anchored at a specific cell in the given worksheet.
      #
      # @param sheet_identifier [Integer, String, untyped] Sheet index or sheet name.
      # @param ref_or_row [String, Symbol, Integer] Cell reference (e.g. "A1") or 0-based row index.
      # @param col [Integer, nil] Optional 0-based column index.
      # @return [Array<Elements::Image>]
      # @api public
      #: (Integer | String | untyped sheet_identifier, String | Symbol | Integer ref_or_row, ?Integer? col) -> Array[Elements::Image]
      def images_at(sheet_identifier, ref_or_row, col = nil)
        ident = sheet_identifier.is_a?(Symbol) ? sheet_identifier.to_s : sheet_identifier
        target_sheet = sheet(ident)
        return [] unless target_sheet

        target_sheet.images_at(ref_or_row, col)
      end

      # Internal load helper without deprecation warning.
      # @api private
      #: () -> Elements::Workbook
      def _load
        loaded_sheets = sheets.map do |s|
          s.respond_to?(:load) ? s.load : s
        end
        close
        with(sheets: loaded_sheets)
      end

      # Loads all sheets into memory, returning an Elements::Workbook where every
      # worksheet is a fully-parsed Elements::Worksheet supporting coordinate random access.
      #
      # @example
      #   wb = Xlsxrb.load("file.xlsx")
      #   puts wb["Sheet1"]["A1"].value
      #
      # @deprecated Use {Xlsxrb.load} instead. Will be removed in 0.2.0.
      # @return [Elements::Workbook]
      # @api public
      #: () -> Elements::Workbook
      def load
        warn "[DEPRECATION] Elements::Workbook#load is deprecated; use Xlsxrb.load(source) instead.", uplevel: 1
        _load
      end

      # @deprecated Use {Xlsxrb.load} instead. Will be removed in 0.2.0.
      #: () -> Elements::Workbook
      def to_workbook
        warn "[DEPRECATION] #to_workbook is deprecated; use Xlsxrb.load(source) instead.", uplevel: 1
        _load
      end

      # Closes any streaming resources associated with worksheets.
      #
      # @return [void]
      # @api public
      #: () -> void
      def close
        sheets.each { |s| s.close if s.respond_to?(:close) }
        nil
      end

      # Returns a new Workbook with the specified sheet updated.
      # Yields the matched worksheet to the block, which must return a new Worksheet.
      #
      # @deprecated Use {Xlsxrb.read} + {Xlsxrb.write} for streaming transformations or {Xlsxrb.write} DSL for report generation. Will be removed in 0.2.0.
      # @example
      #   new_wb = wb.update_sheet("Sheet1") do |sheet|
      #     sheet.update_cell("A1", value: "New Title")
      #   end
      #
      # @param identifier [Integer, String] 0-based index or sheet name.
      # @yield [sheet]
      # @yieldparam sheet [Elements::Worksheet] The worksheet to update.
      # @yieldreturn [Elements::Worksheet] The modified worksheet.
      # @return [Elements::Workbook] A new Workbook instance.
      # @api public
      #: (Integer | String identifier) ?{ (Elements::Worksheet) -> Elements::Worksheet } -> Elements::Workbook
      def update_sheet(identifier)
        warn "[DEPRECATION] Elements::Workbook#update_sheet is deprecated and will be removed in 0.2.0; use Xlsxrb.read + Xlsxrb.write for streaming transformations or Xlsxrb.write DSL for report generation.", uplevel: 1
        raise ArgumentError, "block is required" unless block_given?

        sheet_to_update = sheet(identifier)
        raise ArgumentError, "sheet not found: #{identifier}" unless sheet_to_update

        sheet_to_update = sheet_to_update.load if sheet_to_update.respond_to?(:load)
        new_sheet = yield sheet_to_update
        raise TypeError, "block must return a Worksheet" unless new_sheet.is_a?(Worksheet)

        new_sheets = sheets.map { |s| s.name == sheet_to_update.name ? new_sheet : s }
        with(sheets: new_sheets)
      end

      # Returns an Array of all worksheet names.
      #
      # @return [Array<String>]
      # @api public
      #: () -> Array[String]
      def sheet_names
        sheets.map(&:name)
      end

      # Save the workbook to an XLSX file.
      #
      # @example
      #   wb.save("output.xlsx")
      #
      # @param filepath [String, IO, StringIO] Destination file path or IO stream.
      # @param shared_strings [Boolean] Whether to intern strings into a shared strings table.
      # @param compression_level [Integer, nil] ZIP deflate compression level (0..9).
      # @param options [Hash] Optional write options forwarded to {Xlsxrb.write} (e.g. `password:`, `encryption_mode:`).
      # @return [void]
      # @api public
      #: (String | IO | StringIO filepath, ?shared_strings: bool, ?compression_level: Integer?, **untyped options) -> void
      def save(filepath, shared_strings: true, compression_level: nil, **options)
        Xlsxrb.write(filepath, self, shared_strings: shared_strings, compression_level: compression_level, **options)
      end

      # Serializes the workbook to a binary XLSX data string in memory.
      #
      # @example Export workbook bytes
      #   data = wb.to_data
      #
      # @param shared_strings [Boolean] Whether to intern strings into a shared strings table.
      # @param compression_level [Integer, nil] ZIP deflate compression level (0..9).
      # @param options [Hash] Additional write options forwarded to {Xlsxrb.write} (e.g. `password:`, `encryption_mode:`).
      # @return [String] Binary XLSX data.
      # @api public
      #: (?shared_strings: bool, ?compression_level: Integer?, **untyped options) -> String
      def to_data(shared_strings: true, compression_level: nil, **options)
        Xlsxrb.write(self, shared_strings: shared_strings, compression_level: compression_level, **options)
      end

      # Validates workbook structure according to OOXML specifications.
      #
      # @param sheets [Array<Elements::Worksheet>]
      # @return [Array<String>] List of error messages.
      # @api private
      #: (untyped sheets) -> Array[String]
      def self.validate(sheets)
        errs = []
        errs << "sheets must be an Array (got #{sheets.class})" unless sheets.is_a?(Array)
        if sheets.is_a?(Array)
          errs << "workbook must have at least one sheet" if sheets.empty?
          names = sheets.map(&:name)
          if names.uniq.size != names.size
            dups = names.select { |n| names.count(n) > 1 }.uniq
            errs << "duplicate sheet name: #{dups.map(&:inspect).join(", ")} — sheet names must be unique"
          end
        end
        errs
      end
    end
  end
end
