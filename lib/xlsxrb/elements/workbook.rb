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
    class Workbook < Data.define(:sheets, :shared_strings, :styles, :unmapped_data, :errors, :defined_names)
      # rubocop:enable Style/DataInheritance
      [Enumerable].each { |m| include m }

      # @param sheets [Array<Elements::Worksheet, StreamSheet>] Worksheets in the workbook.
      # @param shared_strings [Array<String>] Shared strings table.
      # @param styles [Hash] Styles definition.
      # @param unmapped_data [Hash] Additional metadata for round-tripping.
      # @param errors [Array<String>, nil] Validation errors.
      # @param defined_names [Array<Hash>, nil] Defined names list.
      # @param date1904 [Boolean] Whether the workbook uses the 1904 date system.
      #: (?sheets: Array[Elements::Worksheet | StreamSheet], ?shared_strings: Array[String], ?styles: Hash[untyped, untyped], ?unmapped_data: Hash[untyped, untyped], ?errors: Array[String]?, ?defined_names: Array[Hash[Symbol, untyped]]?, ?date1904: bool) -> void
      def initialize(sheets: [], shared_strings: [], styles: {}, unmapped_data: {}, errors: nil, defined_names: nil, date1904: false)
        dns = defined_names || unmapped_data[:defined_names] || unmapped_data.dig(:facade, :defined_names) || []
        computed_errors = errors || self.class.validate(sheets)
        unmapped = unmapped_data
        if date1904 && !unmapped.dig(:workbook_properties, :date1904) && !unmapped.dig(:facade, :workbook_properties, :date1904)
          unmapped = unmapped.dup
          unmapped[:workbook_properties] = (unmapped[:workbook_properties] || {}).merge(date1904: true)
        end
        super(sheets: sheets.freeze, shared_strings: shared_strings.freeze, styles: styles,
              unmapped_data: unmapped, errors: computed_errors.freeze, defined_names: dns.freeze)
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
      alias each_sheet each

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
        props = unmapped_data[:workbook_properties] || unmapped_data.dig(:facade, :workbook_properties)
        props&.[](:date1904) ? true : false
      end

      # Returns the defined name matching name (and optionally sheet), or nil.
      #
      # @example
      #   wb.defined_name("SalesTotal")
      #   wb.defined_name("Print_Area", sheet: "Sheet1")
      #
      # @param name [String] The defined name.
      # @param sheet [String, Integer, nil] Optional sheet name or 0-based index.
      # @return [Hash[Symbol, untyped], nil]
      # @api public
      #: (String name, ?sheet: (String | Integer)?) -> Hash[Symbol, untyped]?
      def defined_name(name, sheet: nil)
        sheet_idx = case sheet
                    when Integer then sheet
                    when String then sheets.find_index { |s| s.name == sheet }
                    end

        if sheet_idx
          defined_names.find { |dn| dn[:name] == name && dn[:local_sheet_id] == sheet_idx }
        else
          defined_names.find { |dn| dn[:name] == name && dn[:local_sheet_id].nil? } ||
            defined_names.find { |dn| dn[:name] == name }
        end
      end

      # Returns the worksheet at the given 0-based index or by name.
      #
      # @example
      #   wb.sheet(0)
      #   wb.sheet("Sales")
      #
      # @param identifier [Integer, String, untyped] 0-based index or sheet name.
      # @return [Elements::Worksheet, StreamSheet, nil]
      # @api public
      #: (?Integer | String | untyped identifier) -> (Elements::Worksheet | StreamSheet)?
      def sheet(identifier = 0)
        case identifier
        when Integer
          sheets[identifier]
        when String
          sheets.find { |s| s.name == identifier }
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

      # Loads all sheets into memory, returning an Elements::Workbook where every
      # worksheet is a fully-parsed Elements::Worksheet supporting coordinate random access.
      #
      # @example
      #   wb = Xlsxrb.read("file.xlsx").load
      #   puts wb["Sheet1"]["A1"].value
      #
      # @return [Elements::Workbook]
      # @api public
      #: () -> Elements::Workbook
      def load
        loaded_sheets = sheets.map { |s| s.respond_to?(:load) ? s.load : s }
        close
        with(sheets: loaded_sheets)
      end
      alias to_workbook load

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
      # @return [void]
      # @api public
      #: (String | IO | StringIO filepath) -> void
      def save(filepath)
        Xlsxrb.write(filepath, self)
      end

      # Validates workbook structure according to OOXML specifications.
      #
      # @param sheets [Array<Elements::Worksheet>]
      # @return [Array<String>] List of error messages.
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
