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
    Workbook = Data.define(:sheets, :shared_strings, :styles, :unmapped_data, :errors) do
      include Enumerable

      # @param sheets [Array<Elements::Worksheet>] Worksheets in the workbook.
      # @param shared_strings [Array<String>] Shared strings table.
      # @param styles [Hash] Styles definition.
      # @param unmapped_data [Hash] Additional metadata for round-tripping.
      # @param errors [Array<String>, nil] Validation errors.
      #: (?sheets: Array[Elements::Worksheet], ?shared_strings: Array[String], ?styles: Hash[untyped, untyped], ?unmapped_data: Hash[untyped, untyped], ?errors: Array[String]?) -> void
      def initialize(sheets: [], shared_strings: [], styles: {}, unmapped_data: {}, errors: nil)
        computed_errors = errors || self.class.validate(sheets)
        super(sheets: sheets.freeze, shared_strings: shared_strings.freeze, styles: styles,
              unmapped_data: unmapped_data, errors: computed_errors.freeze)
      end

      # Iterate over worksheets.
      #
      # @example
      #   workbook.each do |sheet|
      #     puts sheet.name
      #   end
      #
      # @yield [sheet]
      # @yieldparam sheet [Elements::Worksheet]
      # @return [Enumerator, void]
      #: () { (Elements::Worksheet) -> void } -> void
      #: () -> Enumerator[Elements::Worksheet, void]
      def each(&)
        sheets.each(&)
      end
      alias_method :each_sheet, :each

      # Returns whether the workbook is valid according to ECMA-376 rules.
      #
      # @return [Boolean]
      #: () -> bool
      def valid?
        errors.empty?
      end

      # Returns the worksheet at the given 0-based index or by name.
      #
      # @example
      #   wb.sheet(0)
      #   wb.sheet("Sales")
      #
      # @param identifier [Integer, String] 0-based index or sheet name.
      # @return [Elements::Worksheet, nil]
      # @api public
      #: (?Integer | String identifier) -> Elements::Worksheet?
      def sheet(identifier = 0)
        case identifier
        when Integer
          sheets[identifier]
        when String
          sheets.find { |s| s.name == identifier }
        end
      end
      alias_method :[], :sheet

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
        with(sheets: loaded_sheets)
      end
      alias_method :to_workbook, :load

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
      #: (Integer | String identifier) { (Elements::Worksheet) -> Elements::Worksheet } -> Elements::Workbook
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
