# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Elements
    # Represents an entire XLSX workbook.
    Workbook = Data.define(:sheets, :shared_strings, :styles, :unmapped_data, :errors) do
      include Enumerable

      def initialize(sheets: [], shared_strings: [], styles: {}, unmapped_data: {}, errors: nil)
        computed_errors = errors || self.class.validate(sheets)
        super(sheets: sheets.freeze, shared_strings: shared_strings.freeze, styles: styles,
              unmapped_data: unmapped_data, errors: computed_errors.freeze)
      end

      def each(&)
        sheets.each(&)
      end

      def valid?
        errors.empty?
      end

      # Returns the sheet at the given 0-based index or by name.
      def sheet(identifier = 0)
        case identifier
        when Integer
          sheets[identifier]
        when String
          sheets.find { |s| s.name == identifier }
        end
      end
      alias_method :[], :sheet

      # Returns a new Workbook with the specified sheet updated.
      # If a block is given, it yields the matched sheet and expects a new Worksheet back.
      def update_sheet(identifier)
        raise ArgumentError, "block is required" unless block_given?

        sheet_to_update = sheet(identifier)
        raise ArgumentError, "sheet not found: #{identifier}" unless sheet_to_update

        new_sheet = yield sheet_to_update
        raise TypeError, "block must return a Worksheet" unless new_sheet.is_a?(Worksheet)

        new_sheets = sheets.map { |s| s == sheet_to_update ? new_sheet : s }
        with(sheets: new_sheets)
      end

      # Returns sheet names.
      def sheet_names
        sheets.map(&:name)
      end

      # Save the workbook to a file.
      def save(filepath)
        Xlsxrb.write(filepath, self)
      end

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
