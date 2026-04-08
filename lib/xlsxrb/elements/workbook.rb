# frozen_string_literal: true

module Xlsxrb
  module Elements
    # Represents an entire XLSX workbook.
    Workbook = Data.define(:sheets, :shared_strings, :styles, :unmapped_data, :errors) do
      def initialize(sheets: [], shared_strings: [], styles: {}, unmapped_data: {}, errors: nil)
        computed_errors = errors || self.class.validate(sheets)
        super(sheets: sheets.freeze, shared_strings: shared_strings.freeze, styles: styles,
              unmapped_data: unmapped_data, errors: computed_errors.freeze)
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

      # Returns sheet names.
      def sheet_names
        sheets.map(&:name)
      end

      def self.validate(sheets)
        errs = []
        errs << "sheets must be an Array" unless sheets.is_a?(Array)
        if sheets.is_a?(Array)
          errs << "workbook must have at least one sheet" if sheets.empty?
          names = sheets.map(&:name)
          errs << "duplicate sheet names" if names.uniq.size != names.size
        end
        errs
      end
    end
  end
end
