# frozen_string_literal: true

module Xlsxrb
  module Elements
    # Represents column formatting in a worksheet.
    # index is 0-based.
    Column = Data.define(:index, :width, :hidden, :custom_width, :outline_level, :unmapped_data, :errors) do
      def initialize(index:, width: nil, hidden: false, custom_width: false, outline_level: nil,
                     unmapped_data: {}, errors: nil)
        computed_errors = errors || self.class.validate(index)
        super(index: index, width: width, hidden: hidden, custom_width: custom_width,
              outline_level: outline_level, unmapped_data: unmapped_data,
              errors: computed_errors.freeze)
      end

      def valid?
        errors.empty?
      end

      def self.validate(index)
        errs = []
        errs << "index must be >= 0" if !index.is_a?(Integer) || index.negative?
        errs << "index must be < 16384" if index.is_a?(Integer) && index >= 16_384
        errs
      end
    end
  end
end
