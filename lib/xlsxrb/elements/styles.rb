# frozen_string_literal: true

# rbs_inline: enabled

require_relative "../number_formatter"
require_relative "../ooxml/utils"

module Xlsxrb
  module Elements
    # Represents the workbook stylesheet definitions (fonts, fills, borders, cellXfs, and numFmts).
    # Inherits from Hash for backward compatibility with existing styles hash structures,
    # while providing precomputed O(1) number format classifications.
    #
    # @api public
    class Styles < Hash #[untyped, untyped]
      # Initializes the styles container, optionally wrapping an existing styles hash.
      #
      # @param initial_hash [Hash, nil] Initial styles dictionary.
      #: (?Hash[untyped, untyped]? initial_hash) -> void
      def initialize(initial_hash = nil)
        super()
        merge!(initial_hash) if initial_hash.is_a?(Hash)
        @format_codes = nil
        @format_types = nil
        @date_flags = nil
        @datetime_flags = nil
        @time_flags = nil
      end

      # Returns the format code string for a given style index.
      #
      # @param style_index [Integer, nil]
      # @return [String, nil]
      # @api public
      #: (Integer? style_index) -> String?
      def number_format(style_index)
        return nil unless style_index.is_a?(Integer) && style_index >= 0

        precompute! unless @format_codes
        @format_codes[style_index]
      end
      alias format_code number_format
      alias format_code_for number_format

      # Returns the classified format category (:date, :datetime, :time, :number, :text, :general) for a style index.
      #
      # @param style_index [Integer, nil]
      # @return [Symbol, nil]
      # @api public
      #: (Integer? style_index) -> Symbol?
      def format_type(style_index)
        return nil unless style_index.is_a?(Integer) && style_index >= 0

        precompute! unless @format_types
        @format_types[style_index]
      end
      alias classification_for format_type

      # Returns whether the style index represents a date, datetime, or time pattern.
      #
      # @param style_index [Integer, nil]
      # @return [Boolean]
      # @api public
      #: (Integer? style_index) -> bool
      def date_format?(style_index)
        return false unless style_index.is_a?(Integer) && style_index >= 0

        precompute! unless @date_flags
        @date_flags[style_index] || false
      end

      # Returns whether the style index represents a datetime pattern.
      #
      # @param style_index [Integer, nil]
      # @return [Boolean]
      # @api public
      #: (Integer? style_index) -> bool
      def datetime_format?(style_index)
        return false unless style_index.is_a?(Integer) && style_index >= 0

        precompute! unless @datetime_flags
        @datetime_flags[style_index] || false
      end

      # Returns whether the style index represents a time pattern.
      #
      # @param style_index [Integer, nil]
      # @return [Boolean]
      # @api public
      #: (Integer? style_index) -> bool
      def time_format?(style_index)
        return false unless style_index.is_a?(Integer) && style_index >= 0

        precompute! unless @time_flags
        @time_flags[style_index] || false
      end

      # Returns whether the style index represents a date-only pattern.
      #
      # @param style_index [Integer, nil]
      # @return [Boolean]
      # @api public
      #: (Integer? style_index) -> bool
      def date_only_format?(style_index)
        format_type(style_index) == :date
      end

      # Precomputes format code and type classification lookups for all cellXfs.
      #
      # @return [self]
      # @api public
      #: () -> self
      def precompute!
        cell_xfs = self[:cell_xfs]
        size = cell_xfs.is_a?(Array) ? cell_xfs.size : 0

        codes = Array.new(size)
        types = Array.new(size)
        d_flags = Array.new(size, false)
        dt_flags = Array.new(size, false)
        t_flags = Array.new(size, false)

        if size.positive?
          cell_style_xfs = self[:cell_style_xfs]
          num_fmts = self[:num_fmts] || {}

          size.times do |i|
            xf = cell_xfs[i]
            next unless xf.is_a?(Hash)

            num_fmt_id = xf[:num_fmt_id]
            if (num_fmt_id.nil? || num_fmt_id.zero?) && xf[:xf_id] && cell_style_xfs.is_a?(Array)
              parent_xf = cell_style_xfs[xf[:xf_id]]
              num_fmt_id = parent_xf[:num_fmt_id] if parent_xf.is_a?(Hash)
            end

            code = num_fmt_id ? (num_fmts[num_fmt_id] || Ooxml::Utils::BUILTIN_NUM_FMT_CODES[num_fmt_id]) : nil

            codes[i] = code
            type = NumberFormatter.format_type(code)
            types[i] = type
            d_flags[i] = %i[date datetime time].include?(type)
            dt_flags[i] = (type == :datetime)
            t_flags[i] = (type == :time)
          end
        end

        @format_codes = codes
        @format_types = types
        @date_flags = d_flags
        @datetime_flags = dt_flags
        @time_flags = t_flags
        self
      end
    end
  end
end
