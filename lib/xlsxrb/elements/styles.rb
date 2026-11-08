# frozen_string_literal: true

# rbs_inline: enabled

require_relative "../number_formatter"
require_relative "../ooxml/utils"

module Xlsxrb
  module Elements
    # Represents the workbook stylesheet definitions (fonts, fills, borders, cellXfs, numFmts, etc.).
    # Encapsulates stylesheet collections, item lookups, and precomputed O(1) number format classifications.
    #
    # @api public
    class Styles
      EMPTY_ARRAY = [].freeze
      EMPTY_HASH = {}.freeze
      private_constant :EMPTY_ARRAY, :EMPTY_HASH

      # Initializes the styles container, optionally wrapping an existing styles dictionary.
      #
      # @param initial_hash [Hash, nil] Initial styles dictionary.
      #: (?Hash[untyped, untyped]? initial_hash) -> void
      def initialize(initial_hash = nil)
        @data = initial_hash.is_a?(Hash) ? initial_hash.dup : {}
        invalidate_cache!
      end

      # Returns array of font entries from the styles.
      #
      # @return [Array<Elements::Font>]
      # @api public
      #: () -> Array[Elements::Font]
      def fonts
        @fonts ||= (@data[:fonts] || EMPTY_ARRAY).map do |f|
          if f.is_a?(Elements::Font)
            f
          else
            f = f.dup
            f[:size] ||= f[:sz] if f.key?(:sz)
            Elements::Font.new(**f)
          end
        end.freeze
      end

      # Returns array of fill entries from the styles.
      #
      # @return [Array<Elements::Fill>]
      # @api public
      #: () -> Array[Elements::Fill]
      def fills
        @fills ||= (@data[:fills] || EMPTY_ARRAY).map do |f|
          if f.is_a?(Elements::Fill)
            f
          else
            f = f.dup
            f[:pattern_type] ||= f[:pattern] if f.key?(:pattern)
            Elements::Fill.new(**f)
          end
        end.freeze
      end

      # Returns array of border entries from the styles.
      #
      # @return [Array<Elements::Border>]
      # @api public
      #: () -> Array[Elements::Border]
      def borders
        @borders ||= (@data[:borders] || EMPTY_ARRAY).map { |b| b.is_a?(Elements::Border) ? b : Elements::Border.new(**b) }.freeze
      end

      # Returns array of cell format (cellXfs) entries from the styles.
      #
      # @return [Array<Elements::CellXf>]
      # @api public
      #: () -> Array[Elements::CellXf]
      def cell_xfs
        @cell_xfs ||= (@data[:cell_xfs] || EMPTY_ARRAY).map { |x| x.is_a?(Elements::CellXf) ? x : Elements::CellXf.new(**x) }.freeze
      end

      # Returns custom number formats as { numFmtId => formatCode }.
      #
      # @return [Hash<Integer, String>]
      # @api public
      #: () -> Hash[Integer, String]
      def number_formats
        raw = @data[:num_fmts]
        if raw.is_a?(Array)
          raw.to_h { |nf| [nf[:num_fmt_id] || nf["num_fmt_id"], nf[:format_code] || nf["format_code"]] }
        else
          raw || EMPTY_HASH
        end
      end

      # Returns array of differential format (dxf) entries from the styles.
      #
      # @return [Array<Elements::Dxf>]
      # @api public
      #: () -> Array[Elements::Dxf]
      def dxfs
        @dxfs ||= (@data[:dxfs] || EMPTY_ARRAY).map { |d| d.is_a?(Elements::Dxf) ? d : Elements::Dxf.new(**d) }.freeze
      end

      # Returns array of cell style format (cellStyleXfs) entries from the styles.
      #
      # @return [Array<Elements::CellXf>]
      # @api public
      #: () -> Array[Elements::CellXf]
      def cell_style_xfs
        @cell_style_xfs ||= (@data[:cell_style_xfs] || EMPTY_ARRAY).map { |x| x.is_a?(Elements::CellXf) ? x : Elements::CellXf.new(**x) }.freeze
      end

      # Returns array of named cell styles from the styles.
      #
      # @return [Array<Elements::CellStyle>]
      # @api public
      #: () -> Array[Elements::CellStyle]
      def cell_styles
        @cell_styles ||= (@data[:cell_styles] || EMPTY_ARRAY).map { |s| s.is_a?(Elements::CellStyle) ? s : Elements::CellStyle.new(**s) }.freeze
      end

      # Returns table styles configuration dictionary from the styles.
      #
      # @return [Elements::TableStyles]
      # @api public
      #: () -> Elements::TableStyles
      def table_styles
        @table_styles ||= begin
          raw = @data[:table_styles] || EMPTY_HASH
          raw.is_a?(Elements::TableStyles) ? raw : Elements::TableStyles.new(**raw)
        end
      end

      # Returns array of indexed color strings (ARGB hex).
      #
      # @return [Array<String>]
      # @api public
      #: () -> Array[String]
      def indexed_colors
        @data[:indexed_colors] || EMPTY_ARRAY
      end

      # Returns array of MRU (most recently used) color dictionaries.
      #
      # @return [Array<Elements::Color>]
      # @api public
      #: () -> Array[Elements::Color]
      def mru_colors
        @mru_colors ||= (@data[:mru_colors] || EMPTY_ARRAY).map { |c| c.is_a?(Elements::Color) ? c : Elements::Color.new(**c) }.freeze
      end

      # Returns colors dictionary or palette.
      #
      # @return [Elements::Colors]
      # @api public
      #: () -> Elements::Colors
      def colors
        @colors ||= begin
          raw = @data[:colors] || EMPTY_HASH
          raw.is_a?(Elements::Colors) ? raw : Elements::Colors.new(**raw)
        end
      end

      # Looks up a font entry by 0-based index.
      #
      # @param id [Integer, nil]
      # @return [Elements::Font, nil]
      # @api public
      #: (Integer? id) -> Elements::Font?
      def font(id)
        return nil unless id.is_a?(Integer) && id >= 0

        fonts[id]
      end

      # Returns the font size in points for the font entry at index id.
      #
      # @param id [Integer, nil]
      # @return [Numeric, nil]
      # @api public
      #: (Integer? id) -> Numeric?
      def font_size(id)
        f = font(id)
        return nil unless f

        sz = f.size
        sz if sz.is_a?(Numeric)
      end

      # Looks up a fill entry by 0-based index.
      #
      # @param id [Integer, nil]
      # @return [Elements::Fill, nil]
      # @api public
      #: (Integer? id) -> Elements::Fill?
      def fill(id)
        return nil unless id.is_a?(Integer) && id >= 0

        fills[id]
      end

      # Looks up a border entry by 0-based index.
      #
      # @param id [Integer, nil]
      # @return [Elements::Border, nil]
      # @api public
      #: (Integer? id) -> Elements::Border?
      def border(id)
        return nil unless id.is_a?(Integer) && id >= 0

        borders[id]
      end

      # Looks up a cell format (cellXf) entry by 0-based index.
      #
      # @param id [Integer, nil]
      # @return [Elements::CellXf, nil]
      # @api public
      #: (Integer? id) -> Elements::CellXf?
      def cell_xf(id)
        return nil unless id.is_a?(Integer) && id >= 0

        cell_xfs[id]
      end

      # Looks up a custom format code by numFmtId.
      #
      # @param id [Integer, nil]
      # @return [String, nil]
      # @api public
      #: (Integer? id) -> String?
      def number_format_code(id)
        return nil unless id.is_a?(Integer)

        number_formats[id]
      end

      # Looks up a differential format (dxf) entry by 0-based index.
      #
      # @param id [Integer, nil]
      # @return [Elements::Dxf, nil]
      # @api public
      #: (Integer? id) -> Elements::Dxf?
      def dxf(id)
        return nil unless id.is_a?(Integer) && id >= 0

        dxfs[id]
      end

      # Looks up a cell style format (cellStyleXf) entry by 0-based index.
      #
      # @param id [Integer, nil]
      # @return [Elements::CellXf, nil]
      # @api public
      #: (Integer? id) -> Elements::CellXf?
      def cell_style_xf(id)
        return nil unless id.is_a?(Integer) && id >= 0

        cell_style_xfs[id]
      end

      # Looks up a named cell style entry by 0-based index.
      #
      # @param id [Integer, nil]
      # @return [Elements::CellStyle, nil]
      # @api public
      #: (Integer? id) -> Elements::CellStyle?
      def cell_style(id)
        return nil unless id.is_a?(Integer) && id >= 0

        cell_styles[id]
      end

      # Appends a font entry to the stylesheet and returns its 0-based index.
      #
      # @param spec [Hash, nil]
      # @return [Integer]
      # @api public
      #: (?Hash[Symbol, untyped]? spec, **untyped kwargs) -> Integer
      def add_font(spec = nil, **kwargs)
        entry = (spec || kwargs).dup
        entry[:size] ||= entry[:sz] if entry.key?(:sz)
        fonts_arr = (@data[:fonts] ||= [])
        fonts_arr << entry
        invalidate_cache!
        fonts_arr.size - 1
      end

      # Appends a fill entry to the stylesheet and returns its 0-based index.
      #
      # @param spec [Hash, nil]
      # @return [Integer]
      # @api public
      #: (?Hash[Symbol, untyped]? spec, **untyped kwargs) -> Integer
      def add_fill(spec = nil, **kwargs)
        entry = (spec || kwargs).dup
        entry[:pattern_type] ||= entry[:pattern] if entry.key?(:pattern)
        fills_arr = (@data[:fills] ||= [])
        fills_arr << entry
        invalidate_cache!
        fills_arr.size - 1
      end

      # Appends a border entry to the stylesheet and returns its 0-based index.
      #
      # @param spec [Hash, nil]
      # @return [Integer]
      # @api public
      #: (?Hash[Symbol, untyped]? spec, **untyped kwargs) -> Integer
      def add_border(spec = nil, **kwargs)
        entry = (spec || kwargs).dup
        borders_arr = (@data[:borders] ||= [])
        borders_arr << entry
        invalidate_cache!
        borders_arr.size - 1
      end

      # Sets a custom number format code and returns the format id.
      #
      # @param id [Integer] Format ID (custom formats typically start at 164).
      # @param code [String] Format code pattern.
      # @return [Integer]
      # @api public
      #: (Integer id, String code) -> Integer
      def add_number_format(id, code)
        raise ArgumentError, "num_fmt_id must be an Integer" unless id.is_a?(Integer)
        raise ArgumentError, "format_code must be a String" unless code.is_a?(String)

        (@data[:num_fmts] ||= {})[id] = code
        invalidate_cache!
        id
      end

      # Appends a cell format (cellXf) entry and returns its 0-based index.
      #
      # @param spec [Hash, nil]
      # @return [Integer]
      # @api public
      #: (?Hash[Symbol, untyped]? spec, **untyped kwargs) -> Integer
      def add_cell_xf(spec = nil, **kwargs)
        entry = (spec || kwargs).dup
        xfs_arr = (@data[:cell_xfs] ||= [])
        xfs_arr << entry
        invalidate_cache!
        xfs_arr.size - 1
      end

      # Appends a differential format (dxf) entry and returns its 0-based index.
      #
      # @param spec [Hash, nil]
      # @return [Integer]
      # @api public
      #: (?Hash[Symbol, untyped]? spec, **untyped kwargs) -> Integer
      def add_dxf(spec = nil, **kwargs)
        entry = (spec || kwargs).dup
        dxfs_arr = (@data[:dxfs] ||= [])
        dxfs_arr << entry
        invalidate_cache!
        dxfs_arr.size - 1
      end

      # Appends a cell style format (cellStyleXf) entry and returns its 0-based index.
      #
      # @param spec [Hash, nil]
      # @return [Integer]
      # @api public
      #: (?Hash[Symbol, untyped]? spec, **untyped kwargs) -> Integer
      def add_cell_style_xf(spec = nil, **kwargs)
        entry = (spec || kwargs).dup
        xfs_arr = (@data[:cell_style_xfs] ||= [])
        xfs_arr << entry
        invalidate_cache!
        xfs_arr.size - 1
      end

      # Appends a named cell style entry and returns its 0-based index.
      #
      # @param spec [Hash, nil]
      # @return [Integer]
      # @api public
      #: (?Hash[Symbol, untyped]? spec, **untyped kwargs) -> Integer
      def add_cell_style(spec = nil, **kwargs)
        entry = (spec || kwargs).dup
        cs_arr = (@data[:cell_styles] ||= [])
        cs_arr << entry
        invalidate_cache!
        cs_arr.size - 1
      end

      # High-level helper to compose and register a cell format (xf), adding subordinate
      # font, fill, border, and number format definitions as needed. Returns the xf_id.
      #
      # @param spec [Hash, nil]
      # @return [Integer]
      # @api public
      #: (?Hash[Symbol, untyped]? spec, **untyped kwargs) -> Integer
      def add_style(spec = nil, **kwargs)
        opts = (spec || kwargs).dup
        font_id = opts[:font_id] || (opts[:font] ? add_font(opts[:font]) : 0)
        fill_id = opts[:fill_id] || (opts[:fill] ? add_fill(opts[:fill]) : 0)
        border_id = opts[:border_id] || (opts[:border] ? add_border(opts[:border]) : 0)
        num_fmt_id = opts[:num_fmt_id] || 0
        if opts[:number_format]
          code = opts[:number_format].to_s
          existing_id = number_formats.key(code) || Ooxml::Utils::BUILTIN_NUM_FMT_CODES.key(code)
          num_fmt_id = existing_id || add_number_format(next_num_fmt_id, code)
        end

        xf_spec = {
          font_id: font_id,
          fill_id: fill_id,
          border_id: border_id,
          num_fmt_id: num_fmt_id
        }
        xf_spec[:alignment] = opts[:alignment] if opts[:alignment]
        xf_spec[:apply_alignment] = true if opts[:alignment] || opts[:apply_alignment]
        xf_spec[:apply_fill] = true if opts[:fill] || opts[:apply_fill]
        xf_spec[:apply_font] = true if opts[:font] || opts[:apply_font]
        xf_spec[:apply_border] = true if opts[:border] || opts[:apply_border]
        xf_spec[:apply_number_format] = true if opts[:number_format] || opts[:apply_number_format]
        xf_spec[:protection] = opts[:protection] if opts[:protection]
        xf_spec[:apply_protection] = true if opts[:protection] || opts[:apply_protection]

        add_cell_xf(xf_spec)
      end

      # Returns true if the styles container has no registered definitions.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def empty?
        @data.empty?
      end

      # Converts the styles container to a standard Hash.
      #
      # @return [Hash]
      # @api public
      #: () -> Hash[Symbol, untyped]
      def to_h
        @data.dup
      end
      alias to_hash to_h

      # Equality check against another Styles.
      #
      # @param other [Object]
      # @return [Boolean]
      # @api public
      #: (untyped other) -> bool
      def ==(other)
        other.is_a?(Styles) ? @data == other.to_h : false
      end
      alias eql? ==

      # Computes hash code based on underlying data.
      #
      # @return [Integer]
      # @api public
      #: () -> Integer
      def hash
        @data.hash
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
      # @api private
      #: () -> self
      def precompute!
        xfs = cell_xfs
        size = xfs.is_a?(Array) ? xfs.size : 0

        codes = Array.new(size)
        types = Array.new(size)
        d_flags = Array.new(size, false)
        dt_flags = Array.new(size, false)
        t_flags = Array.new(size, false)

        if size.positive?
          cs_xfs = cell_style_xfs
          n_fmts = number_formats

          size.times do |i|
            xf = xfs[i]
            next unless xf.is_a?(Hash) || xf.is_a?(Elements::CellXf)

            num_fmt_id = xf.is_a?(Elements::CellXf) ? xf.num_fmt_id : xf[:num_fmt_id]
            xf_id = xf.is_a?(Elements::CellXf) ? xf.xf_id : xf[:xf_id]
            if (num_fmt_id.nil? || num_fmt_id.zero?) && xf_id && cs_xfs.is_a?(Array)
              parent_xf = cs_xfs[xf_id]
              if parent_xf.is_a?(Elements::CellXf)
                num_fmt_id = parent_xf.num_fmt_id
              elsif parent_xf.is_a?(Hash)
                num_fmt_id = parent_xf[:num_fmt_id]
              end
            end

            code = num_fmt_id ? (n_fmts[num_fmt_id] || Ooxml::Utils::BUILTIN_NUM_FMT_CODES[num_fmt_id]) : nil

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

      private

      def next_num_fmt_id
        max_id = [163, *number_formats.keys].max
        max_id + 1
      end

      def invalidate_cache!
        @fonts = nil
        @fills = nil
        @borders = nil
        @cell_xfs = nil
        @dxfs = nil
        @cell_style_xfs = nil
        @cell_styles = nil
        @table_styles = nil
        @mru_colors = nil
        @colors = nil
        @format_codes = nil
        @format_types = nil
        @date_flags = nil
        @datetime_flags = nil
        @time_flags = nil
      end
    end
  end
end
