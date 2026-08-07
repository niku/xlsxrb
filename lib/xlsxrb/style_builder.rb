# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  # Helper class for building cell styles with a fluent DSL.
  # Encapsulates font, fill, border, alignment, and number format properties.
  class StyleBuilder
    COLORS = {
      black: "FF000000",
      white: "FFFFFFFF",
      red: "FFFF0000",
      green: "FF00FF00",
      blue: "FF0000FF",
      yellow: "FFFFFF00",
      cyan: "FF00FFFF",
      magenta: "FFFF00FF",
      gray: "FF808080",
      grey: "FF808080"
    }.freeze

    def resolve_color(color)
      return nil unless color

      if color.is_a?(Symbol) || (color.is_a?(String) && color.start_with?(":") && color.length > 1)
        key = color.to_s.sub(/^:/, "").to_sym
        return COLORS[key] || color.to_s
      end
      color.to_s
    end

    # : (?String? name) -> void
    def initialize(name = nil)
      @name = name
      @font_props = {}
      @fill_props = {}
      @border_props = {}
      @num_fmt_id = nil
      @alignment = {}
    end

    attr_reader :name, :font_props, :fill_props, :border_props, :num_fmt_id, :alignment

    # Applies option-style definitions so callers can use add_style(name, **opts)
    # as an alternative to block-based fluent chaining.
    # : (**untyped opts) -> self
    def apply_options!(**opts)
      if opts.key?(:font)
        font_opts = opts[:font] || {}
        bold(font_opts[:bold]) if font_opts.key?(:bold)
        italic(font_opts[:italic]) if font_opts.key?(:italic)
        size(font_opts[:size]) if font_opts.key?(:size)
        font_name(font_opts[:name]) if font_opts.key?(:name)
        font_color(font_opts[:color]) if font_opts.key?(:color)
        underline(font_opts[:underline]) if font_opts.key?(:underline)
        strike(font_opts[:strike]) if font_opts.key?(:strike)
        vert_align(font_opts[:vert_align]) if font_opts.key?(:vert_align)
      else
        bold(opts[:bold]) if opts.key?(:bold)
        italic(opts[:italic]) if opts.key?(:italic)
        size(opts[:size]) if opts.key?(:size)
        font_name(opts[:font_name]) if opts.key?(:font_name)
        font_color(opts[:font_color]) if opts.key?(:font_color)
        underline(opts[:underline]) if opts.key?(:underline)
        strike(opts[:strike]) if opts.key?(:strike)
        vert_align(opts[:vert_align]) if opts.key?(:vert_align)
      end

      if opts.key?(:fill)
        fill_opts = opts[:fill] || {}
        if fill_opts.key?(:color)
          fill_color(fill_opts[:color])
        else
          fill_pattern(
            fill_opts[:pattern] || "solid",
            fg_color: fill_opts[:fg_color],
            bg_color: fill_opts[:bg_color]
          )
        end
      else
        fill_color(opts[:fill_color]) if opts.key?(:fill_color)
        if opts.key?(:fill_pattern)
          pattern = opts[:fill_pattern] || {}
          fill_pattern(
            pattern[:pattern] || "solid",
            fg_color: pattern[:fg_color],
            bg_color: pattern[:bg_color]
          )
        end
      end
      fill_gradient(**opts[:fill_gradient]) if opts.key?(:fill_gradient) && opts[:fill_gradient]

      if opts.key?(:border)
        border_opts = opts[:border] || {}
        border_all(**border_opts[:all]) if border_opts.key?(:all) && border_opts[:all]
        border_left(**border_opts[:left]) if border_opts.key?(:left) && border_opts[:left]
        border_right(**border_opts[:right]) if border_opts.key?(:right) && border_opts[:right]
        border_top(**border_opts[:top]) if border_opts.key?(:top) && border_opts[:top]
        border_bottom(**border_opts[:bottom]) if border_opts.key?(:bottom) && border_opts[:bottom]
      else
        border_all(**opts[:border_all]) if opts.key?(:border_all) && opts[:border_all]
        border_left(**opts[:border_left]) if opts.key?(:border_left) && opts[:border_left]
        border_right(**opts[:border_right]) if opts.key?(:border_right) && opts[:border_right]
        border_top(**opts[:border_top]) if opts.key?(:border_top) && opts[:border_top]
        border_bottom(**opts[:border_bottom]) if opts.key?(:border_bottom) && opts[:border_bottom]
      end

      number_format(opts[:number_format]) if opts.key?(:number_format)

      if opts.key?(:alignment)
        align_opts = opts[:alignment] || {}
        align_horizontal(align_opts[:horizontal]) if align_opts.key?(:horizontal)
        align_vertical(align_opts[:vertical]) if align_opts.key?(:vertical)
        wrap_text(align_opts[:wrap_text]) if align_opts.key?(:wrap_text)
        text_rotation(align_opts[:text_rotation]) if align_opts.key?(:text_rotation)
        indent(align_opts[:indent]) if align_opts.key?(:indent)
        shrink_to_fit(align_opts[:shrink_to_fit]) if align_opts.key?(:shrink_to_fit)
      else
        align_horizontal(opts[:align_horizontal]) if opts.key?(:align_horizontal)
        align_vertical(opts[:align_vertical]) if opts.key?(:align_vertical)
        wrap_text(opts[:wrap_text]) if opts.key?(:wrap_text)
        text_rotation(opts[:text_rotation]) if opts.key?(:text_rotation)
        indent(opts[:indent]) if opts.key?(:indent)
        shrink_to_fit(opts[:shrink_to_fit]) if opts.key?(:shrink_to_fit)
      end

      self
    end

    # --- Font Properties ---

    # : (**untyped opts) -> self
    def font(**opts)
      bold(opts[:bold]) if opts.key?(:bold)
      italic(opts[:italic]) if opts.key?(:italic)
      size(opts[:size]) if opts.key?(:size)
      font_name(opts[:name]) if opts.key?(:name)
      font_color(opts[:color]) if opts.key?(:color)
      underline(opts[:underline]) if opts.key?(:underline)
      strike(opts[:strike]) if opts.key?(:strike)
      vert_align(opts[:vert_align]) if opts.key?(:vert_align)
      self
    end

    # rubocop:disable Style/OptionalBooleanParameter
    # : (?bool value) -> self
    def bold(value = true)
      @font_props[:bold] = value
      self
    end

    # : (?bool value) -> self
    def italic(value = true)
      @font_props[:italic] = value
      self
    end

    # : (Numeric size_value) -> self
    def size(size_value)
      @font_props[:sz] = size_value.to_i
      self
    end

    # : (String name) -> self
    def font_name(name)
      @font_props[:name] = name
      self
    end

    # : (String | Symbol color) -> self
    def font_color(color)
      @font_props[:color] = resolve_color(color)
      self
    end

    # : (?String val) -> self
    def underline(val = "single")
      @font_props[:underline] = val
      self
    end

    # : (?bool value) -> self
    def strike(value = true)
      @font_props[:strike] = value
      self
    end

    # : (String value) -> self
    def vert_align(value)
      @font_props[:vert_align] = value
      self
    end
    # rubocop:enable Style/OptionalBooleanParameter

    # --- Fill Properties ---

    # : (String | Symbol pattern, ?fg_color: String | Symbol | nil, ?bg_color: String | Symbol | nil) -> self
    def fill_pattern(pattern, fg_color: nil, bg_color: nil)
      @fill_props[:pattern] = pattern
      @fill_props[:fg_color] = fg_color if fg_color
      @fill_props[:bg_color] = bg_color if bg_color
      self
    end

    # : (String | Symbol color) -> self
    def fill_color(color)
      @fill_props[:pattern] = "solid"
      @fill_props[:fg_color] = resolve_color(color)
      self
    end

    # : (?pattern: String | Symbol, ?fg_color: String | Symbol | nil, ?bg_color: String | Symbol | nil) -> self
    def fill(pattern: "solid", fg_color: nil, bg_color: nil)
      fill_pattern(pattern, fg_color: resolve_color(fg_color), bg_color: resolve_color(bg_color))
    end

    # : (type: String, ?degree: Numeric | nil, ?stops: Array[untyped]) -> self
    def fill_gradient(type:, degree: nil, stops: [])
      @fill_props[:gradient] = {
        type: type,
        degree: degree,
        stops: stops
      }.compact
      self
    end

    # --- Border Properties ---

    # : (**untyped opts) -> self
    def border(**opts)
      border_all(**opts[:all]) if opts.key?(:all) && opts[:all]
      border_left(**opts[:left]) if opts.key?(:left) && opts[:left]
      border_right(**opts[:right]) if opts.key?(:right) && opts[:right]
      border_top(**opts[:top]) if opts.key?(:top) && opts[:top]
      border_bottom(**opts[:bottom]) if opts.key?(:bottom) && opts[:bottom]
      self
    end

    # : (?style: String | Symbol, ?color: String | Symbol | nil) -> self
    def border_all(style: "thin", color: nil)
      color_opt = color ? { color: resolve_color(color) } : {}
      @border_props[:left] = { style: style, **color_opt }
      @border_props[:right] = { style: style, **color_opt }
      @border_props[:top] = { style: style, **color_opt }
      @border_props[:bottom] = { style: style, **color_opt }
      self
    end

    # : (?style: String | Symbol, ?color: String | Symbol | nil) -> self
    def border_left(style: "thin", color: nil)
      @border_props[:left] = { style: style, color: resolve_color(color) }.compact
      self
    end

    # : (?style: String | Symbol, ?color: String | Symbol | nil) -> self
    def border_right(style: "thin", color: nil)
      @border_props[:right] = { style: style, color: resolve_color(color) }.compact
      self
    end

    # : (?style: String | Symbol, ?color: String | Symbol | nil) -> self
    def border_top(style: "thin", color: nil)
      @border_props[:top] = { style: style, color: resolve_color(color) }.compact
      self
    end

    # : (?style: String | Symbol, ?color: String | Symbol | nil) -> self
    def border_bottom(style: "thin", color: nil)
      @border_props[:bottom] = { style: style, color: resolve_color(color) }.compact
      self
    end

    # rubocop:disable Naming/MethodParameterName
    # : (?style: String | Symbol, ?color: String | Symbol | nil, ?up: bool, ?down: bool) -> self
    def border_diagonal(style: "thin", color: nil, up: false, down: false)
      @border_props[:diagonal] = { style: style, color: resolve_color(color) }.compact
      @border_props[:diagonal_up] = true if up
      @border_props[:diagonal_down] = true if down
      self
    end
    # rubocop:enable Naming/MethodParameterName

    # --- Alignment Properties ---

    # : (String | Symbol value) -> self
    def align_horizontal(value)
      @alignment[:horizontal] = value
      self
    end

    # : (String | Symbol value) -> self
    def align_vertical(value)
      @alignment[:vertical] = value
      self
    end

    # rubocop:disable Style/OptionalBooleanParameter
    # : (?bool value) -> self
    def wrap_text(value = true)
      @alignment[:wrap_text] = value
      self
    end

    # : (?bool value) -> self
    def shrink_to_fit(value = true)
      @alignment[:shrink_to_fit] = value
      self
    end
    # rubocop:enable Style/OptionalBooleanParameter

    # : (Numeric value) -> self
    def text_rotation(value)
      @alignment[:text_rotation] = value
      self
    end

    # : (Numeric value) -> self
    def indent(value)
      @alignment[:indent] = value.to_i
      self
    end

    # --- Number Format ---

    # : (String | Integer num_fmt_id) -> self
    def number_format(num_fmt_id)
      @num_fmt_id = num_fmt_id
      self
    end
    alias num_fmt number_format

    # Register this style with the given Writer, returning the style_id.
    # writer:: Xlsxrb::Ooxml::Writer instance
    # : (untyped writer) -> Integer
    def register_with(writer)
      font_id = 0
      fill_id = 0
      border_id = 0

      font_id = writer.add_font(**@font_props) if @font_props.any?
      fill_id = writer.add_fill(**@fill_props) if @fill_props.any?
      border_id = writer.add_border(**@border_props) if @border_props.any?

      resolved_num_fmt_id = if @num_fmt_id.is_a?(String)
                              writer.add_number_format(@num_fmt_id)
                            else
                              @num_fmt_id
                            end

      cell_style_opts = {
        num_fmt_id: resolved_num_fmt_id,
        font_id: font_id,
        fill_id: fill_id,
        border_id: border_id
      }
      cell_style_opts[:alignment] = @alignment if @alignment.any?

      writer.add_cell_style(**cell_style_opts)
    end
  end
end
