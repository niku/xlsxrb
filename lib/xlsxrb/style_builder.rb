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
    def apply_options!(**opts)
      bold(opts[:bold]) if opts.key?(:bold)
      italic(opts[:italic]) if opts.key?(:italic)
      size(opts[:size]) if opts.key?(:size)
      font_name(opts[:font_name]) if opts.key?(:font_name)
      font_color(opts[:font_color]) if opts.key?(:font_color)
      underline(opts[:underline]) if opts.key?(:underline)
      strike(opts[:strike]) if opts.key?(:strike)
      vert_align(opts[:vert_align]) if opts.key?(:vert_align)

      fill_color(opts[:fill_color]) if opts.key?(:fill_color)
      if opts.key?(:fill_pattern)
        pattern = opts[:fill_pattern] || {}
        fill_pattern(
          pattern[:pattern] || "solid",
          fg_color: pattern[:fg_color],
          bg_color: pattern[:bg_color]
        )
      end
      fill_gradient(**opts[:fill_gradient]) if opts.key?(:fill_gradient) && opts[:fill_gradient]

      border_all(**opts[:border_all]) if opts.key?(:border_all) && opts[:border_all]
      border_left(**opts[:border_left]) if opts.key?(:border_left) && opts[:border_left]
      border_right(**opts[:border_right]) if opts.key?(:border_right) && opts[:border_right]
      border_top(**opts[:border_top]) if opts.key?(:border_top) && opts[:border_top]
      border_bottom(**opts[:border_bottom]) if opts.key?(:border_bottom) && opts[:border_bottom]

      number_format(opts[:number_format]) if opts.key?(:number_format)

      align_horizontal(opts[:align_horizontal]) if opts.key?(:align_horizontal)
      align_vertical(opts[:align_vertical]) if opts.key?(:align_vertical)
      wrap_text(opts[:wrap_text]) if opts.key?(:wrap_text)
      text_rotation(opts[:text_rotation]) if opts.key?(:text_rotation)
      indent(opts[:indent]) if opts.key?(:indent)
      shrink_to_fit(opts[:shrink_to_fit]) if opts.key?(:shrink_to_fit)

      self
    end

    # --- Font Properties ---

    # rubocop:disable Style/OptionalBooleanParameter
    def bold(value = true)
      @font_props[:bold] = value
      self
    end

    def italic(value = true)
      @font_props[:italic] = value
      self
    end

    def size(size_value)
      @font_props[:sz] = size_value.to_i
      self
    end

    def font_name(name)
      @font_props[:name] = name
      self
    end

    def font_color(color)
      @font_props[:color] = resolve_color(color)
      self
    end

    def underline(val = "single")
      @font_props[:underline] = val
      self
    end

    def strike(value = true)
      @font_props[:strike] = value
      self
    end

    def vert_align(value)
      @font_props[:vert_align] = value
      self
    end
    # rubocop:enable Style/OptionalBooleanParameter

    # --- Fill Properties ---

    def fill_pattern(pattern, fg_color: nil, bg_color: nil)
      @fill_props[:pattern] = pattern
      @fill_props[:fg_color] = fg_color if fg_color
      @fill_props[:bg_color] = bg_color if bg_color
      self
    end

    def fill_color(color)
      @fill_props[:pattern] = "solid"
      @fill_props[:fg_color] = resolve_color(color)
      self
    end

    def fill(pattern: "solid", fg_color: nil, bg_color: nil)
      fill_pattern(pattern, fg_color: resolve_color(fg_color), bg_color: resolve_color(bg_color))
    end

    def fill_gradient(type:, degree: nil, stops: [])
      @fill_props[:gradient] = {
        type: type,
        degree: degree,
        stops: stops
      }.compact
      self
    end

    # --- Border Properties ---

    def border_all(style: "thin", color: nil)
      color_opt = color ? { color: resolve_color(color) } : {}
      @border_props[:left] = { style: style, **color_opt }
      @border_props[:right] = { style: style, **color_opt }
      @border_props[:top] = { style: style, **color_opt }
      @border_props[:bottom] = { style: style, **color_opt }
      self
    end

    def border_left(style: "thin", color: nil)
      @border_props[:left] = { style: style, color: resolve_color(color) }.compact
      self
    end

    def border_right(style: "thin", color: nil)
      @border_props[:right] = { style: style, color: resolve_color(color) }.compact
      self
    end

    def border_top(style: "thin", color: nil)
      @border_props[:top] = { style: style, color: resolve_color(color) }.compact
      self
    end

    def border_bottom(style: "thin", color: nil)
      @border_props[:bottom] = { style: style, color: resolve_color(color) }.compact
      self
    end

    # rubocop:disable Naming/MethodParameterName
    def border_diagonal(style: "thin", color: nil, up: false, down: false)
      @border_props[:diagonal] = { style: style, color: resolve_color(color) }.compact
      @border_props[:diagonal_up] = true if up
      @border_props[:diagonal_down] = true if down
      self
    end
    # rubocop:enable Naming/MethodParameterName

    # --- Alignment Properties ---

    def align_horizontal(value)
      @alignment[:horizontal] = value
      self
    end

    def align_vertical(value)
      @alignment[:vertical] = value
      self
    end

    # rubocop:disable Style/OptionalBooleanParameter
    def wrap_text(value = true)
      @alignment[:wrap_text] = value
      self
    end

    def shrink_to_fit(value = true)
      @alignment[:shrink_to_fit] = value
      self
    end
    # rubocop:enable Style/OptionalBooleanParameter

    def text_rotation(value)
      @alignment[:text_rotation] = value
      self
    end

    def indent(value)
      @alignment[:indent] = value.to_i
      self
    end

    # --- Number Format ---

    def number_format(num_fmt_id)
      @num_fmt_id = num_fmt_id
      self
    end
    alias num_fmt number_format

    # Register this style with the given Writer, returning the style_id.
    # writer:: Xlsxrb::Ooxml::Writer instance
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
