# frozen_string_literal: true

module Xlsxrb
  # Helper class for building cell styles with a fluent DSL.
  # Encapsulates font, fill, border, and number format properties.
  class StyleBuilder
    def initialize(name = nil)
      @name = name
      @font_props = {}
      @fill_props = {}
      @border_props = {}
      @num_fmt_id = nil
    end

    attr_reader :name, :font_props, :fill_props, :border_props, :num_fmt_id

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
      self
    end

    # --- Font Properties ---

    def bold(value = true)
      @font_props[:bold] = value
      self
    end

    def italic(value = true)
      @font_props[:italic] = value
      self
    end

    def size(sz)
      @font_props[:sz] = sz.to_i
      self
    end

    def font_name(name)
      @font_props[:name] = name
      self
    end

    def font_color(color)
      @font_props[:color] = color
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

    # --- Fill Properties ---

    def fill_pattern(pattern, fg_color: nil, bg_color: nil)
      @fill_props[:pattern] = pattern
      @fill_props[:fg_color] = fg_color if fg_color
      @fill_props[:bg_color] = bg_color if bg_color
      self
    end

    def fill_color(color)
      @fill_props[:pattern] = "solid"
      @fill_props[:fg_color] = color
      self
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
      color_opt = color ? { color: color } : {}
      @border_props[:left] = { style: style, **color_opt }
      @border_props[:right] = { style: style, **color_opt }
      @border_props[:top] = { style: style, **color_opt }
      @border_props[:bottom] = { style: style, **color_opt }
      self
    end

    def border_left(style: "thin", color: nil)
      @border_props[:left] = { style: style, color: color }.compact
      self
    end

    def border_right(style: "thin", color: nil)
      @border_props[:right] = { style: style, color: color }.compact
      self
    end

    def border_top(style: "thin", color: nil)
      @border_props[:top] = { style: style, color: color }.compact
      self
    end

    def border_bottom(style: "thin", color: nil)
      @border_props[:bottom] = { style: style, color: color }.compact
      self
    end

    # --- Number Format ---

    def number_format(num_fmt_id)
      @num_fmt_id = num_fmt_id
      self
    end

    # Register this style with the given Writer, returning the style_id.
    # writer:: Xlsxrb::Ooxml::Writer instance
    def register_with(writer)
      font_id = 0
      fill_id = 0
      border_id = 0

      font_id = writer.add_font(**@font_props) if @font_props.any?
      fill_id = writer.add_fill(**@fill_props) if @fill_props.any?
      border_id = writer.add_border(**@border_props) if @border_props.any?

      writer.add_cell_style(
        num_fmt_id: @num_fmt_id,
        font_id: font_id,
        fill_id: fill_id,
        border_id: border_id
      )
    end
  end
end
