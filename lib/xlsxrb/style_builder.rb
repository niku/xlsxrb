# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  # Helper class for building cell styles with a fluent DSL.
  # Encapsulates font, fill, border, alignment, and number format properties.
  # @api public
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

    CSS_COLORS = {
      aliceblue: "F0F8FF",
      antiquewhite: "FAEBD7",
      aqua: "00FFFF",
      aquamarine: "7FFFD4",
      azure: "F0FFFF",
      beige: "F5F5DC",
      bisque: "FFE4C4",
      black: "000000",
      blanchedalmond: "FFEBCD",
      blue: "0000FF",
      blueviolet: "8A2BE2",
      brown: "A52A2A",
      burlywood: "DEB887",
      cadetblue: "5F9EA0",
      chartreuse: "7FFF00",
      chocolate: "D2691E",
      coral: "FF7F50",
      cornflowerblue: "6495ED",
      cornsilk: "FFF8DC",
      crimson: "DC143C",
      cyan: "00FFFF",
      darkblue: "00008B",
      darkcyan: "008B8B",
      darkgoldenrod: "B8860B",
      darkgray: "A9A9A9",
      darkgrey: "A9A9A9",
      darkgreen: "006400",
      darkkhaki: "BDB76B",
      darkmagenta: "8B008B",
      darkolivegreen: "556B2F",
      darkorange: "FF8C00",
      darkorchid: "9932CC",
      darkred: "8B0000",
      darksalmon: "E9967A",
      darkseagreen: "8FBC8F",
      darkslateblue: "483D8B",
      darkslategray: "2F4F4F",
      darkslategrey: "2F4F4F",
      darkturquoise: "00CED1",
      darkviolet: "9400D3",
      deeppink: "FF1493",
      deepskyblue: "00BFFF",
      dimgray: "696969",
      dimgrey: "696969",
      dodgerblue: "1E90FF",
      firebrick: "B22222",
      floralwhite: "FFFAF0",
      forestgreen: "228B22",
      fuchsia: "FF00FF",
      gainsboro: "DCDCDC",
      ghostwhite: "F8F8FF",
      gold: "FFD700",
      goldenrod: "DAA520",
      gray: "808080",
      grey: "808080",
      green: "00FF00",
      greenyellow: "ADFF2F",
      honeydew: "F0FFF0",
      hotpink: "FF69B4",
      indianred: "CD5C5C",
      indigo: "4B0082",
      ivory: "FFFFF0",
      khaki: "F0E68C",
      lavender: "E6E6FA",
      lavenderblush: "FFF0F5",
      lawngreen: "7CFC00",
      lemonchiffon: "FFFACD",
      lightblue: "ADD8E6",
      lightcoral: "F08080",
      lightcyan: "E0FFFF",
      lightgoldenrodyellow: "FAFAD2",
      lightgray: "D3D3D3",
      lightgrey: "D3D3D3",
      lightgreen: "90EE90",
      lightpink: "FFB6C1",
      lightsalmon: "FFA07A",
      lightseagreen: "20B2AA",
      lightskyblue: "87CEFA",
      lightslategray: "778899",
      lightslategrey: "778899",
      lightsteelblue: "B0C4DE",
      lightyellow: "FFFFE0",
      lime: "00FF00",
      limegreen: "32CD32",
      linen: "FAF0E6",
      magenta: "FF00FF",
      maroon: "800000",
      mediumaquamarine: "66CDAA",
      mediumblue: "0000CD",
      mediumorchid: "BA55D3",
      mediumpurple: "9370DB",
      mediumseagreen: "3CB371",
      mediumslateblue: "7B68EE",
      mediumspringgreen: "00FA9A",
      mediumturquoise: "48D1CC",
      mediumvioletred: "C71585",
      midnightblue: "191970",
      mintcream: "F5FFFA",
      mistyrose: "FFE4E1",
      moccasin: "FFE4B5",
      navajowhite: "FFDEAD",
      navy: "000080",
      oldlace: "FDF5E6",
      olive: "808000",
      olivedrab: "6B8E23",
      orange: "FFA500",
      orangered: "FF4500",
      orchid: "DA70D6",
      palegoldenrod: "EEE8AA",
      palegreen: "98FB98",
      paleturquoise: "AFEEEE",
      palevioletred: "DB7093",
      papayawhip: "FFEFD5",
      peachpuff: "FFDAB9",
      peru: "CD853F",
      pink: "FFC0CB",
      plum: "DDA0DD",
      powderblue: "B0E0E6",
      purple: "800080",
      rebeccapurple: "663399",
      red: "FF0000",
      rosybrown: "BC8F8F",
      royalblue: "4169E1",
      saddlebrown: "8B4513",
      salmon: "FA8072",
      sandybrown: "F4A460",
      seagreen: "2E8B57",
      seashell: "FFF5EE",
      sienna: "A0522D",
      silver: "C0C0C0",
      skyblue: "87CEEB",
      slateblue: "6A5ACD",
      slategray: "708090",
      slategrey: "708090",
      snow: "FFFAFA",
      springgreen: "00FF7F",
      steelblue: "4682B4",
      tan: "D2B48C",
      teal: "008080",
      thistle: "D8BFD8",
      tomato: "FF6347",
      turquoise: "40E0D0",
      violet: "EE82EE",
      wheat: "F5DEB3",
      white: "FFFFFF",
      whitesmoke: "F5F5F5",
      yellow: "FFFF00",
      yellowgreen: "9ACD32"
    }.freeze

    BORDER_STYLES = {
      none: :none,
      thin: :thin,
      medium: :medium,
      dashed: :dashed,
      dotted: :dotted,
      thick: :thick,
      double: :double,
      hair: :hair,
      medium_dashed: :mediumDashed,
      mediumdashed: :mediumDashed,
      mediumDashed: :mediumDashed,
      dash_dot: :dashDot,
      dashdot: :dashDot,
      dashDot: :dashDot,
      medium_dash_dot: :mediumDashDot,
      mediumdashdot: :mediumDashDot,
      mediumDashDot: :mediumDashDot,
      dash_dot_dot: :dashDotDot,
      dashdotdot: :dashDotDot,
      dashDotDot: :dashDotDot,
      medium_dash_dot_dot: :mediumDashDotDot,
      mediumdashdotdot: :mediumDashDotDot,
      mediumDashDotDot: :mediumDashDotDot,
      slant_dash_dot: :slantDashDot,
      slantdashdot: :slantDashDot,
      slantDashDot: :slantDashDot
    }.freeze

    BORDER_INDEX_MAP = {
      0 => :none,
      1 => :thin,
      2 => :medium,
      3 => :dashed,
      4 => :dotted,
      5 => :thick,
      6 => :double,
      7 => :hair,
      8 => :mediumDashed,
      9 => :dashDot,
      10 => :mediumDashDot,
      11 => :dashDotDot,
      12 => :mediumDashDotDot,
      13 => :slantDashDot
    }.freeze

    # Normalizes a color definition (CSS color name, hex code, or RGB integer) to an OpenXML hex color string.
    #
    # @example
    #   StyleBuilder.normalize_color(:navy)                 #=> "FF000080"
    #   StyleBuilder.normalize_color("#FF0000")             #=> "FFFF0000"
    #   StyleBuilder.normalize_color("#F00")                #=> "FFFF0000"
    #   StyleBuilder.normalize_color("FF0000", alpha: false) #=> "FF0000"
    #   StyleBuilder.normalize_color(0xFF0000)              #=> "FFFF0000"
    #
    # @param color [String, Symbol, Integer, nil]
    # @param alpha [Boolean] Whether to include the alpha channel ("AARRGGBB" vs "RRGGBB").
    # @return [String, nil]
    # @api public
    #: (String | Symbol | Integer | nil color, ?alpha: bool) -> String?
    def self.normalize_color(color, alpha: true)
      return nil if color.nil?

      hex = case color
            when Symbol, defined?(String) && String
              str = color.to_s.strip
              return nil if str.empty?

              # Strip leading colon (e.g. ":red")
              str = str.sub(/^:/, "")
              return nil if str.empty?

              sym_key = str.downcase.to_sym
              if CSS_COLORS.key?(sym_key)
                CSS_COLORS[sym_key]
              else
                # Hex string parsing
                clean_hex = str.sub(/^#/, "")
                case clean_hex.length
                when 3
                  if clean_hex.match?(/\A[0-9a-fA-F]{3}\z/)
                    r, g, b = clean_hex.chars
                    "#{r}#{r}#{g}#{g}#{b}#{b}".upcase
                  end
                when 6
                  clean_hex.upcase if clean_hex.match?(/\A[0-9a-fA-F]{6}\z/)
                when 8
                  if clean_hex.match?(/\A[0-9a-fA-F]{8}\z/)
                    return alpha ? clean_hex.upcase : clean_hex[2..7].upcase
                  end
                end
              end
            when Integer
              format("%06X", color & 0xFFFFFF)
            end

      return nil unless hex

      alpha ? "FF#{hex}" : hex
    end

    # Normalizes a border style symbol, string, or integer index to an OOXML ST_BorderStyle symbol.
    #
    # @example
    #   StyleBuilder.normalize_border_style(:border_thin)      #=> :thin
    #   StyleBuilder.normalize_border_style(1)                 #=> :thin
    #   StyleBuilder.normalize_border_style(:medium_dashed)    #=> :mediumDashed
    #   StyleBuilder.normalize_border_style("double")          #=> :double
    #   StyleBuilder.normalize_border_style(0)                 #=> :none
    #
    # @param style [Symbol, String, Integer, nil]
    # @return [Symbol, nil]
    # @api public
    #: (Symbol | String | Integer | nil style) -> Symbol?
    def self.normalize_border_style(style)
      return nil if style.nil?

      return BORDER_INDEX_MAP[style] if style.is_a?(Integer)

      str = style.to_s.sub(/^border_/, "")
      sym = str.to_sym
      BORDER_STYLES[sym] || BORDER_STYLES[str.downcase.to_sym]
    end

    # Resolves a color symbol, string, or prefix to an aRGB hex string.
    #
    # @param color [Symbol, String, Object, nil]
    # @return [String, nil]
    #: (untyped color) -> String?
    def self.resolve_color(color)
      return nil unless color

      if color.is_a?(Symbol) || (color.is_a?(String) && color.start_with?(":") && color.length > 1)
        key = color.to_s.sub(/^:/, "").to_sym
        return COLORS[key] || (CSS_COLORS[key] ? "FF#{CSS_COLORS[key]}" : nil) || color.to_s
      end

      return normalize_color(color, alpha: true) || color.to_s if color.is_a?(String) && color.start_with?("#")

      color.to_s
    end

    def resolve_color(color)
      self.class.resolve_color(color)
    end

    #: (?String? name) -> void
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
    # Applies option-style definitions so callers can use add_style(name, **opts)
    # as an alternative to block-based fluent chaining.
    # @param opts [Hash] The styling options.
    # @return [self]
    # @api public
    #: (**untyped) -> self
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

    # Configures multiple font properties at once.
    # @param opts [Hash] The font properties.
    # @return [self]
    # @api public
    #: (**untyped) -> self
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
    # Sets the font to bold.
    # @param value [Boolean] Whether to apply bold.
    # @return [self]
    # @api public
    #: (?bool) -> self
    def bold(value = true)
      @font_props[:bold] = value
      self
    end

    # Sets the font to italic.
    # @param value [Boolean] Whether to apply italic.
    # @return [self]
    # @api public
    #: (?bool) -> self
    def italic(value = true)
      @font_props[:italic] = value
      self
    end

    # Sets the font size.
    # @param size_value [Numeric] The size.
    # @return [self]
    # @api public
    #: (Numeric) -> self
    def size(size_value)
      @font_props[:sz] = size_value.to_i
      self
    end

    # Sets the font name.
    # @param name [String] The font name.
    # @return [self]
    # @api public
    #: (String) -> self
    def font_name(name)
      @font_props[:name] = name
      self
    end

    # Sets the font color.
    # @param color [String, Symbol] The color.
    # @return [self]
    # @api public
    #: (String | Symbol) -> self
    def font_color(color)
      @font_props[:color] = resolve_color(color)
      self
    end

    # Sets the font underline style.
    # @param val [String] The underline style (e.g., 'single').
    # @return [self]
    # @api public
    #: (?String) -> self
    def underline(val = "single")
      @font_props[:underline] = val
      self
    end

    # Sets the font strikethrough.
    # @param value [Boolean] Whether to apply strike.
    # @return [self]
    # @api public
    #: (?bool) -> self
    def strike(value = true)
      @font_props[:strike] = value
      self
    end

    # Sets the vertical alignment of the font.
    # @param value [String] The alignment value.
    # @return [self]
    # @api public
    #: (String) -> self
    def vert_align(value)
      @font_props[:vert_align] = value
      self
    end
    # rubocop:enable Style/OptionalBooleanParameter

    # --- Fill Properties ---

    # Sets the fill pattern.
    # @param pattern [String, Symbol] The pattern type.
    # @param fg_color [String, Symbol, nil] The foreground color.
    # @param bg_color [String, Symbol, nil] The background color.
    # @return [self]
    # @api public
    # Configures fill properties.
    # @param pattern [String, Symbol] The pattern type.
    # @param fg_color [String, Symbol, nil] The foreground color.
    # @param bg_color [String, Symbol, nil] The background color.
    # @return [self]
    # @api public
    #: (String | Symbol pattern, ?fg_color: String | Symbol | nil, ?bg_color: String | Symbol | nil) -> self
    def fill_pattern(pattern, fg_color: nil, bg_color: nil)
      @fill_props[:pattern] = pattern
      @fill_props[:fg_color] = fg_color if fg_color
      @fill_props[:bg_color] = bg_color if bg_color
      self
    end

    # Sets a solid fill color.
    # @param color [String, Symbol] The color.
    # @return [self]
    # @api public
    #: (String | Symbol) -> self
    def fill_color(color)
      @fill_props[:pattern] = "solid"
      @fill_props[:fg_color] = resolve_color(color)
      self
    end

    #: (?pattern: String | Symbol, ?fg_color: String | Symbol | nil, ?bg_color: String | Symbol | nil) -> self
    def fill(pattern: "solid", fg_color: nil, bg_color: nil)
      fill_pattern(pattern, fg_color: resolve_color(fg_color), bg_color: resolve_color(bg_color))
    end

    # Sets a gradient fill.
    # @param type [String] The gradient type.
    # @param degree [Numeric, nil] The degree.
    # @param stops [Array] The gradient stops.
    # @return [self]
    # @api public
    #: (type: String, ?degree: Numeric | nil, ?stops: Array[untyped]) -> self
    def fill_gradient(type:, degree: nil, stops: [])
      @fill_props[:gradient] = {
        type: type,
        degree: degree,
        stops: stops
      }.compact
      self
    end

    # --- Border Properties ---

    # Configures multiple border properties.
    # @param opts [Hash] Border options.
    # @return [self]
    # @api public
    #: (**untyped) -> self
    def border(**opts)
      border_all(**opts[:all]) if opts.key?(:all) && opts[:all]
      border_left(**opts[:left]) if opts.key?(:left) && opts[:left]
      border_right(**opts[:right]) if opts.key?(:right) && opts[:right]
      border_top(**opts[:top]) if opts.key?(:top) && opts[:top]
      border_bottom(**opts[:bottom]) if opts.key?(:bottom) && opts[:bottom]
      self
    end

    # Sets all borders.
    # @param style [String, Symbol, Integer] The border style.
    # @param color [String, Symbol, nil] The color.
    # @return [self]
    # @api public
    #: (?style: String | Symbol | Integer, ?color: String | Symbol | nil) -> self
    def border_all(style: "thin", color: nil)
      norm_style = self.class.normalize_border_style(style)&.to_s || style.to_s
      color_opt = color ? { color: resolve_color(color) } : {}
      @border_props[:left] = { style: norm_style, **color_opt }
      @border_props[:right] = { style: norm_style, **color_opt }
      @border_props[:top] = { style: norm_style, **color_opt }
      @border_props[:bottom] = { style: norm_style, **color_opt }
      self
    end

    # Sets the left border.
    # @param style [String, Symbol, Integer] The border style.
    # @param color [String, Symbol, nil] The color.
    # @return [self]
    # @api public
    #: (?style: String | Symbol | Integer, ?color: String | Symbol | nil) -> self
    def border_left(style: "thin", color: nil)
      norm_style = self.class.normalize_border_style(style)&.to_s || style.to_s
      @border_props[:left] = { style: norm_style, color: resolve_color(color) }.compact
      self
    end

    # Sets the right border.
    # @param style [String, Symbol, Integer] The border style.
    # @param color [String, Symbol, nil] The color.
    # @return [self]
    # @api public
    #: (?style: String | Symbol | Integer, ?color: String | Symbol | nil) -> self
    def border_right(style: "thin", color: nil)
      norm_style = self.class.normalize_border_style(style)&.to_s || style.to_s
      @border_props[:right] = { style: norm_style, color: resolve_color(color) }.compact
      self
    end

    # Sets the top border.
    # @param style [String, Symbol, Integer] The border style.
    # @param color [String, Symbol, nil] The color.
    # @return [self]
    # @api public
    #: (?style: String | Symbol | Integer, ?color: String | Symbol | nil) -> self
    def border_top(style: "thin", color: nil)
      norm_style = self.class.normalize_border_style(style)&.to_s || style.to_s
      @border_props[:top] = { style: norm_style, color: resolve_color(color) }.compact
      self
    end

    # Sets the bottom border.
    # @param style [String, Symbol, Integer] The border style.
    # @param color [String, Symbol, nil] The color.
    # @return [self]
    # @api public
    #: (?style: String | Symbol | Integer, ?color: String | Symbol | nil) -> self
    def border_bottom(style: "thin", color: nil)
      norm_style = self.class.normalize_border_style(style)&.to_s || style.to_s
      @border_props[:bottom] = { style: norm_style, color: resolve_color(color) }.compact
      self
    end

    # rubocop:disable Naming/MethodParameterName
    # Sets diagonal borders.
    # @param style [String, Symbol, Integer] The border style.
    # @param color [String, Symbol, nil] The color.
    # @param up [Boolean] Diagonal up.
    # @param down [Boolean] Diagonal down.
    # @return [self]
    # @api public
    #: (?style: String | Symbol | Integer, ?color: String | Symbol | nil, ?up: bool, ?down: bool) -> self
    def border_diagonal(style: "thin", color: nil, up: false, down: false)
      norm_style = self.class.normalize_border_style(style)&.to_s || style.to_s
      @border_props[:diagonal] = { style: norm_style, color: resolve_color(color) }.compact
      @border_props[:diagonal_up] = true if up
      @border_props[:diagonal_down] = true if down
      self
    end
    # rubocop:enable Naming/MethodParameterName

    # --- Alignment Properties ---

    # Sets horizontal alignment.
    # @param value [String, Symbol] The alignment.
    # @return [self]
    # @api public
    #: (String | Symbol) -> self
    def align_horizontal(value)
      @alignment[:horizontal] = value
      self
    end

    # Sets vertical alignment.
    # @param value [String, Symbol] The alignment.
    # @return [self]
    # @api public
    #: (String | Symbol) -> self
    def align_vertical(value)
      @alignment[:vertical] = value
      self
    end

    # rubocop:disable Style/OptionalBooleanParameter
    # Sets text wrapping.
    # @param value [Boolean] Whether to wrap text.
    # @return [self]
    # @api public
    #: (?bool) -> self
    def wrap_text(value = true)
      @alignment[:wrap_text] = value
      self
    end

    # Sets shrink to fit.
    # @param value [Boolean] Whether to shrink text.
    # @return [self]
    # @api public
    #: (?bool) -> self
    def shrink_to_fit(value = true)
      @alignment[:shrink_to_fit] = value
      self
    end
    # rubocop:enable Style/OptionalBooleanParameter

    # Sets text rotation.
    # @param value [Numeric] The rotation angle.
    # @return [self]
    # @api public
    #: (Numeric) -> self
    def text_rotation(value)
      @alignment[:text_rotation] = value
      self
    end

    # Sets text indent.
    # @param value [Numeric] The indent level.
    # @return [self]
    # @api public
    #: (Numeric) -> self
    def indent(value)
      @alignment[:indent] = value.to_i
      self
    end

    # --- Number Format ---

    # Sets the number format.
    # @param num_fmt_id [String, Integer] The format id or format string.
    # @return [self]
    # @api public
    #: (String | Integer) -> self
    def number_format(num_fmt_id)
      @num_fmt_id = num_fmt_id
      self
    end
    alias num_fmt number_format

    # Register this style with the given Writer, returning the style_id.
    # writer:: Xlsxrb::Ooxml::Writer instance
    #: (untyped writer) -> Integer
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
