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

    UNDERLINE_MAP = {
      single: "single",
      double: "double",
      single_accounting: "singleAccounting",
      double_accounting: "doubleAccounting",
      none: "none"
    }.freeze

    VERT_ALIGN_MAP = {
      superscript: "superscript",
      super: "superscript",
      subscript: "subscript",
      sub: "subscript",
      baseline: "baseline"
    }.freeze
    OMITTED = Object.new.freeze
    private_constant :COLORS, :CSS_COLORS, :BORDER_STYLES, :BORDER_INDEX_MAP, :OMITTED

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
    # @api private
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
    # @api private
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

      return Colors.palette_color(color, alpha: true) if color.is_a?(Integer) && defined?(Colors) && color >= 0 && color < Colors::INDEXED_COLORS.size

      if color.is_a?(Symbol) || (color.is_a?(String) && color.start_with?(":") && color.length > 1)
        key = color.to_s.sub(/^:/, "").to_sym
        return COLORS[key] || (CSS_COLORS[key] ? "FF#{CSS_COLORS[key]}" : nil) || color.to_s
      end

      return normalize_color(color, alpha: true) || color.to_s if color.is_a?(String) && color.start_with?("#")

      color.to_s
    end

    #: (?String? name) -> void
    def initialize(name = nil)
      @name = name
      @font_props = {}
      @fill_props = {}
      @border_props = {}
      @num_fmt_id = nil
      @alignment = {}
      @protection = {}
    end

    #: (String | Symbol)?
    attr_reader :name

    # @api private
    #: Hash[Symbol, untyped]
    attr_reader :font_props

    # @api private
    #: Hash[Symbol, untyped]
    attr_reader :fill_props

    # @api private
    #: Hash[Symbol, untyped]
    attr_reader :border_props

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
        if font_opts.key?(:font_size)
          warn "[Xlsxrb DEPRECATION] StyleBuilder font option :font_size is deprecated; use :size instead.", uplevel: 1
          font_size(font_opts[:font_size])
        elsif font_opts.key?(:size)
          font_size(font_opts[:size])
        elsif font_opts.key?(:sz)
          warn "[DEPRECATION] :sz is deprecated; use :font_size or :size instead.", uplevel: 1
          font_size(font_opts[:sz])
        end
        if font_opts.key?(:font_name)
          warn "[Xlsxrb DEPRECATION] StyleBuilder font option :font_name is deprecated; use :name instead.", uplevel: 1
          font_name(font_opts[:font_name])
        elsif font_opts.key?(:name)
          font_name(font_opts[:name])
        end
        if font_opts.key?(:font_color)
          warn "[Xlsxrb DEPRECATION] StyleBuilder font option :font_color is deprecated; use :color instead.", uplevel: 1
          font_color(font_opts[:font_color])
        elsif font_opts.key?(:color)
          font_color(font_opts[:color])
        end
        underline(font_opts[:underline]) if font_opts.key?(:underline)
        strike(font_opts[:strike]) if font_opts.key?(:strike)
        if font_opts.key?(:vertical_align)
          warn "[Xlsxrb DEPRECATION] StyleBuilder font option :vertical_align is deprecated; use :vert_align instead.", uplevel: 1
          vert_align(font_opts[:vertical_align])
        elsif font_opts.key?(:vert_align)
          vert_align(font_opts[:vert_align])
        end
        superscript(font_opts[:superscript]) if font_opts.key?(:superscript)
        subscript(font_opts[:subscript]) if font_opts.key?(:subscript)
      else
        bold(opts[:bold]) if opts.key?(:bold)
        italic(opts[:italic]) if opts.key?(:italic)
        if opts.key?(:size)
          warn "[Xlsxrb DEPRECATION] StyleBuilder option :size is deprecated; use :font_size instead.", uplevel: 1
          font_size(opts[:size])
        elsif opts.key?(:font_size)
          font_size(opts[:font_size])
        elsif opts.key?(:sz)
          warn "[DEPRECATION] :sz is deprecated; use :font_size or :size instead.", uplevel: 1
          font_size(opts[:sz])
        end
        if opts.key?(:name)
          warn "[Xlsxrb DEPRECATION] StyleBuilder option :name is deprecated; use :font_name instead.", uplevel: 1
          font_name(opts[:name])
        elsif opts.key?(:font_name)
          font_name(opts[:font_name])
        end
        if opts.key?(:color)
          warn "[Xlsxrb DEPRECATION] StyleBuilder option :color is deprecated; use :font_color instead.", uplevel: 1
          font_color(opts[:color])
        elsif opts.key?(:font_color)
          font_color(opts[:font_color])
        end
        underline(opts[:underline]) if opts.key?(:underline)
        strike(opts[:strike]) if opts.key?(:strike)
        if opts.key?(:vertical_align)
          warn "[Xlsxrb DEPRECATION] StyleBuilder option :vertical_align is deprecated; use :vert_align instead.", uplevel: 1
          vert_align(opts[:vertical_align])
        elsif opts.key?(:vert_align)
          vert_align(opts[:vert_align])
        end
        superscript(opts[:superscript]) if opts.key?(:superscript)
        subscript(opts[:subscript]) if opts.key?(:subscript)
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

      if opts.key?(:number_format_id)
        warn "[Xlsxrb DEPRECATION] StyleBuilder option :number_format_id is deprecated; use :number_format instead.", uplevel: 1
        number_format(opts[:number_format_id])
      elsif opts.key?(:number_format)
        number_format(opts[:number_format])
      end

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

      if opts.key?(:protection)
        prot_opts = opts[:protection] || {}
        locked(prot_opts[:locked]) if prot_opts.key?(:locked)
        hidden(prot_opts[:hidden]) if prot_opts.key?(:hidden)
      else
        locked(opts[:locked]) if opts.key?(:locked)
        hidden(opts[:hidden]) if opts.key?(:hidden)
      end

      self
    end

    # --- Font Properties ---

    # Configures multiple font properties at once.
    # @param opts [Hash] The font properties.
    # @return [self]
    # @api public
    #: (**untyped) -> (self | Hash[Symbol, untyped])
    def font(**opts)
      return @font_props if opts.empty?

      bold(opts[:bold]) if opts.key?(:bold)
      italic(opts[:italic]) if opts.key?(:italic)
      if opts.key?(:font_size)
        warn "[Xlsxrb DEPRECATION] StyleBuilder font option :font_size is deprecated; use :size instead.", uplevel: 1
        font_size(opts[:font_size])
      elsif opts.key?(:size)
        font_size(opts[:size])
      elsif opts.key?(:sz)
        warn "[DEPRECATION] :sz is deprecated; use :font_size or :size instead.", uplevel: 1
        font_size(opts[:sz])
      end
      if opts.key?(:font_name)
        warn "[Xlsxrb DEPRECATION] StyleBuilder font option :font_name is deprecated; use :name instead.", uplevel: 1
        font_name(opts[:font_name])
      elsif opts.key?(:name)
        font_name(opts[:name])
      end
      if opts.key?(:font_color)
        warn "[Xlsxrb DEPRECATION] StyleBuilder font option :font_color is deprecated; use :color instead.", uplevel: 1
        font_color(opts[:font_color])
      elsif opts.key?(:color)
        font_color(opts[:color])
      end
      underline(opts[:underline]) if opts.key?(:underline)
      strike(opts[:strike]) if opts.key?(:strike)
      if opts.key?(:vertical_align)
        warn "[Xlsxrb DEPRECATION] StyleBuilder font option :vertical_align is deprecated; use :vert_align instead.", uplevel: 1
        vert_align(opts[:vertical_align])
      elsif opts.key?(:vert_align)
        vert_align(opts[:vert_align])
      end
      superscript(opts[:superscript]) if opts.key?(:superscript)
      subscript(opts[:subscript]) if opts.key?(:subscript)
      self
    end

    def font=(opts)
      font(**opts)
    end

    # Sets the font to bold.
    # @param value [Boolean] Whether to apply bold.
    # @return [self]
    # @api public
    #: (?untyped) -> (self | bool)
    def bold(value = OMITTED)
      if value.equal?(OMITTED)
        return bold? unless @font_props[:bold].nil?

        Xlsxrb::Deprecation.warn("Calling StyleBuilder#bold with zero arguments as a mutator is deprecated; pass true explicitly (e.g. bold(true))", uplevel: 1)
        @font_props[:bold] = true
        return self
      end

      @font_props[:bold] = value ? true : false
      self
    end

    # Sets the font to bold.
    # @param value [Boolean]
    # @return [Boolean]
    # @api public
    #: (bool value) -> bool
    def bold=(value)
      @font_props[:bold] = value
    end

    # Returns whether bold is enabled.
    # @return [Boolean]
    # @api public
    #: () -> bool
    def bold?
      @font_props[:bold] ? true : false
    end

    # Sets the font to italic.
    # @param value [Boolean] Whether to apply italic.
    # @return [self]
    # @api public
    #: (?untyped) -> (self | bool)
    def italic(value = OMITTED)
      if value.equal?(OMITTED)
        return italic? unless @font_props[:italic].nil?

        Xlsxrb::Deprecation.warn("Calling StyleBuilder#italic with zero arguments as a mutator is deprecated; pass true explicitly (e.g. italic(true))", uplevel: 1)
        @font_props[:italic] = true
        return self
      end

      @font_props[:italic] = value ? true : false
      self
    end

    # Sets the font to italic.
    # @param value [Boolean]
    # @return [Boolean]
    # @api public
    #: (bool value) -> bool
    def italic=(value)
      @font_props[:italic] = value
    end

    # Returns whether italic is enabled.
    # @return [Boolean]
    # @api public
    #: () -> bool
    def italic?
      @font_props[:italic] ? true : false
    end

    # Sets or gets the font size in points.
    # @param size_value [Integer, Float, nil] The size in points (supports fractional sizes such as 10.5).
    # @return [Integer, Float, self, nil]
    # @api public
    #: (?Numeric? size_value) -> (Numeric | self | nil)
    def font_size(size_value = nil)
      if size_value.nil?
        @font_props[:size] || @font_props[:sz]
      else
        val = size_value.is_a?(Float) && (size_value % 1).zero? ? size_value.to_i : size_value
        @font_props[:size] = val
        @font_props[:sz] = val
        self
      end
    end

    # Sets the font size in points.
    # @param size_value [Integer, Float] The size in points (supports fractional sizes such as 10.5).
    # @return [Integer, Float]
    # @api public
    #: (Numeric size_value) -> Numeric
    def font_size=(size_value)
      val = size_value.is_a?(Float) && (size_value % 1).zero? ? size_value.to_i : size_value
      @font_props[:size] = val
      @font_props[:sz] = val
      size_value
    end

    # @deprecated Use {#font_size} instead. Will be removed in 0.2.0.
    #: (?Numeric? size_value) -> (Numeric | self | nil)
    def size(size_value = nil)
      warn "[DEPRECATION] #size is deprecated; use #font_size instead.", uplevel: 1
      font_size(size_value)
    end

    # @deprecated Use {#font_size=} instead. Will be removed in 0.2.0.
    #: (Numeric size_value) -> Numeric
    def size=(size_value)
      warn "[DEPRECATION] #size= is deprecated; use #font_size= instead.", uplevel: 1
      self.font_size = size_value
    end

    # @deprecated Use {#font_size} instead. Will be removed in 0.2.0.
    #: (?Numeric? size_value) -> (Numeric | self | nil)
    def sz(size_value = nil)
      warn "[DEPRECATION] #sz is deprecated; use #font_size instead.", uplevel: 1
      font_size(size_value)
    end

    # @deprecated Use {#font_size=} instead. Will be removed in 0.2.0.
    #: (Numeric size_value) -> Numeric
    def sz=(size_value)
      warn "[DEPRECATION] #sz= is deprecated; use #font_size= instead.", uplevel: 1
      self.font_size = size_value
    end

    # Sets the font name.
    # Gets or sets the font name.
    # @param name [String, Object] The font name.
    # @return [String, self, nil]
    # @api public
    #: (?untyped name) -> (String | self | nil)
    def font_name(name = OMITTED)
      return @font_props[:name] if name.equal?(OMITTED)

      @font_props[:name] = name
      self
    end

    # Sets the font name.
    # @param name [String]
    # @return [String]
    # @api public
    #: (String name) -> String
    def font_name=(name)
      @font_props[:name] = name
    end

    # Gets or sets the font color.
    # @param color [String, Symbol, Object] The color.
    # @return [String, Symbol, self, nil]
    # @api public
    #: (?untyped color) -> (String | Symbol | self | nil)
    def font_color(color = OMITTED)
      return @font_props[:color] if color.equal?(OMITTED)

      @font_props[:color] = resolve_color(color)
      self
    end

    # Sets the font color.
    # @param color [String, Symbol]
    # @return [String, Symbol]
    # @api public
    #: (String | Symbol color) -> (String | Symbol)
    def font_color=(color)
      @font_props[:color] = resolve_color(color)
      color
    end

    # Sets the font underline style.
    # @param val [String, Symbol, Boolean] The underline style (e.g., 'single', :double, :single_accounting, :double_accounting).
    # @return [self]
    # @api public
    #: (?untyped) -> (self | String | Symbol | bool | nil)
    def underline(val = OMITTED)
      if val.equal?(OMITTED)
        return @font_props[:underline] unless @font_props[:underline].nil?

        Xlsxrb::Deprecation.warn("Calling StyleBuilder#underline with zero arguments as a mutator is deprecated; pass true explicitly (e.g. underline(true))", uplevel: 1)
        @font_props[:underline] = "single"
        return self
      end

      mapped = if [true, "single"].include?(val)
                 "single"
               elsif val == false
                 nil
               else
                 UNDERLINE_MAP[val.to_sym] || val.to_s
               end
      @font_props[:underline] = mapped
      self
    end

    # Sets the font underline style.
    # @param val [String, Symbol, Boolean]
    # @return [String, Symbol, Boolean]
    # @api public
    #: (String | Symbol | bool val) -> (String | Symbol | bool)
    def underline=(val)
      underline(val)
      val
    end

    # Returns whether underline is enabled.
    # @return [Boolean]
    # @api public
    #: () -> bool
    def underline?
      !@font_props[:underline].nil? && @font_props[:underline] != false
    end

    # Sets the font strikethrough.
    # @param value [Boolean] Whether to apply strike.
    # @return [self]
    # @api public
    #: (?untyped) -> (self | bool)
    def strike(value = OMITTED)
      if value.equal?(OMITTED)
        return strike? unless @font_props[:strike].nil?

        Xlsxrb::Deprecation.warn("Calling StyleBuilder#strike with zero arguments as a mutator is deprecated; pass true explicitly (e.g. strike(true))", uplevel: 1)
        @font_props[:strike] = true
        return self
      end

      @font_props[:strike] = value ? true : false
      self
    end

    # Sets the font strikethrough.
    # @param value [Boolean]
    # @return [Boolean]
    # @api public
    #: (bool value) -> bool
    def strike=(value)
      @font_props[:strike] = value
    end

    # Returns whether strikethrough is enabled.
    # @return [Boolean]
    # @api public
    #: () -> bool
    def strike?
      @font_props[:strike] ? true : false
    end

    # Gets or sets the vertical alignment of the font.
    # @param value [String, Symbol, Object] The alignment value ('superscript', 'subscript', 'baseline', :superscript, :subscript).
    # @return [String, self, nil]
    # @api public
    #: (?untyped value) -> (String | self | nil)
    def vert_align(value = OMITTED)
      return @font_props[:vert_align] if value.equal?(OMITTED)

      mapped = VERT_ALIGN_MAP[value.to_sym] || value.to_s
      @font_props[:vert_align] = mapped
      self
    end

    def vert_align=(value)
      vert_align(value)
      value
    end

    # Sets the font to superscript.
    # @param value [Boolean] Whether to apply superscript.
    # @return [self]
    # @api public
    #: (?untyped) -> (self | bool)
    def superscript(value = OMITTED)
      if value.equal?(OMITTED)
        return superscript? unless @font_props[:vert_align].nil?

        Xlsxrb::Deprecation.warn("Calling StyleBuilder#superscript with zero arguments as a mutator is deprecated; pass true explicitly (e.g. superscript(true))", uplevel: 1)
        @font_props[:vert_align] = "superscript"
        return self
      end

      @font_props[:vert_align] = value ? "superscript" : "baseline"
      self
    end

    def superscript=(value)
      superscript(value)
      value
    end

    def superscript?
      @font_props[:vert_align] == "superscript"
    end

    # Sets the font to subscript.
    # @param value [Boolean] Whether to apply subscript.
    # @return [self]
    # @api public
    #: (?untyped) -> (self | bool)
    def subscript(value = OMITTED)
      if value.equal?(OMITTED)
        return subscript? unless @font_props[:vert_align].nil?

        Xlsxrb::Deprecation.warn("Calling StyleBuilder#subscript with zero arguments as a mutator is deprecated; pass true explicitly (e.g. subscript(true))", uplevel: 1)
        @font_props[:vert_align] = "subscript"
        return self
      end

      @font_props[:vert_align] = value ? "subscript" : "baseline"
      self
    end

    def subscript=(value)
      subscript(value)
      value
    end

    def subscript?
      @font_props[:vert_align] == "subscript"
    end
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
    #: (?untyped pattern, ?fg_color: String | Symbol | nil, ?bg_color: String | Symbol | nil) -> untyped
    def fill_pattern(pattern = OMITTED, fg_color: nil, bg_color: nil)
      return @fill_props[:pattern] if pattern.equal?(OMITTED)

      @fill_props[:pattern] = pattern
      @fill_props[:fg_color] = fg_color if fg_color
      @fill_props[:bg_color] = bg_color if bg_color
      self
    end

    # Sets the fill pattern.
    # @param pattern [String, Symbol, Hash]
    # @return [Object]
    # @api public
    #: (untyped pattern) -> untyped
    def fill_pattern=(pattern)
      if pattern.is_a?(Hash)
        fill_pattern(pattern[:pattern] || "solid", fg_color: pattern[:fg_color], bg_color: pattern[:bg_color])
      else
        @fill_props[:pattern] = pattern
      end
      pattern
    end

    # Gets or sets a solid fill color.
    # @param color [String, Symbol, Object] The color.
    # @return [String, Symbol, self, nil]
    # @api public
    #: (?untyped color) -> (String | Symbol | self | nil)
    def fill_color(color = OMITTED)
      return @fill_props[:fg_color] if color.equal?(OMITTED)

      @fill_props[:pattern] = "solid"
      @fill_props[:fg_color] = resolve_color(color)
      self
    end

    # Sets a solid fill color.
    # @param color [String, Symbol]
    # @return [String, Symbol]
    # @api public
    #: (String | Symbol color) -> (String | Symbol)
    def fill_color=(color)
      fill_color(color)
      color
    end

    #: (?pattern: untyped, ?fg_color: String | Symbol | nil, ?bg_color: String | Symbol | nil) -> (self | Hash[Symbol, untyped])
    def fill(pattern: OMITTED, fg_color: nil, bg_color: nil)
      return @fill_props if pattern.equal?(OMITTED) && fg_color.nil? && bg_color.nil?

      pat = pattern.equal?(OMITTED) ? "solid" : pattern
      fill_pattern(pat, fg_color: resolve_color(fg_color), bg_color: resolve_color(bg_color))
    end

    def fill=(opts)
      fill(**opts)
    end

    # Sets a gradient fill.
    # @param type [String] The gradient type.
    # @param degree [Numeric, nil] The degree.
    # @param stops [Array] The gradient stops.
    # @return [self]
    # @api public
    #: (?type: untyped, ?degree: Numeric | nil, ?stops: Array[untyped]) -> (self | Hash[Symbol, untyped] | nil)
    def fill_gradient(type: OMITTED, degree: nil, stops: [])
      return @fill_props[:gradient] if type.equal?(OMITTED)

      @fill_props[:gradient] = {
        type: type,
        degree: degree,
        stops: stops
      }.compact
      self
    end

    # --- Border Properties ---

    # Configures multiple border properties.
    # @param all [Hash, nil]
    # @param left [Hash, nil]
    # @param right [Hash, nil]
    # @param top [Hash, nil]
    # @param bottom [Hash, nil]
    # @param diagonal [Hash, nil]
    # @return [self]
    # @api public
    #: (?all: Hash[Symbol, untyped]?, ?left: Hash[Symbol, untyped]?, ?right: Hash[Symbol, untyped]?, ?top: Hash[Symbol, untyped]?, ?bottom: Hash[Symbol, untyped]?, ?diagonal: Hash[Symbol, untyped]?) -> (self | Hash[Symbol, untyped])
    def border(all: nil, left: nil, right: nil, top: nil, bottom: nil, diagonal: nil)
      return @border_props if all.nil? && left.nil? && right.nil? && top.nil? && bottom.nil? && diagonal.nil?

      border_all(**all) if all
      border_left(**left) if left
      border_right(**right) if right
      border_top(**top) if top
      border_bottom(**bottom) if bottom
      border_diagonal(**diagonal) if diagonal
      self
    end

    def border=(opts)
      border(**opts)
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

    # Configures or returns alignment properties.
    # When called with 0 arguments/keywords, returns the current alignment Hash.
    # When called with keyword arguments, merges the properties and returns self.
    #
    # @param horizontal [String, Symbol, nil]
    # @param vertical [String, Symbol, nil]
    # @param wrap_text [Boolean, nil]
    # @param text_rotation [Integer, nil]
    # @param indent [Integer, nil]
    # @param shrink_to_fit [Boolean, nil]
    # @param reading_order [Integer, nil]
    # @param justify_last_line [Boolean, nil]
    # @param relative_indent [Integer, nil]
    # @return [Hash, self]
    # @api public
    #: (?horizontal: (String | Symbol)?, ?vertical: (String | Symbol)?, ?wrap_text: bool?, ?text_rotation: Integer?, ?indent: Integer?, ?shrink_to_fit: bool?, ?reading_order: Integer?, ?justify_last_line: bool?, ?relative_indent: Integer?) -> (Hash[Symbol, untyped] | self)
    def alignment(horizontal: nil, vertical: nil, wrap_text: nil, text_rotation: nil, indent: nil, shrink_to_fit: nil, reading_order: nil, justify_last_line: nil, relative_indent: nil)
      return @alignment if horizontal.nil? && vertical.nil? && wrap_text.nil? && text_rotation.nil? && indent.nil? && shrink_to_fit.nil? && reading_order.nil? && justify_last_line.nil? && relative_indent.nil?

      align_horizontal(horizontal) unless horizontal.nil?
      align_vertical(vertical) unless vertical.nil?
      wrap_text(wrap_text) unless wrap_text.nil?
      text_rotation(text_rotation) unless text_rotation.nil?
      indent(indent) unless indent.nil?
      shrink_to_fit(shrink_to_fit) unless shrink_to_fit.nil?
      @alignment[:reading_order] = reading_order unless reading_order.nil?
      @alignment[:justify_last_line] = justify_last_line unless justify_last_line.nil?
      @alignment[:relative_indent] = relative_indent unless relative_indent.nil?
      self
    end

    # Sets alignment options.
    #
    # @param opts [Hash]
    # @return [Hash]
    # @api public
    #: (Hash[Symbol, untyped] opts) -> Hash[Symbol, untyped]
    def alignment=(opts)
      alignment(**opts)
      opts
    end

    # Gets or sets horizontal alignment.
    # @param value [String, Symbol, Object] The alignment.
    # @return [String, Symbol, self, nil]
    # @api public
    #: (?untyped value) -> (String | Symbol | self | nil)
    def align_horizontal(value = OMITTED)
      return @alignment[:horizontal] if value.equal?(OMITTED)

      @alignment[:horizontal] = value
      self
    end

    # Sets horizontal alignment.
    # @param value [String, Symbol]
    # @return [String, Symbol]
    # @api public
    #: (String | Symbol value) -> (String | Symbol)
    def align_horizontal=(value)
      @alignment[:horizontal] = value
    end

    # Gets or sets vertical alignment.
    # @param value [String, Symbol, Object] The alignment.
    # @return [String, Symbol, self, nil]
    # @api public
    #: (?untyped value) -> (String | Symbol | self | nil)
    def align_vertical(value = OMITTED)
      return @alignment[:vertical] if value.equal?(OMITTED)

      @alignment[:vertical] = value
      self
    end

    # Sets vertical alignment.
    # @param value [String, Symbol]
    # @return [String, Symbol]
    # @api public
    #: (String | Symbol value) -> (String | Symbol)
    def align_vertical=(value)
      @alignment[:vertical] = value
    end

    # Sets text wrapping.
    # @param value [Boolean] Whether to wrap text.
    # @return [self]
    # @api public
    #: (?untyped) -> (self | bool)
    def wrap_text(value = OMITTED)
      if value.equal?(OMITTED)
        return wrap_text? unless @alignment[:wrap_text].nil?

        Xlsxrb::Deprecation.warn("Calling StyleBuilder#wrap_text with zero arguments as a mutator is deprecated; pass true explicitly (e.g. wrap_text(true))", uplevel: 1)
        @alignment[:wrap_text] = true
        return self
      end

      @alignment[:wrap_text] = value ? true : false
      self
    end

    # Sets text wrapping.
    # @param value [Boolean]
    # @return [Boolean]
    # @api public
    #: (bool value) -> bool
    def wrap_text=(value)
      @alignment[:wrap_text] = value
    end

    # Returns whether text wrapping is enabled.
    # @return [Boolean]
    # @api public
    #: () -> bool
    def wrap_text?
      @alignment[:wrap_text] ? true : false
    end

    # Sets shrink to fit.
    # @param value [Boolean] Whether to shrink text.
    # @return [self]
    # @api public
    #: (?untyped) -> (self | bool)
    def shrink_to_fit(value = OMITTED)
      if value.equal?(OMITTED)
        return shrink_to_fit? unless @alignment[:shrink_to_fit].nil?

        Xlsxrb::Deprecation.warn("Calling StyleBuilder#shrink_to_fit with zero arguments as a mutator is deprecated; pass true explicitly (e.g. shrink_to_fit(true))", uplevel: 1)
        @alignment[:shrink_to_fit] = true
        return self
      end

      @alignment[:shrink_to_fit] = value ? true : false
      self
    end

    # Sets shrink to fit.
    # @param value [Boolean]
    # @return [Boolean]
    # @api public
    #: (bool value) -> bool
    def shrink_to_fit=(value)
      @alignment[:shrink_to_fit] = value
    end

    # Returns whether shrink to fit is enabled.
    # @return [Boolean]
    # @api public
    #: () -> bool
    def shrink_to_fit?
      @alignment[:shrink_to_fit] ? true : false
    end

    # Gets or sets text rotation.
    # @param value [Numeric, Object] The rotation angle.
    # @return [Numeric, self, nil]
    # @api public
    #: (?untyped value) -> (Numeric | self | nil)
    def text_rotation(value = OMITTED)
      return @alignment[:text_rotation] if value.equal?(OMITTED)

      @alignment[:text_rotation] = value
      self
    end

    # Sets text rotation.
    # @param value [Numeric]
    # @return [Numeric]
    # @api public
    #: (Numeric value) -> Numeric
    def text_rotation=(value)
      @alignment[:text_rotation] = value
    end

    # Gets or sets text indent.
    # @param value [Numeric, Object] The indent level.
    # @return [Numeric, self, nil]
    # @api public
    #: (?untyped value) -> (Numeric | self | nil)
    def indent(value = OMITTED)
      return @alignment[:indent] if value.equal?(OMITTED)

      @alignment[:indent] = value.to_i
      self
    end

    # Sets text indent.
    # @param value [Numeric]
    # @return [Numeric]
    # @api public
    #: (Numeric value) -> Numeric
    def indent=(value)
      @alignment[:indent] = value.to_i
    end

    # --- Protection Properties ---

    # Configures or returns cell protection properties.
    # When called with 0 arguments/keywords, returns the current protection Hash.
    # When called with keyword arguments, merges the properties and returns self.
    #
    # @param locked [Boolean, nil]
    # @param hidden [Boolean, nil]
    # @return [Hash, self]
    # @api public
    #: (?locked: bool?, ?hidden: bool?) -> (Hash[Symbol, untyped] | self)
    def protection(locked: nil, hidden: nil)
      return @protection if locked.nil? && hidden.nil?

      @protection[:locked] = locked unless locked.nil?
      @protection[:hidden] = hidden unless hidden.nil?
      self
    end

    # Sets protection options.
    #
    # @param opts [Hash]
    # @return [Hash]
    # @api public
    #: (Hash[Symbol, untyped] opts) -> Hash[Symbol, untyped]
    def protection=(opts)
      protection(**opts)
      opts
    end

    # Sets whether the cell is locked.
    #
    # @param value [Boolean]
    # @return [self]
    # @api public
    #: (?untyped value) -> (self | bool)
    def locked(value = OMITTED)
      if value.equal?(OMITTED)
        return locked? unless @protection[:locked].nil?

        Xlsxrb::Deprecation.warn("Calling StyleBuilder#locked with zero arguments as a mutator is deprecated; pass true explicitly (e.g. locked(true))", uplevel: 1)
        @protection[:locked] = true
        return self
      end

      @protection[:locked] = value ? true : false
      self
    end

    # Sets whether the cell is locked.
    #
    # @param value [Boolean]
    # @return [Boolean]
    # @api public
    #: (bool value) -> bool
    def locked=(value)
      @protection[:locked] = value
    end

    # Returns whether the cell is locked.
    # @return [Boolean]
    # @api public
    #: () -> bool
    def locked?
      @protection[:locked] != false
    end

    # Sets whether the cell formula is hidden.
    #
    # @param value [Boolean]
    # @return [self]
    # @api public
    #: (?untyped value) -> (self | bool)
    def hidden(value = OMITTED)
      if value.equal?(OMITTED)
        return hidden? unless @protection[:hidden].nil?

        Xlsxrb::Deprecation.warn("Calling StyleBuilder#hidden with zero arguments as a mutator is deprecated; pass true explicitly (e.g. hidden(true))", uplevel: 1)
        @protection[:hidden] = true
        return self
      end

      @protection[:hidden] = value ? true : false
      self
    end

    # Sets whether the cell formula is hidden.
    #
    # @param value [Boolean]
    # @return [Boolean]
    # @api public
    #: (bool value) -> bool
    def hidden=(value)
      @protection[:hidden] = value
    end

    # Returns whether the cell formula is hidden.
    # @return [Boolean]
    # @api public
    #: () -> bool
    def hidden?
      @protection[:hidden] ? true : false
    end

    # --- Number Format ---

    # Sets or gets the number format.
    # @param num_fmt_id [String, Integer, nil] The format id or format string.
    # @return [String, Integer, self, nil]
    # @api public
    #: (?String | Integer | nil num_fmt_id) -> (String | Integer | self | nil)
    def number_format(num_fmt_id = nil)
      if num_fmt_id.nil?
        @num_fmt_id
      else
        @num_fmt_id = num_fmt_id
        self
      end
    end

    # Sets the number format.
    # @param num_fmt_id [String, Integer]
    # @return [String, Integer]
    # @api public
    #: (String | Integer num_fmt_id) -> (String | Integer)
    def number_format=(num_fmt_id)
      @num_fmt_id = num_fmt_id
    end

    # Sets or gets the number format ID.
    # @deprecated Use {#number_format} instead. Will be removed in 0.2.0.
    # @param id [String, Integer, nil]
    # @return [String, Integer, self, nil]
    # @api public
    #: (?String | Integer | nil id) -> (String | Integer | self | nil)
    def number_format_id(id = nil)
      warn "[Xlsxrb DEPRECATION] StyleBuilder#number_format_id is deprecated; use #number_format instead.", uplevel: 1
      number_format(id)
    end

    # Sets the number format ID.
    # @deprecated Use {#number_format=} instead. Will be removed in 0.2.0.
    # @param id [String, Integer]
    # @return [String, Integer]
    # @api public
    #: (String | Integer id) -> (String | Integer)
    def number_format_id=(id)
      warn "[Xlsxrb DEPRECATION] StyleBuilder#number_format_id= is deprecated; use #number_format= instead.", uplevel: 1
      @num_fmt_id = id
    end

    # @deprecated Use {#number_format} instead. Will be removed in 0.2.0.
    #: (?String | Integer | nil num_fmt_id) -> (String | Integer | self | nil)
    def num_fmt(num_fmt_id = nil)
      warn "[DEPRECATION] #num_fmt is deprecated; use #number_format instead.", uplevel: 1
      number_format(num_fmt_id)
    end

    # @deprecated Use {#number_format=} instead. Will be removed in 0.2.0.
    #: (String | Integer num_fmt_id) -> (String | Integer)
    def num_fmt=(num_fmt_id)
      warn "[DEPRECATION] #num_fmt= is deprecated; use #number_format= instead.", uplevel: 1
      self.number_format = num_fmt_id
    end

    # @deprecated Use {#number_format_id} instead. Will be removed in 0.2.0.
    #: (?String | Integer | nil id) -> (String | Integer | self | nil)
    def num_fmt_id(id = nil)
      warn "[DEPRECATION] #num_fmt_id is deprecated; use #number_format_id instead.", uplevel: 1
      number_format_id(id)
    end

    # @deprecated Use {#number_format_id=} instead. Will be removed in 0.2.0.
    #: (String | Integer id) -> (String | Integer)
    def num_fmt_id=(id)
      warn "[DEPRECATION] #num_fmt_id= is deprecated; use #number_format_id= instead.", uplevel: 1
      self.number_format_id = id
    end

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
      cell_style_opts[:protection] = @protection if @protection.any?

      writer.add_cell_style(**cell_style_opts)
    end

    private

    def resolve_color(color)
      self.class.resolve_color(color)
    end
  end
end
