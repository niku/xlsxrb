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
        font_size(font_opts[:size]) if font_opts.key?(:size)
        font_name(font_opts[:name]) if font_opts.key?(:name)
        font_color(font_opts[:color]) if font_opts.key?(:color)
        underline(font_opts[:underline]) if font_opts.key?(:underline)
        strike(font_opts[:strike]) if font_opts.key?(:strike)
        vert_align(font_opts[:vert_align]) if font_opts.key?(:vert_align)
        superscript(font_opts[:superscript]) if font_opts.key?(:superscript)
        subscript(font_opts[:subscript]) if font_opts.key?(:subscript)
      else
        bold(opts[:bold]) if opts.key?(:bold)
        italic(opts[:italic]) if opts.key?(:italic)
        font_size(opts[:font_size]) if opts.key?(:font_size)
        font_name(opts[:font_name]) if opts.key?(:font_name)
        font_color(opts[:font_color]) if opts.key?(:font_color)
        underline(opts[:underline]) if opts.key?(:underline)
        strike(opts[:strike]) if opts.key?(:strike)
        vert_align(opts[:vert_align]) if opts.key?(:vert_align)
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

    # Configures multiple font properties at once or returns font properties.
    # @param opts [Hash] The font properties.
    # @return [Hash, void]
    # @api public
    #: (**untyped) -> (Hash[Symbol, untyped] | void)
    def font(**opts)
      return @font_props if opts.empty?

      bold(opts[:bold]) if opts.key?(:bold)
      italic(opts[:italic]) if opts.key?(:italic)
      font_size(opts[:size]) if opts.key?(:size)
      font_name(opts[:name]) if opts.key?(:name)
      font_color(opts[:color]) if opts.key?(:color)
      underline(opts[:underline]) if opts.key?(:underline)
      strike(opts[:strike]) if opts.key?(:strike)
      vert_align(opts[:vert_align]) if opts.key?(:vert_align)
      superscript(opts[:superscript]) if opts.key?(:superscript)
      subscript(opts[:subscript]) if opts.key?(:subscript)
      nil
    end

    # Sets font options.
    # @param opts [Hash]
    # @return [Hash]
    # @api public
    #: (Hash[Symbol, untyped] opts) -> Hash[Symbol, untyped]
    def font=(opts)
      font(**opts)
      opts
    end

    # Sets or returns whether font is bold.
    # @param value [Boolean, Object]
    # @return [Boolean, void]
    # @api public
    #: (?untyped value) -> (bool | void)
    def bold(value = OMITTED)
      return bold? if value.equal?(OMITTED)

      @font_props[:bold] = value ? true : false
      nil
    end

    # Sets the font to bold.
    # @param value [Boolean]
    # @return [Boolean]
    # @api public
    #: (bool value) -> bool
    def bold=(value)
      @font_props[:bold] = value ? true : false
    end

    # Returns whether bold is enabled.
    # @return [Boolean]
    # @api public
    #: () -> bool
    def bold?
      @font_props[:bold] ? true : false
    end

    # Sets or returns whether font is italic.
    # @param value [Boolean, Object]
    # @return [Boolean, void]
    # @api public
    #: (?untyped value) -> (bool | void)
    def italic(value = OMITTED)
      return italic? if value.equal?(OMITTED)

      @font_props[:italic] = value ? true : false
      nil
    end

    # Sets the font to italic.
    # @param value [Boolean]
    # @return [Boolean]
    # @api public
    #: (bool value) -> bool
    def italic=(value)
      @font_props[:italic] = value ? true : false
    end

    # Returns whether italic is enabled.
    # @return [Boolean]
    # @api public
    #: () -> bool
    def italic?
      @font_props[:italic] ? true : false
    end

    # Sets or gets the font size in points.
    # @param size_value [Integer, Float, Object, nil] The size in points (supports fractional sizes such as 10.5).
    # @return [Integer, Float, void, nil]
    # @api public
    #: (?Numeric? size_value) -> (Numeric | void | nil)
    def font_size(size_value = OMITTED)
      return @font_props[:size] if size_value.equal?(OMITTED) || size_value.nil?

      val = size_value.is_a?(Float) && (size_value % 1).zero? ? size_value.to_i : size_value
      @font_props[:size] = val
      nil
    end

    # Sets the font size in points.
    # @param size_value [Integer, Float] The size in points (supports fractional sizes such as 10.5).
    # @return [Integer, Float]
    # @api public
    #: (Numeric size_value) -> Numeric
    def font_size=(size_value)
      val = size_value.is_a?(Float) && (size_value % 1).zero? ? size_value.to_i : size_value
      @font_props[:size] = val
      size_value
    end

    # Gets or sets the font name.
    # @param name [String, Object] The font name.
    # @return [String, void, nil]
    # @api public
    #: (?untyped name) -> (String | void | nil)
    def font_name(name = OMITTED)
      return @font_props[:name] if name.equal?(OMITTED)

      @font_props[:name] = name
      nil
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
    # @return [String, Symbol, void, nil]
    # @api public
    #: (?untyped color) -> (String | Symbol | void | nil)
    def font_color(color = OMITTED)
      return @font_props[:color] if color.equal?(OMITTED)

      @font_props[:color] = resolve_color(color)
      nil
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

    # Sets or gets the font underline style.
    # @param val [String, Symbol, Boolean, Object] The underline style (e.g., 'single', :double, :single_accounting, :double_accounting).
    # @return [String, Symbol, Boolean, void, nil]
    # @api public
    #: (?untyped val) -> (String | Symbol | bool | void | nil)
    def underline(val = OMITTED)
      return @font_props[:underline] if val.equal?(OMITTED)

      mapped = if [true, "single"].include?(val)
                 "single"
               elsif val == false
                 nil
               else
                 UNDERLINE_MAP[val.to_sym] || val.to_s
               end
      @font_props[:underline] = mapped
      nil
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

    # Sets or returns whether font has strikethrough.
    # @param value [Boolean, Object]
    # @return [Boolean, void]
    # @api public
    #: (?untyped value) -> (bool | void)
    def strike(value = OMITTED)
      return strike? if value.equal?(OMITTED)

      @font_props[:strike] = value ? true : false
      nil
    end

    # Sets the font strikethrough.
    # @param value [Boolean]
    # @return [Boolean]
    # @api public
    #: (bool value) -> bool
    def strike=(value)
      @font_props[:strike] = value ? true : false
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
    # @return [String, void, nil]
    # @api public
    #: (?untyped value) -> (String | void | nil)
    def vert_align(value = OMITTED)
      return @font_props[:vert_align] if value.equal?(OMITTED)

      mapped = VERT_ALIGN_MAP[value.to_sym] || value.to_s
      @font_props[:vert_align] = mapped
      nil
    end

    # Sets the vertical alignment of the font.
    # @param val [String, Symbol]
    # @return [Object]
    # @api public
    #: (untyped val) -> untyped
    def vert_align=(val)
      vert_align(val)
      val
    end

    # Sets or gets whether the font is superscript.
    # @param value [Boolean, Object] Whether to apply superscript.
    # @return [Boolean, void]
    # @api public
    #: (?untyped value) -> (bool | void)
    def superscript(value = OMITTED)
      return superscript? if value.equal?(OMITTED)

      @font_props[:vert_align] = value ? "superscript" : "baseline"
      nil
    end

    # Returns whether superscript is enabled.
    # @return [Boolean]
    # @api public
    #: () -> bool
    def superscript?
      @font_props[:vert_align] == "superscript"
    end

    # Sets the font to superscript.
    # @param val [Boolean]
    # @return [Boolean]
    # @api public
    #: (untyped val) -> untyped
    def superscript=(val)
      superscript(val)
      val
    end

    # Sets or gets whether the font is subscript.
    # @param value [Boolean, Object] Whether to apply subscript.
    # @return [Boolean, void]
    # @api public
    #: (?untyped value) -> (bool | void)
    def subscript(value = OMITTED)
      return subscript? if value.equal?(OMITTED)

      @font_props[:vert_align] = value ? "subscript" : "baseline"
      nil
    end

    # Returns whether subscript is enabled.
    # @return [Boolean]
    # @api public
    #: () -> bool
    def subscript?
      @font_props[:vert_align] == "subscript"
    end

    # Sets the font to subscript.
    # @param val [Boolean]
    # @return [Boolean]
    # @api public
    #: (untyped val) -> untyped
    def subscript=(val)
      subscript(val)
      val
    end

    # --- Fill Properties ---

    # Sets or gets the fill pattern.
    # @param pattern [String, Symbol, Object] The pattern type.
    # @param fg_color [String, Symbol, nil] The foreground color.
    # @param bg_color [String, Symbol, nil] The background color.
    # @return [String, Symbol, void, nil]
    # @api public
    #: (?untyped pattern, ?fg_color: String | Symbol | nil, ?bg_color: String | Symbol | nil) -> (untyped | void)
    def fill_pattern(pattern = OMITTED, fg_color: nil, bg_color: nil)
      return @fill_props[:pattern] if pattern.equal?(OMITTED)

      @fill_props[:pattern] = pattern
      @fill_props[:fg_color] = fg_color if fg_color
      @fill_props[:bg_color] = bg_color if bg_color
      nil
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
    # @return [String, Symbol, void, nil]
    # @api public
    #: (?untyped color) -> (String | Symbol | void | nil)
    def fill_color(color = OMITTED)
      return @fill_props[:fg_color] if color.equal?(OMITTED)

      @fill_props[:pattern] = "solid"
      @fill_props[:fg_color] = resolve_color(color)
      nil
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

    # Configures fill properties or returns current fill properties.
    # @param pattern [String, Symbol, Object]
    # @param fg_color [String, Symbol, nil]
    # @param bg_color [String, Symbol, nil]
    # @return [Hash, void]
    # @api public
    #: (?pattern: untyped, ?fg_color: String | Symbol | nil, ?bg_color: String | Symbol | nil) -> (Hash[Symbol, untyped] | void)
    def fill(pattern: OMITTED, fg_color: nil, bg_color: nil)
      return @fill_props if pattern.equal?(OMITTED) && fg_color.nil? && bg_color.nil?

      pat = pattern.equal?(OMITTED) ? "solid" : pattern
      fill_pattern(pat, fg_color: resolve_color(fg_color), bg_color: resolve_color(bg_color))
      nil
    end

    # Sets fill options.
    # @param opts [Hash, Object]
    # @return [Object]
    # @api public
    #: (untyped opts) -> untyped
    def fill=(opts)
      if opts.is_a?(Hash)
        fill(**opts)
      else
        fill(pattern: opts)
      end
      opts
    end

    # Sets or gets a gradient fill.
    # @param type [String, Object] The gradient type.
    # @param degree [Numeric, nil] The degree.
    # @param stops [Array] The gradient stops.
    # @return [Hash, void, nil]
    # @api public
    #: (?type: untyped, ?degree: Numeric | nil, ?stops: Array[untyped]) -> (Hash[Symbol, untyped] | void | nil)
    def fill_gradient(type: OMITTED, degree: nil, stops: [])
      return @fill_props[:gradient] if type.equal?(OMITTED)

      @fill_props[:gradient] = {
        type: type,
        degree: degree,
        stops: stops
      }.compact
      nil
    end

    # --- Border Properties ---

    # Configures multiple border properties or returns current border properties.
    #
    # @param all [Hash, nil] Border options for all sides.
    # @param left [Hash, nil] Options for left border.
    # @param right [Hash, nil] Options for right border.
    # @param top [Hash, nil] Options for top border.
    # @param bottom [Hash, nil] Options for bottom border.
    # @param diagonal [Hash, nil] Options for diagonal border.
    # @return [Hash, void]
    # @api public
    #: (?all: Hash[Symbol, untyped]?, ?left: Hash[Symbol, untyped]?, ?right: Hash[Symbol, untyped]?, ?top: Hash[Symbol, untyped]?, ?bottom: Hash[Symbol, untyped]?, ?diagonal: Hash[Symbol, untyped]?) -> (Hash[Symbol, untyped] | void)
    def border(all: nil, left: nil, right: nil, top: nil, bottom: nil, diagonal: nil)
      return @border_props if all.nil? && left.nil? && right.nil? && top.nil? && bottom.nil? && diagonal.nil?

      border_all(**all) if all
      border_left(**left) if left
      border_right(**right) if right
      border_top(**top) if top
      border_bottom(**bottom) if bottom
      border_diagonal(**diagonal) if diagonal
      nil
    end

    # Sets border options.
    # @param opts [Hash]
    # @return [Hash]
    # @api public
    #: (Hash[Symbol, untyped] opts) -> Hash[Symbol, untyped]
    def border=(opts)
      border(**opts)
      opts
    end

    # Sets all borders.
    # @param style [String, Symbol, Integer] The border style.
    # @param color [String, Symbol, nil] The color.
    # @return [void]
    # @api public
    #: (?style: String | Symbol | Integer, ?color: String | Symbol | nil) -> void
    def border_all(style: "thin", color: nil)
      norm_style = self.class.normalize_border_style(style)&.to_s || style.to_s
      color_opt = color ? { color: resolve_color(color) } : {}
      @border_props[:left] = { style: norm_style, **color_opt }
      @border_props[:right] = { style: norm_style, **color_opt }
      @border_props[:top] = { style: norm_style, **color_opt }
      @border_props[:bottom] = { style: norm_style, **color_opt }
      nil
    end

    # Sets the left border.
    # @param style [String, Symbol, Integer] The border style.
    # @param color [String, Symbol, nil] The color.
    # @return [void]
    # @api public
    #: (?style: String | Symbol | Integer, ?color: String | Symbol | nil) -> void
    def border_left(style: "thin", color: nil)
      norm_style = self.class.normalize_border_style(style)&.to_s || style.to_s
      @border_props[:left] = { style: norm_style, color: resolve_color(color) }.compact
      nil
    end

    # Sets the right border.
    # @param style [String, Symbol, Integer] The border style.
    # @param color [String, Symbol, nil] The color.
    # @return [void]
    # @api public
    #: (?style: String | Symbol | Integer, ?color: String | Symbol | nil) -> void
    def border_right(style: "thin", color: nil)
      norm_style = self.class.normalize_border_style(style)&.to_s || style.to_s
      @border_props[:right] = { style: norm_style, color: resolve_color(color) }.compact
      nil
    end

    # Sets the top border.
    # @param style [String, Symbol, Integer] The border style.
    # @param color [String, Symbol, nil] The color.
    # @return [void]
    # @api public
    #: (?style: String | Symbol | Integer, ?color: String | Symbol | nil) -> void
    def border_top(style: "thin", color: nil)
      norm_style = self.class.normalize_border_style(style)&.to_s || style.to_s
      @border_props[:top] = { style: norm_style, color: resolve_color(color) }.compact
      nil
    end

    # Sets the bottom border.
    # @param style [String, Symbol, Integer] The border style.
    # @param color [String, Symbol, nil] The color.
    # @return [void]
    # @api public
    #: (?style: String | Symbol | Integer, ?color: String | Symbol | nil) -> void
    def border_bottom(style: "thin", color: nil)
      norm_style = self.class.normalize_border_style(style)&.to_s || style.to_s
      @border_props[:bottom] = { style: norm_style, color: resolve_color(color) }.compact
      nil
    end

    # rubocop:disable Naming/MethodParameterName
    # Sets diagonal borders.
    # @param style [String, Symbol, Integer] The border style.
    # @param color [String, Symbol, nil] The color.
    # @param up [Boolean] Diagonal up.
    # @param down [Boolean] Diagonal down.
    # @return [void]
    # @api public
    #: (?style: String | Symbol | Integer, ?color: String | Symbol | nil, ?up: bool, ?down: bool) -> void
    def border_diagonal(style: "thin", color: nil, up: false, down: false)
      norm_style = self.class.normalize_border_style(style)&.to_s || style.to_s
      @border_props[:diagonal] = { style: norm_style, color: resolve_color(color) }.compact
      @border_props[:diagonal_up] = true if up
      @border_props[:diagonal_down] = true if down
      nil
    end
    # rubocop:enable Naming/MethodParameterName

    # --- Alignment Properties ---

    # Configures or returns alignment properties.
    # When called with 0 arguments/keywords, returns the current alignment Hash.
    # When called with keyword arguments, merges the properties and returns nil.
    #
    # @param horizontal [String, Symbol, nil] Horizontal alignment (:left, :center, :right, :fill, :justify, etc.).
    # @param vertical [String, Symbol, nil] Vertical alignment (:top, :center, :bottom, :justify, :distributed).
    # @param wrap_text [Boolean, nil] Whether to wrap text within the cell.
    # @param text_rotation [Integer, nil] Text rotation in degrees (0-180 or 255 for vertical).
    # @param indent [Integer, nil] Indentation level (0-250).
    # @param shrink_to_fit [Boolean, nil] Whether to shrink text to fit cell width.
    # @param reading_order [Integer, nil] Reading order (0: context, 1: left-to-right, 2: right-to-left).
    # @param justify_last_line [Boolean, nil] Whether to justify the last line of text.
    # @param relative_indent [Integer, nil] Relative indentation level.
    # @return [Hash, void]
    # @api public
    #: (?horizontal: (String | Symbol)?, ?vertical: (String | Symbol)?, ?wrap_text: bool?, ?text_rotation: Integer?, ?indent: Integer?, ?shrink_to_fit: bool?, ?reading_order: Integer?, ?justify_last_line: bool?, ?relative_indent: Integer?) -> (Hash[Symbol, untyped] | void)
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
      nil
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
    # @return [String, Symbol, void, nil]
    # @api public
    #: (?untyped value) -> (String | Symbol | void | nil)
    def align_horizontal(value = OMITTED)
      return @alignment[:horizontal] if value.equal?(OMITTED)

      @alignment[:horizontal] = value
      nil
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
    # @return [String, Symbol, void, nil]
    # @api public
    #: (?untyped value) -> (String | Symbol | void | nil)
    def align_vertical(value = OMITTED)
      return @alignment[:vertical] if value.equal?(OMITTED)

      @alignment[:vertical] = value
      nil
    end

    # Sets vertical alignment.
    # @param value [String, Symbol]
    # @return [String, Symbol]
    # @api public
    #: (String | Symbol value) -> (String | Symbol)
    def align_vertical=(value)
      @alignment[:vertical] = value
    end

    # Sets or gets text wrapping.
    # @param value [Boolean, Object] Whether to wrap text.
    # @return [Boolean, void]
    # @api public
    #: (?untyped value) -> (bool | void)
    def wrap_text(value = OMITTED)
      return wrap_text? if value.equal?(OMITTED)

      @alignment[:wrap_text] = value ? true : false
      nil
    end

    # Sets text wrapping.
    # @param value [Boolean]
    # @return [Boolean]
    # @api public
    #: (bool value) -> bool
    def wrap_text=(value)
      @alignment[:wrap_text] = value ? true : false
    end

    # Returns whether text wrapping is enabled.
    # @return [Boolean]
    # @api public
    #: () -> bool
    def wrap_text?
      @alignment[:wrap_text] ? true : false
    end

    # Sets or gets shrink to fit.
    # @param value [Boolean, Object] Whether to shrink text.
    # @return [Boolean, void]
    # @api public
    #: (?untyped value) -> (bool | void)
    def shrink_to_fit(value = OMITTED)
      return shrink_to_fit? if value.equal?(OMITTED)

      @alignment[:shrink_to_fit] = value ? true : false
      nil
    end

    # Sets shrink to fit.
    # @param value [Boolean]
    # @return [Boolean]
    # @api public
    #: (bool value) -> bool
    def shrink_to_fit=(value)
      @alignment[:shrink_to_fit] = value ? true : false
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
    # @return [Numeric, void, nil]
    # @api public
    #: (?untyped value) -> (Numeric | void | nil)
    def text_rotation(value = OMITTED)
      return @alignment[:text_rotation] if value.equal?(OMITTED)

      @alignment[:text_rotation] = value
      nil
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
    # @return [Numeric, void, nil]
    # @api public
    #: (?untyped value) -> (Numeric | void | nil)
    def indent(value = OMITTED)
      return @alignment[:indent] if value.equal?(OMITTED)

      @alignment[:indent] = value.to_i
      nil
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
    # When called with keyword arguments, merges the properties and returns nil.
    #
    # @param locked [Boolean, nil]
    # @param hidden [Boolean, nil]
    # @return [Hash, void]
    # @api public
    #: (?locked: bool?, ?hidden: bool?) -> (Hash[Symbol, untyped] | void)
    def protection(locked: nil, hidden: nil)
      return @protection if locked.nil? && hidden.nil?

      @protection[:locked] = locked unless locked.nil?
      @protection[:hidden] = hidden unless hidden.nil?
      nil
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

    # Sets or gets whether the cell is locked.
    #
    # @param value [Boolean, Object]
    # @return [Boolean, void]
    # @api public
    #: (?untyped value) -> (bool | void)
    def locked(value = OMITTED)
      return locked? if value.equal?(OMITTED)

      @protection[:locked] = value ? true : false
      nil
    end

    # Sets whether the cell is locked.
    #
    # @param value [Boolean]
    # @return [Boolean]
    # @api public
    #: (bool value) -> bool
    def locked=(value)
      @protection[:locked] = value ? true : false
    end

    # Returns whether the cell is locked.
    # @return [Boolean]
    # @api public
    #: () -> bool
    def locked?
      @protection[:locked] != false
    end

    # Sets or gets whether the cell formula is hidden.
    #
    # @param value [Boolean, Object]
    # @return [Boolean, void]
    # @api public
    #: (?untyped value) -> (bool | void)
    def hidden(value = OMITTED)
      return hidden? if value.equal?(OMITTED)

      @protection[:hidden] = value ? true : false
      nil
    end

    # Sets whether the cell formula is hidden.
    #
    # @param value [Boolean]
    # @return [Boolean]
    # @api public
    #: (bool value) -> bool
    def hidden=(value)
      @protection[:hidden] = value ? true : false
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
    # @param num_fmt_id [String, Integer, Object, nil] The format id or format string.
    # @return [String, Integer, void, nil]
    # @api public
    #: (?String | Integer | Object | nil num_fmt_id) -> (String | Integer | void | nil)
    def number_format(num_fmt_id = OMITTED)
      return @num_fmt_id if num_fmt_id.equal?(OMITTED) || num_fmt_id.nil?

      @num_fmt_id = num_fmt_id
      nil
    end

    # Sets the number format.
    # @param num_fmt_id [String, Integer]
    # @return [String, Integer]
    # @api public
    #: (String | Integer num_fmt_id) -> (String | Integer)
    def number_format=(num_fmt_id)
      @num_fmt_id = num_fmt_id
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
