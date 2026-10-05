# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  # ECMA-376 Part 4 §18.8.27 Standard Indexed Color Palette and utilities.
  module Colors
    # Standard ECMA-376 indexed colors (indices 0 through 63 as 6-character RRGGBB hex strings).
    # Indices 8 through 63 represent the classic 56-color palette.
    INDEXED_COLORS = [
      "000000", # 0
      "FFFFFF", # 1
      "FF0000", # 2
      "00FF00", # 3
      "0000FF", # 4
      "FFFF00", # 5
      "FF00FF", # 6
      "00FFFF", # 7
      "000000", # 8: Black
      "FFFFFF", # 9: White
      "FF0000", # 10: Red
      "00FF00", # 11: Green
      "0000FF", # 12: Blue
      "FFFF00", # 13: Yellow
      "FF00FF", # 14: Magenta
      "00FFFF", # 15: Cyan
      "800000", # 16: Brown / Dark Red
      "008000", # 17: Olive / Dark Green
      "000080", # 18: Navy / Dark Blue
      "808000", # 19: Dark Yellow
      "800080", # 20: Purple
      "008080", # 21: Teal
      "C0C0C0", # 22: Silver / Light Gray
      "808080", # 23: Gray
      "9999FF", # 24
      "993366", # 25
      "FFFFCC", # 26
      "CCFFFF", # 27
      "660066", # 28
      "FF8080", # 29
      "0066CC", # 30
      "CCCCFF", # 31
      "000080", # 32
      "FF00FF", # 33
      "FFFF00", # 34
      "00FFFF", # 35
      "800080", # 36
      "800000", # 37
      "008080", # 38
      "0000FF", # 39
      "00CCFF", # 40
      "CCFFFF", # 41
      "CCFFCC", # 42
      "FFFFC0", # 43
      "CCE5FF", # 44
      "99CCFF", # 45
      "FF99CC", # 46
      "CC99FF", # 47
      "FFCC99", # 48
      "3366FF", # 49
      "33CCCC", # 50
      "99CC00", # 51
      "FFCC00", # 52
      "FF9900", # 53
      "FF6600", # 54
      "666699", # 55
      "969696", # 56
      "003366", # 57
      "339966", # 58
      "003300", # 59
      "333300", # 60
      "993300", # 61
      "993366", # 62
      "333333"  # 63
    ].freeze

    # Resolves a palette index (0..63) to an RRGGBB hex string (or AARRGGBB if alpha: true).
    #
    # @example
    #   Xlsxrb::Colors.palette_color(12)               #=> "0000FF"
    #   Xlsxrb::Colors.palette_color(12, alpha: true)  #=> "FF0000FF"
    #
    # @param index [Object, nil] Palette index (0..63).
    # @param alpha [Boolean] Whether to prefix "FF" alpha channel.
    # @return [String, nil]
    # @api public
    #: (untyped index, ?alpha: bool) -> String?
    def self.palette_color(index, alpha: false)
      return nil unless index.is_a?(Integer) && index >= 0 && index < INDEXED_COLORS.size

      hex = INDEXED_COLORS[index]
      alpha ? "FF#{hex}" : hex
    end

    # Converts any supported color representation (palette index 0..63, CSS name, #RGB, #RRGGBB, #AARRGGBB, integer RGB)
    # to a normalized hex string.
    #
    # @example
    #   Xlsxrb::Colors.to_hex(12)                      #=> "0000FF"
    #   Xlsxrb::Colors.to_hex(12, alpha: true)         #=> "FF0000FF"
    #   Xlsxrb::Colors.to_hex(:navy)                   #=> "000080"
    #   Xlsxrb::Colors.to_hex("#FF0000")               #=> "FF0000"
    #   Xlsxrb::Colors.to_hex(0xFF0000)                #=> "FF0000"
    #
    # @param color [Integer, String, Symbol, nil]
    # @param alpha [Boolean] Whether to include the alpha channel (default false -> "RRGGBB").
    # @return [String, nil]
    # @api public
    #: (Integer | String | Symbol | nil color, ?alpha: bool) -> String?
    def self.to_hex(color, alpha: false)
      return nil if color.nil?

      return palette_color(color, alpha: alpha) if color.is_a?(Integer) && color >= 0 && color < INDEXED_COLORS.size

      StyleBuilder.normalize_color(color, alpha: alpha)
    end
  end
end
