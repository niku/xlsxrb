# frozen_string_literal: true

require "test_helper"

class StyleBuilderTest < Test::Unit::TestCase
  cover Xlsxrb::StyleBuilder

  # --- resolve_color ---

  test "resolve_color resolves symbols, prefixed strings, hex, and nil" do
    b = Xlsxrb::StyleBuilder.new

    assert_nil(b.resolve_color(nil))
    assert_equal("FF000000", b.resolve_color(:black))
    assert_equal("FFFFFFFF", b.resolve_color(:white))
    assert_equal("FFFF0000", b.resolve_color(:red))
    assert_equal("FF00FF00", b.resolve_color(:green))
    assert_equal("FF0000FF", b.resolve_color(:blue))
    assert_equal("FFFFFF00", b.resolve_color(:yellow))
    assert_equal("FF00FFFF", b.resolve_color(:cyan))
    assert_equal("FFFF00FF", b.resolve_color(:magenta))
    assert_equal("FF808080", b.resolve_color(:gray))
    assert_equal("FF808080", b.resolve_color(:grey))

    # Colon-prefixed string and single colon boundary
    assert_equal("FFFF0000", b.resolve_color(":red"))
    assert_equal("FF0000FF", b.resolve_color(":blue"))
    assert_equal(":not_in_colors", b.resolve_color(":not_in_colors"))
    assert_equal(":", b.resolve_color(":"))

    # Unknown symbol falls back to symbol to_s
    assert_equal("custom_sym", b.resolve_color(:custom_sym))

    # Custom hex, empty string, and arbitrary types
    assert_equal("FF112233", b.resolve_color("FF112233"))
    assert_equal("custom_color", b.resolve_color("custom_color"))
    assert_equal("", b.resolve_color(""))
    assert_equal("12345", b.resolve_color(12_345))

    # Pure class method direct calls
    assert_nil(Xlsxrb::StyleBuilder.resolve_color(nil))
    assert_equal("FFFF0000", Xlsxrb::StyleBuilder.resolve_color(:red))
    assert_equal("FF0000FF", Xlsxrb::StyleBuilder.resolve_color(":blue"))
    assert_equal("FF112233", Xlsxrb::StyleBuilder.resolve_color("FF112233"))
    assert_equal("FF000080", Xlsxrb::StyleBuilder.resolve_color(:navy))
    assert_equal("FFFF0000", Xlsxrb::StyleBuilder.resolve_color("#FF0000"))
    assert_equal("FF0000FF", Xlsxrb::StyleBuilder.resolve_color(12))
    assert_equal("FF000000", Xlsxrb::StyleBuilder.resolve_color(0))
    assert_equal("FFFFFFFF", Xlsxrb::StyleBuilder.resolve_color(1))
    assert_equal("FF333333", Xlsxrb::StyleBuilder.resolve_color(63))
  end

  test "normalize_color normalizes CSS colors, hex codes, integers, and alpha" do
    assert_nil(Xlsxrb::StyleBuilder.normalize_color(nil))
    assert_nil(Xlsxrb::StyleBuilder.normalize_color(""))
    assert_nil(Xlsxrb::StyleBuilder.normalize_color(":"))
    assert_nil(Xlsxrb::StyleBuilder.normalize_color("unknown_color"))

    # CSS color names (with and without alpha)
    assert_equal("FF000080", Xlsxrb::StyleBuilder.normalize_color(:navy))
    assert_equal("000080", Xlsxrb::StyleBuilder.normalize_color(:navy, alpha: false))
    assert_equal("FFFFD700", Xlsxrb::StyleBuilder.normalize_color("gold"))
    assert_equal("FFDC143C", Xlsxrb::StyleBuilder.normalize_color(:crimson))
    assert_equal("FF663399", Xlsxrb::StyleBuilder.normalize_color(:rebeccapurple))

    # Hex strings (3-digit, 6-digit, 8-digit)
    assert_equal("FFFF0000", Xlsxrb::StyleBuilder.normalize_color("#F00"))
    assert_equal("FF0000", Xlsxrb::StyleBuilder.normalize_color("#F00", alpha: false))
    assert_equal("FFFF0000", Xlsxrb::StyleBuilder.normalize_color("F00"))
    assert_equal("FFFF0000", Xlsxrb::StyleBuilder.normalize_color("#FF0000"))
    assert_equal("FF0000", Xlsxrb::StyleBuilder.normalize_color("#FF0000", alpha: false))
    assert_equal("FFFF0000", Xlsxrb::StyleBuilder.normalize_color("FF0000"))
    assert_equal("FFFF0000", Xlsxrb::StyleBuilder.normalize_color("#FFFF0000"))
    assert_equal("FF0000", Xlsxrb::StyleBuilder.normalize_color("#FFFF0000", alpha: false))
    assert_equal("FF123456", Xlsxrb::StyleBuilder.normalize_color("FF123456"))
    assert_equal("123456", Xlsxrb::StyleBuilder.normalize_color("FF123456", alpha: false))

    # Integer RGB
    assert_equal("FFFF0000", Xlsxrb::StyleBuilder.normalize_color(0xFF0000))
    assert_equal("FF0000", Xlsxrb::StyleBuilder.normalize_color(0xFF0000, alpha: false))
    assert_equal("FF000080", Xlsxrb::StyleBuilder.normalize_color(0x000080))
  end

  test "normalize_border_style normalizes border symbols, aliases, strings, and indices" do
    assert_nil(Xlsxrb::StyleBuilder.normalize_border_style(nil))
    assert_nil(Xlsxrb::StyleBuilder.normalize_border_style(:invalid_style))

    # Integer mappings 0..13
    assert_equal(:none, Xlsxrb::StyleBuilder.normalize_border_style(0))
    assert_equal(:thin, Xlsxrb::StyleBuilder.normalize_border_style(1))
    assert_equal(:medium, Xlsxrb::StyleBuilder.normalize_border_style(2))
    assert_equal(:dashed, Xlsxrb::StyleBuilder.normalize_border_style(3))
    assert_equal(:dotted, Xlsxrb::StyleBuilder.normalize_border_style(4))
    assert_equal(:thick, Xlsxrb::StyleBuilder.normalize_border_style(5))
    assert_equal(:double, Xlsxrb::StyleBuilder.normalize_border_style(6))
    assert_equal(:hair, Xlsxrb::StyleBuilder.normalize_border_style(7))
    assert_equal(:mediumDashed, Xlsxrb::StyleBuilder.normalize_border_style(8))
    assert_equal(:dashDot, Xlsxrb::StyleBuilder.normalize_border_style(9))
    assert_equal(:mediumDashDot, Xlsxrb::StyleBuilder.normalize_border_style(10))
    assert_equal(:dashDotDot, Xlsxrb::StyleBuilder.normalize_border_style(11))
    assert_equal(:mediumDashDotDot, Xlsxrb::StyleBuilder.normalize_border_style(12))
    assert_equal(:slantDashDot, Xlsxrb::StyleBuilder.normalize_border_style(13))

    # Symbol and String styles
    assert_equal(:thin, Xlsxrb::StyleBuilder.normalize_border_style(:thin))
    assert_equal(:thin, Xlsxrb::StyleBuilder.normalize_border_style("thin"))
    assert_equal(:thin, Xlsxrb::StyleBuilder.normalize_border_style(:border_thin))
    assert_equal(:mediumDashed, Xlsxrb::StyleBuilder.normalize_border_style(:medium_dashed))
    assert_equal(:mediumDashed, Xlsxrb::StyleBuilder.normalize_border_style(:border_medium_dashed))
    assert_equal(:mediumDashed, Xlsxrb::StyleBuilder.normalize_border_style("mediumDashed"))
  end

  # --- Font properties & chaining ---

  test "font fluent DSL and individual property setters" do
    b = Xlsxrb::StyleBuilder.new("header")
    ret = b.bold.italic.size(14).font_name("Arial").font_color(:red).underline("double").strike.vert_align("superscript")

    assert_same(b, ret)
    assert_equal(true, b.font_props[:bold])
    assert_equal(true, b.font_props[:italic])
    assert_equal(14, b.font_props[:sz])
    assert_equal("Arial", b.font_props[:name])
    assert_equal("FFFF0000", b.font_props[:color])
    assert_equal("double", b.font_props[:underline])
    assert_equal(true, b.font_props[:strike])
    assert_equal("superscript", b.font_props[:vert_align])

    # font(**opts) bulk helper
    b2 = Xlsxrb::StyleBuilder.new
    b2.font(bold: false, size: 10.5, name: "Calibri", color: :blue)
    assert_equal(false, b2.font_props[:bold])
    assert_equal(10, b2.font_props[:sz])
    assert_equal("Calibri", b2.font_props[:name])
    assert_equal("FF0000FF", b2.font_props[:color])
  end

  # --- Fill properties ---

  test "fill fluent DSL with color, pattern, and gradient" do
    b = Xlsxrb::StyleBuilder.new
    b.fill_color(:yellow)
    assert_equal("solid", b.fill_props[:pattern])
    assert_equal("FFFFFF00", b.fill_props[:fg_color])

    b.fill(pattern: "gray125", fg_color: :white, bg_color: :black)
    assert_equal("gray125", b.fill_props[:pattern])
    assert_equal("FFFFFFFF", b.fill_props[:fg_color])
    assert_equal("FF000000", b.fill_props[:bg_color])

    b.fill_gradient(type: "linear", degree: 90, stops: [{ position: 0, color: "FF000000" }])
    assert_equal("linear", b.fill_props[:gradient][:type])
    assert_equal(90, b.fill_props[:gradient][:degree])
    assert_equal(1, b.fill_props[:gradient][:stops].size)
  end

  # --- Border properties ---

  test "border fluent DSL and directional borders" do
    b = Xlsxrb::StyleBuilder.new
    b.border_all(style: "thick", color: :blue)
    assert_equal({ style: "thick", color: "FF0000FF" }, b.border_props[:left])
    assert_equal({ style: "thick", color: "FF0000FF" }, b.border_props[:right])
    assert_equal({ style: "thick", color: "FF0000FF" }, b.border_props[:top])
    assert_equal({ style: "thick", color: "FF0000FF" }, b.border_props[:bottom])

    b.border_left(style: "double", color: :red)
    b.border_right(style: "dashed")
    b.border_top(style: "dotted", color: :green)
    b.border_bottom(style: "thin")
    b.border_diagonal(style: "medium", color: :black, up: true, down: true)

    assert_equal({ style: "double", color: "FFFF0000" }, b.border_props[:left])
    assert_equal({ style: "dashed" }, b.border_props[:right])
    assert_equal({ style: "dotted", color: "FF00FF00" }, b.border_props[:top])
    assert_equal({ style: "thin" }, b.border_props[:bottom])
    assert_equal({ style: "medium", color: "FF000000" }, b.border_props[:diagonal])
    assert_equal(true, b.border_props[:diagonal_up])
    assert_equal(true, b.border_props[:diagonal_down])
  end

  test "border methods accept integer styles and symbol aliases" do
    b = Xlsxrb::StyleBuilder.new
    b.border_all(style: 1, color: :navy)
    assert_equal({ style: "thin", color: "FF000080" }, b.border_props[:left])
    assert_equal({ style: "thin", color: "FF000080" }, b.border_props[:right])

    b.border_left(style: :border_medium, color: "#FF0000")
    b.border_right(style: :medium_dashed)
    assert_equal({ style: "medium", color: "FFFF0000" }, b.border_props[:left])
    assert_equal({ style: "mediumDashed" }, b.border_props[:right])
  end

  # --- Alignment properties ---

  test "alignment fluent DSL" do
    b = Xlsxrb::StyleBuilder.new
    b.align_horizontal("center")
     .align_vertical("top")
     .wrap_text(true)
     .shrink_to_fit(false)
     .text_rotation(45)
     .indent(2.5)

    assert_equal("center", b.alignment[:horizontal])
    assert_equal("top", b.alignment[:vertical])
    assert_equal(true, b.alignment[:wrap_text])
    assert_equal(false, b.alignment[:shrink_to_fit])
    assert_equal(45, b.alignment[:text_rotation])
    assert_equal(2, b.alignment[:indent])
  end

  # --- Number format & apply_options! ---

  test "number format and apply_options! option hash" do
    b = Xlsxrb::StyleBuilder.new
    b.number_format("$#,##0.00")
    assert_equal("$#,##0.00", b.num_fmt_id)

    # apply_options! with nested hashes
    b_opt = Xlsxrb::StyleBuilder.new
    b_opt.apply_options!(
      font: { bold: true, color: :red },
      fill: { color: :yellow },
      border: { all: { style: "thin", color: :black } },
      alignment: { horizontal: "right", wrap_text: true },
      number_format: 49
    )

    assert_equal(true, b_opt.font_props[:bold])
    assert_equal("FFFF0000", b_opt.font_props[:color])
    assert_equal("FFFFFF00", b_opt.fill_props[:fg_color])
    assert_equal({ style: "thin", color: "FF000000" }, b_opt.border_props[:left])
    assert_equal("right", b_opt.alignment[:horizontal])
    assert_equal(true, b_opt.alignment[:wrap_text])
    assert_equal(49, b_opt.num_fmt_id)
  end
end
