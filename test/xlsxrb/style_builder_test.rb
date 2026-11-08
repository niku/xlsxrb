# frozen_string_literal: true

require "test_helper"

class StyleBuilderTest < Test::Unit::TestCase
  cover Xlsxrb::StyleBuilder

  # --- resolve_color ---

  test "resolve_color resolves symbols, prefixed strings, hex, and nil" do
    b = Xlsxrb::StyleBuilder.new

    assert_nil(b.send(:resolve_color, nil))
    assert_equal("FF000000", b.send(:resolve_color, :black))
    assert_equal("FFFFFFFF", b.send(:resolve_color, :white))
    assert_equal("FFFF0000", b.send(:resolve_color, :red))
    assert_equal("FF00FF00", b.send(:resolve_color, :green))
    assert_equal("FF0000FF", b.send(:resolve_color, :blue))
    assert_equal("FFFFFF00", b.send(:resolve_color, :yellow))
    assert_equal("FF00FFFF", b.send(:resolve_color, :cyan))
    assert_equal("FFFF00FF", b.send(:resolve_color, :magenta))
    assert_equal("FF808080", b.send(:resolve_color, :gray))
    assert_equal("FF808080", b.send(:resolve_color, :grey))

    # Colon-prefixed string and single colon boundary
    assert_equal("FFFF0000", b.send(:resolve_color, ":red"))
    assert_equal("FF0000FF", b.send(:resolve_color, ":blue"))
    assert_equal(":not_in_colors", b.send(:resolve_color, ":not_in_colors"))
    assert_equal(":", b.send(:resolve_color, ":"))

    # Unknown symbol falls back to symbol to_s
    assert_equal("custom_sym", b.send(:resolve_color, :custom_sym))

    # Custom hex, empty string, and arbitrary types
    assert_equal("FF112233", b.send(:resolve_color, "FF112233"))
    assert_equal("custom_color", b.send(:resolve_color, "custom_color"))
    assert_equal("", b.send(:resolve_color, ""))
    assert_equal("12345", b.send(:resolve_color, 12_345))

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

  # --- Font properties & getters ---

  test "font property setters and 0-argument getters" do
    b = Xlsxrb::StyleBuilder.new("header")
    assert_equal(false, b.bold?)
    assert_equal(false, b.italic?)
    assert_nil(b.font_size)
    assert_nil(b.font_name)
    assert_nil(b.font_color)
    assert_nil(b.underline)
    assert_equal(false, b.strike?)
    assert_nil(b.vert_align)
    assert_equal(false, b.superscript?)
    assert_equal(false, b.subscript?)

    refute_respond_to(b, :bold)
    refute_respond_to(b, :italic)
    refute_respond_to(b, :strike)
    refute_respond_to(b, :superscript)
    refute_respond_to(b, :subscript)

    b.bold = true
    b.italic = true
    b.font_size = 14
    b.font_name = "Arial"
    b.font_color = :red
    b.underline = "double"
    b.strike = true
    b.vert_align = "superscript"

    assert_equal(true, b.bold?)
    assert_equal(true, b.font_props[:bold])
    assert_equal(true, b.italic?)
    assert_equal(true, b.font_props[:italic])
    assert_equal(14, b.font_size)
    assert_equal("Arial", b.font_name)
    assert_equal("Arial", b.font_props[:name])
    assert_equal("FFFF0000", b.font_color)
    assert_equal("FFFF0000", b.font_props[:color])
    assert_equal("double", b.underline)
    assert_equal("double", b.font_props[:underline])
    assert_equal(true, b.strike?)
    assert_equal(true, b.font_props[:strike])
    assert_equal("superscript", b.vert_align)
    assert_equal("superscript", b.font_props[:vert_align])
    assert_equal(true, b.superscript?)

    b.subscript = true
    assert_equal("subscript", b.vert_align)
    assert_equal(true, b.subscript?)
    assert_equal(false, b.superscript?)

    # font(**opts) bulk helper and 0-arg getter
    b2 = Xlsxrb::StyleBuilder.new
    assert_equal({}, b2.font)
    assert_nil(b2.font(bold: false, size: 10.5, name: "Calibri", color: :blue))
    assert_equal(false, b2.font_props[:bold])
    assert_equal(10.5, b2.font_props[:size])
    assert_equal("Calibri", b2.font_props[:name])
    assert_equal("FF0000FF", b2.font_props[:color])
    assert_equal(b2.font_props, b2.font)
  end

  # --- Fill properties ---

  test "fill property setters, getters, and gradient" do
    b = Xlsxrb::StyleBuilder.new
    assert_nil(b.fill_color)
    assert_nil(b.fill_pattern)
    assert_equal({}, b.fill)

    b.fill_color = :yellow
    assert_equal("solid", b.fill_props[:pattern])
    assert_equal("FFFFFF00", b.fill_props[:fg_color])
    assert_equal("FFFFFF00", b.fill_color)

    assert_nil(b.fill(pattern: "gray125", fg_color: :white, bg_color: :black))
    assert_equal("gray125", b.fill_props[:pattern])
    assert_equal("FFFFFFFF", b.fill_props[:fg_color])
    assert_equal("FF000000", b.fill_props[:bg_color])
    assert_equal(b.fill_props, b.fill)

    assert_nil(b.fill_gradient(type: "linear", degree: 90, stops: [{ position: 0, color: "FF000000" }]))
    assert_equal("linear", b.fill_props[:gradient][:type])
    assert_equal(90, b.fill_props[:gradient][:degree])
    assert_equal(1, b.fill_props[:gradient][:stops].size)
    assert_equal(b.fill_props[:gradient], b.fill_gradient)
  end

  # --- Border properties ---

  test "border directional setters and 0-argument getters" do
    b = Xlsxrb::StyleBuilder.new
    assert_equal({}, b.border)

    assert_nil(b.border_all(style: "thick", color: :blue))
    assert_equal({ style: "thick", color: "FF0000FF" }, b.border_props[:left])
    assert_equal({ style: "thick", color: "FF0000FF" }, b.border_props[:right])
    assert_equal({ style: "thick", color: "FF0000FF" }, b.border_props[:top])
    assert_equal({ style: "thick", color: "FF0000FF" }, b.border_props[:bottom])

    assert_nil(b.border_left(style: "double", color: :red))
    assert_nil(b.border_right(style: "dashed"))
    assert_nil(b.border_top(style: "dotted", color: :green))
    assert_nil(b.border_bottom(style: "thin"))
    assert_nil(b.border_diagonal(style: "medium", color: :black, up: true, down: true))

    assert_equal({ style: "double", color: "FFFF0000" }, b.border_props[:left])
    assert_equal({ style: "dashed" }, b.border_props[:right])
    assert_equal({ style: "dotted", color: "FF00FF00" }, b.border_props[:top])
    assert_equal({ style: "thin" }, b.border_props[:bottom])
    assert_equal({ style: "medium", color: "FF000000" }, b.border_props[:diagonal])
    assert_equal(true, b.border_props[:diagonal_up])
    assert_equal(true, b.border_props[:diagonal_down])
    assert_equal(b.border_props, b.border)
  end

  test "border methods accept integer styles and symbol aliases" do
    b = Xlsxrb::StyleBuilder.new
    assert_nil(b.border_all(style: 1, color: :navy))
    assert_equal({ style: "thin", color: "FF000080" }, b.border_props[:left])
    assert_equal({ style: "thin", color: "FF000080" }, b.border_props[:right])

    assert_nil(b.border_left(style: :border_medium, color: "#FF0000"))
    assert_nil(b.border_right(style: :medium_dashed))
    assert_equal({ style: "medium", color: "FFFF0000" }, b.border_props[:left])
    assert_equal({ style: "mediumDashed" }, b.border_props[:right])
  end

  # --- Alignment properties ---

  test "alignment property setters and 0-argument getters" do
    b = Xlsxrb::StyleBuilder.new
    assert_equal({}, b.alignment)
    assert_nil(b.align_horizontal)
    assert_nil(b.align_vertical)
    assert_equal(false, b.wrap_text?)
    assert_equal(false, b.shrink_to_fit?)
    assert_nil(b.text_rotation)
    assert_nil(b.indent)

    refute_respond_to(b, :wrap_text)
    refute_respond_to(b, :shrink_to_fit)

    b.align_horizontal = "center"
    b.align_vertical = "top"
    b.wrap_text = true
    b.shrink_to_fit = false
    b.text_rotation = 45
    b.indent = 2.5

    assert_equal("center", b.alignment[:horizontal])
    assert_equal("center", b.align_horizontal)
    assert_equal("top", b.alignment[:vertical])
    assert_equal("top", b.align_vertical)
    assert_equal(true, b.alignment[:wrap_text])
    assert_equal(true, b.wrap_text?)
    assert_equal(false, b.alignment[:shrink_to_fit])
    assert_equal(false, b.shrink_to_fit?)
    assert_equal(45, b.alignment[:text_rotation])
    assert_equal(45, b.text_rotation)
    assert_equal(2, b.alignment[:indent])
    assert_equal(2, b.indent)
  end

  # --- Protection properties ---

  test "protection property setters and 0-argument getters" do
    b = Xlsxrb::StyleBuilder.new
    assert_equal({}, b.protection)
    assert_equal(true, b.locked?)
    assert_equal(false, b.hidden?)

    refute_respond_to(b, :locked)
    refute_respond_to(b, :hidden)

    b.locked = false
    b.hidden = true
    assert_equal(false, b.locked?)
    assert_equal(true, b.hidden?)
    assert_equal({ locked: false, hidden: true }, b.protection)

    b2 = Xlsxrb::StyleBuilder.new
    assert_nil(b2.protection(locked: true, hidden: false))
    assert_equal({ locked: true, hidden: false }, b2.protection)
  end

  # --- Number format & apply_options! ---

  test "number format and apply_options! option hash" do
    b = Xlsxrb::StyleBuilder.new
    assert_nil(b.number_format)
    b.number_format = "$#,##0.00"
    assert_equal("$#,##0.00", b.number_format)

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
    assert_equal(49, b_opt.number_format)
  end

  test "fractional font_size preserves Float and serializes to sz val in styles XML" do
    b1 = Xlsxrb::StyleBuilder.new
    b1.font_size = 10.5
    assert_equal(10.5, b1.font_size)
    assert_equal(10.5, b1.font_props[:size])

    b2 = Xlsxrb::StyleBuilder.new
    b2.font_size = 10.5
    assert_equal(10.5, b2.font_size)
    assert_equal(10.5, b2.font_props[:size])

    # Normalization of whole Float to Integer
    b3 = Xlsxrb::StyleBuilder.new
    b3.font_size = 12.0
    assert_equal(12, b3.font_size)
    assert_instance_of(Integer, b3.font_props[:size])

    # XML serialization contains <sz val="10.5"/>
    wb = Xlsxrb.build do |b|
      b.style(:half_pt, font_size: 10.5)
      b.sheet("Sheet1") do |s|
        s.row(["Test"], row_style: :half_pt)
      end
    end
    data = Xlsxrb.write(wb)
    styles_xml = Xlsxrb::Ooxml::ZipReader.open(data) { |reader| reader.read_entry("xl/styles.xml") }
    assert_include(styles_xml, '<sz val="10.5"/>')
  end

  test "border and alignment accept explicit keyword arguments" do
    sb = Xlsxrb::StyleBuilder.new
    assert_nil(
      sb.border(
        left: { style: "thin", color: :red },
        right: { style: "double", color: :blue },
        top: { style: "dashed" },
        bottom: { style: "thick" },
        diagonal: { style: "thin", up: true }
      )
    )
    assert_equal("thin", sb.border_props[:left][:style])
    assert_equal("FFFF0000", sb.border_props[:left][:color])
    assert_equal("double", sb.border_props[:right][:style])
    assert_equal("0000FF", sb.border_props[:right][:color].sub(/\AFF/, ""))
    assert_equal("dashed", sb.border_props[:top][:style])
    assert_equal("thick", sb.border_props[:bottom][:style])
    assert_equal("thin", sb.border_props[:diagonal][:style])
    assert_equal(true, sb.border_props[:diagonal_up])

    assert_nil(
      sb.alignment(
        horizontal: :center,
        vertical: :top,
        wrap_text: true,
        text_rotation: 45,
        indent: 2,
        shrink_to_fit: false,
        reading_order: 1,
        justify_last_line: true,
        relative_indent: 3
      )
    )
    assert_equal(:center, sb.alignment[:horizontal])
    assert_equal(:top, sb.alignment[:vertical])
    assert_equal(true, sb.alignment[:wrap_text])
    assert_equal(45, sb.alignment[:text_rotation])
    assert_equal(2, sb.alignment[:indent])
    assert_equal(false, sb.alignment[:shrink_to_fit])
    assert_equal(1, sb.alignment[:reading_order])
    assert_equal(true, sb.alignment[:justify_last_line])
    assert_equal(3, sb.alignment[:relative_indent])
  end
end
