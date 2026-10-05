# frozen_string_literal: true

require "test_helper"

class ColorsTest < Test::Unit::TestCase
  cover Xlsxrb::Colors

  test "INDEXED_COLORS has exactly 64 entries per ECMA-376 Part 4 §18.8.27" do
    assert_equal(64, Xlsxrb::Colors::INDEXED_COLORS.size)
    assert_true(Xlsxrb::Colors::INDEXED_COLORS.frozen?)
    Xlsxrb::Colors::INDEXED_COLORS.each_with_index do |hex, i|
      assert_match(/\A[0-9A-F]{6}\z/, hex, "Expected 6-char uppercase hex for index #{i}")
    end
  end

  test "palette_color resolves valid indices with and without alpha" do
    assert_equal("000000", Xlsxrb::Colors.palette_color(0))
    assert_equal("FF000000", Xlsxrb::Colors.palette_color(0, alpha: true))
    assert_equal("FFFFFF", Xlsxrb::Colors.palette_color(1))
    assert_equal("FFFFFFFF", Xlsxrb::Colors.palette_color(1, alpha: true))
    assert_equal("0000FF", Xlsxrb::Colors.palette_color(12))
    assert_equal("FF0000FF", Xlsxrb::Colors.palette_color(12, alpha: true))
    assert_equal("333333", Xlsxrb::Colors.palette_color(63))
    assert_equal("FF333333", Xlsxrb::Colors.palette_color(63, alpha: true))

    # Out of bounds and non-integers
    assert_nil(Xlsxrb::Colors.palette_color(-1))
    assert_nil(Xlsxrb::Colors.palette_color(64))
    assert_nil(Xlsxrb::Colors.palette_color(100))
    assert_nil(Xlsxrb::Colors.palette_color(nil))
    assert_nil(Xlsxrb::Colors.palette_color("12"))
  end

  test "to_hex resolves palette indices, CSS names, hex strings, and RGB integers" do
    assert_nil(Xlsxrb::Colors.to_hex(nil))

    # Palette indices 0..63
    assert_equal("0000FF", Xlsxrb::Colors.to_hex(12))
    assert_equal("FF0000FF", Xlsxrb::Colors.to_hex(12, alpha: true))
    assert_equal("000000", Xlsxrb::Colors.to_hex(0))
    assert_equal("FFFFFF", Xlsxrb::Colors.to_hex(1))

    # CSS color names
    assert_equal("000080", Xlsxrb::Colors.to_hex(:navy))
    assert_equal("FF000080", Xlsxrb::Colors.to_hex(:navy, alpha: true))
    assert_equal("FF0000", Xlsxrb::Colors.to_hex("red"))

    # Hex strings
    assert_equal("FF0000", Xlsxrb::Colors.to_hex("#FF0000"))
    assert_equal("FFFF0000", Xlsxrb::Colors.to_hex("#FF0000", alpha: true))
    assert_equal("FF0000", Xlsxrb::Colors.to_hex("FF0000"))
    assert_equal("FF0000", Xlsxrb::Colors.to_hex("#F00"))

    # RGB integers (> 63)
    assert_equal("FF0000", Xlsxrb::Colors.to_hex(0xFF0000))
    assert_equal("FFFF0000", Xlsxrb::Colors.to_hex(0xFF0000, alpha: true))
    assert_equal("000080", Xlsxrb::Colors.to_hex(0x000080))
  end

  test "top-level Xlsxrb facade methods delegate to Colors" do
    assert_equal("0000FF", Xlsxrb.palette_color(12))
    assert_equal("FF0000FF", Xlsxrb.palette_color(12, alpha: true))
    assert_equal("0000FF", Xlsxrb.to_hex_color(12))
    assert_equal("0000FF", Xlsxrb.color_to_hex(12))
    assert_equal("000080", Xlsxrb.to_hex_color(:navy))
    assert_equal("FF0000", Xlsxrb.to_hex_color("#FF0000"))
  end
end
