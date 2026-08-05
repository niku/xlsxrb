# frozen_string_literal: true

require "test_helper"

class StyleColorTest < Test::Unit::TestCase
  test "standard color symbols are supported" do
    builder = Xlsxrb::StyleBuilder.new("red_style")
    builder.font_color(:red)
    builder.fill(fg_color: :blue)
    builder.border_all(color: :green)

    assert_equal("FFFF0000", builder.font_props[:color])
    assert_equal("FF0000FF", builder.fill_props[:fg_color])
    assert_equal("FF00FF00", builder.border_props[:top][:color])

    # Check invalid symbol defaults to to_s
    builder.font_color(:purple)
    assert_equal("purple", builder.font_props[:color])
  end
end
