# frozen_string_literal: true

require "test_helper"

class ChartBuilderTest < Test::Unit::TestCase
  cover Xlsxrb::ChartBuilder

  test "chart_builder builds standard bar chart options" do
    builder = Xlsxrb::ChartBuilder.new
    builder.type(:bar)
    builder.title("Q1 Revenue")
    builder.legend(position: "r")
    builder.style(10)
    builder.plot_visible_only(true)
    builder.display_blanks_as("zero")
    builder.show_legend_key(false)

    builder.series do |s|
      s.name("Target")
      s.categories("Sheet1!$A$2:$A$5")
      s.values("Sheet1!$B$2:$B$5")
      s.fill(color: "00AA00")
      s.smooth(false)
    end

    # Pre-built series hash overload
    builder.series({ name: "Actual", values: "Sheet1!$C$2:$C$5" })

    opts = builder.options
    assert_equal(:bar, opts[:type])
    assert_equal("Q1 Revenue", opts[:title])
    assert_equal({ position: "r" }, opts[:legend])
    assert_equal(10, opts[:style])
    assert_equal(true, opts[:plot_visible_only])
    assert_equal("zero", opts[:display_blanks_as])
    assert_equal(false, opts[:show_legend_key])
    assert_equal(2, opts[:series].size)

    s1 = opts[:series][0]
    assert_equal("Target", s1[:name])
    assert_equal("Sheet1!$A$2:$A$5", s1[:categories])
    assert_equal("Sheet1!$B$2:$B$5", s1[:values])
    assert_equal({ color: "00AA00" }, s1[:fill])
    assert_equal(false, s1[:smooth])

    s2 = opts[:series][1]
    assert_equal("Actual", s2[:name])
    assert_equal("Sheet1!$C$2:$C$5", s2[:values])
  end

  test "chart_builder positional vs kwargs options" do
    builder = Xlsxrb::ChartBuilder.new
    # Positional string argument
    builder.legend("bottom")
    builder.plot_area("custom_plot_area")
    builder.chart_space("custom_space")
    builder.category_axis("cat_axis_spec")
    builder.value_axis("val_axis_spec")
    builder.view3d("standard_3d")
    builder.data_labels("show_val")

    opts = builder.options
    assert_equal("bottom", opts[:legend])
    assert_equal("custom_plot_area", opts[:plot_area])
    assert_equal("custom_space", opts[:chart_space])
    assert_equal("cat_axis_spec", opts[:category_axis])
    assert_equal("val_axis_spec", opts[:value_axis])
    assert_equal("standard_3d", opts[:view3d])
    assert_equal("show_val", opts[:data_labels])
  end

  test "series_builder advanced properties and shapes" do
    sb = Xlsxrb::ChartBuilder::SeriesBuilder.new
    sb.name("Line1")
    sb.marker(symbol: "circle", size: 5)
    sb.line(color: "FF0000", width: 2.25)
    sb.trendline(type: "linear")
    sb.data_labels(value: true)
    sb.shape("pyramid")
    sb.type("scatter")

    opts = sb.options
    assert_equal("Line1", opts[:name])
    assert_equal({ symbol: "circle", size: 5 }, opts[:marker])
    assert_equal({ color: "FF0000", width: 2.25 }, opts[:line])
    assert_equal({ type: "linear" }, opts[:trendline])
    assert_equal({ value: true }, opts[:data_labels])
    assert_equal("pyramid", opts[:shape])
    assert_equal("scatter", opts[:type])
  end
end
