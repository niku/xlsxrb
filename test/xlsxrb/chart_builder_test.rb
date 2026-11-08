# frozen_string_literal: true

require "test_helper"

class ChartBuilderTest < Test::Unit::TestCase
  cover Xlsxrb::ChartBuilder

  test "chart_builder builds standard bar chart options" do
    builder = Xlsxrb::ChartBuilder.new
    builder.type = :bar
    builder.title = "Q1 Revenue"
    assert_nil(builder.legend(position: "r"))
    builder.style = 10
    builder.plot_visible_only = true
    builder.display_blanks_as = "zero"
    builder.show_legend_key = false

    assert_equal(:bar, builder.type)
    assert_equal("Q1 Revenue", builder.title)
    assert_equal({ position: "r" }, builder.legend)
    assert_equal(10, builder.style)
    assert_equal(true, builder.plot_visible_only?)
    assert_equal("zero", builder.display_blanks_as)
    assert_equal(false, builder.show_legend_key?)
    assert_false(builder.respond_to?(:plot_visible_only))
    assert_false(builder.respond_to?(:show_legend_key))

    ret_s1 = builder.series do |s|
      s.name = "Target"
      s.categories = "Sheet1!$A$2:$A$5"
      s.values = "Sheet1!$B$2:$B$5"
      assert_nil(s.fill(color: "00AA00"))
      s.smooth = false

      assert_equal("Target", s.name)
      assert_equal("Sheet1!$A$2:$A$5", s.categories)
      assert_equal("Sheet1!$B$2:$B$5", s.values)
      assert_equal({ color: "00AA00" }, s.fill)
      assert_equal(false, s.smooth?)
      assert_false(s.respond_to?(:smooth))
    end
    assert_nil(ret_s1)

    # Keyword argument series overload
    ret_s2 = builder.series(name: "Actual", values: "Sheet1!$C$2:$C$5")
    assert_nil(ret_s2)

    opts = builder.options
    assert_equal(:bar, opts[:type])
    assert_equal("Q1 Revenue", opts[:title])
    assert_equal({ position: "r" }, opts[:legend])
    assert_equal(10, opts[:style])
    assert_equal(true, opts[:plot_visible_only])
    assert_equal("zero", opts[:display_blanks_as])
    assert_equal(false, opts[:show_legend_key])
    assert_equal(2, opts[:series].size)
    assert_equal(2, builder.series.size)

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

  test "chart_builder direct setters and title kwargs" do
    builder = Xlsxrb::ChartBuilder.new
    builder.legend = "bottom"
    builder.plot_area = "custom_plot_area"
    builder.chart_space = "custom_space"
    builder.category_axis = "cat_axis_spec"
    builder.value_axis = "val_axis_spec"
    builder.view3d = "standard_3d"
    builder.data_labels = "show_val"

    opts = builder.options
    assert_equal("bottom", opts[:legend])
    assert_equal("bottom", builder.legend)
    assert_equal("custom_plot_area", opts[:plot_area])
    assert_equal("custom_plot_area", builder.plot_area)
    assert_equal("custom_space", opts[:chart_space])
    assert_equal("custom_space", builder.chart_space)
    assert_equal("cat_axis_spec", opts[:category_axis])
    assert_equal("cat_axis_spec", builder.category_axis)
    assert_equal("val_axis_spec", opts[:value_axis])
    assert_equal("val_axis_spec", builder.value_axis)
    assert_equal("standard_3d", opts[:view3d])
    assert_equal("standard_3d", builder.view3d)
    assert_equal("show_val", opts[:data_labels])
    assert_equal("show_val", builder.data_labels)

    # title(**kwargs)
    assert_nil(builder.title(text: "Sales", font_size: 14))
    assert_equal({ text: "Sales", font_size: 14 }, builder.title)

    # series= direct array assignment
    builder.series = [{ name: "S1" }]
    assert_equal([{ name: "S1" }], builder.series)

    # Empty builder returns [] for series
    b_empty = Xlsxrb::ChartBuilder.new
    assert_equal([], b_empty.series)
  end

  test "series_builder advanced properties and shapes" do
    sb = Xlsxrb::ChartBuilder::SeriesBuilder.new
    sb.name = "Line1"
    assert_nil(sb.marker(symbol: "circle", size: 5))
    assert_nil(sb.line(color: "FF0000", width: 2.25))
    assert_nil(sb.trendline(type: "linear"))
    assert_nil(sb.data_labels(value: true))
    sb.shape = "pyramid"
    sb.type = "scatter"

    assert_equal("Line1", sb.name)
    assert_equal({ symbol: "circle", size: 5 }, sb.marker)
    assert_equal({ color: "FF0000", width: 2.25 }, sb.line)
    assert_equal({ type: "linear" }, sb.trendline)
    assert_equal({ value: true }, sb.data_labels)
    assert_equal("pyramid", sb.shape)
    assert_equal("scatter", sb.type)

    opts = sb.options
    assert_equal("Line1", opts[:name])
    assert_equal({ symbol: "circle", size: 5 }, opts[:marker])
    assert_equal({ color: "FF0000", width: 2.25 }, opts[:line])
    assert_equal({ type: "linear" }, opts[:trendline])
    assert_equal({ value: true }, opts[:data_labels])
    assert_equal("pyramid", opts[:shape])
    assert_equal("scatter", opts[:type])

    # Direct setters
    sb.marker = "none"
    assert_equal("none", sb.marker)
    sb.fill = "blue"
    assert_equal("blue", sb.fill)
    sb.line = "dashed"
    assert_equal("dashed", sb.line)
    sb.trendline = "exponential"
    assert_equal("exponential", sb.trendline)
    sb.data_labels = "all"
    assert_equal("all", sb.data_labels)
  end

  test "chart_builder secondary value axis and series axis binding" do
    builder = Xlsxrb::ChartBuilder.new
    builder.type = :combo
    assert_nil(builder.secondary_value_axis(title: "Secondary Units", major_gridlines: false))

    builder.series do |s|
      s.name = "Primary Bar"
      s.values = "Sheet1!$B$2:$B$10"
    end

    builder.series do |s|
      s.name = "Secondary Line"
      s.type = :line
      s.values = "Sheet1!$C$2:$C$10"
      s.axis = :secondary
      assert_equal(:secondary, s.axis)
    end

    builder.series do |s|
      s.name = "Tertiary Series"
      s.axis = :secondary
    end

    opts = builder.options
    assert_equal(:combo, opts[:type])
    assert_equal({ title: "Secondary Units", major_gridlines: false }, opts[:secondary_value_axis])
    assert_equal(3, opts[:series].size)
    assert_nil(opts[:series][0][:secondary_axis])
    assert_equal(:secondary, opts[:series][1][:axis])
    assert_equal(true, opts[:series][1][:secondary_axis])
    assert_equal(:secondary, opts[:series][2][:axis])
    assert_equal(true, opts[:series][2][:secondary_axis])

    # Direct setter test
    b2 = Xlsxrb::ChartBuilder.new
    b2.secondary_value_axis = "sec_axis_spec"
    assert_equal("sec_axis_spec", b2.options[:secondary_value_axis])
    assert_equal("sec_axis_spec", b2.secondary_value_axis)
  end
end
