# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  # Builder for block-style chart definitions.
  #
  # @example Create a chart definition
  #   builder = ChartBuilder.new
  #   builder.type(:bar)
  #   builder.title("Quarterly Sales")
  #   builder.series do |s|
  #     s.name("2026")
  #     s.categories("Sheet1!$A$2:$A$5")
  #     s.values("Sheet1!$B$2:$B$5")
  #   end
  #
  # @api public
  class ChartBuilder
    # @return [Hash{Symbol => Object}] The configured chart options.
    #: () -> void
    def initialize
      @options = {}
    end

    # @return [Hash{Symbol => Object}]
    #: Hash[Symbol, untyped]
    attr_reader :options

    # Sets or gets the chart type (e.g. :line, :bar, :pie, :area, :radar, :scatter).
    #
    # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#type=} instead.
    # @param value [Symbol, String, nil] Chart type.
    # @return [Symbol, String, nil]
    # @api public
    #: (?Symbol | String | nil value) -> (Symbol | String | nil)
    def type(value = nil)
      if value.nil?
        @options[:type]
      else
        warn "[Xlsxrb DEPRECATION] ChartBuilder#type(value) is deprecated; use #type= instead.", uplevel: 1
        @options[:type] = value
      end
    end

    # Sets the chart type.
    #
    # @param value [Symbol, String, nil]
    # @return [Symbol, String, nil]
    # @api public
    #: (Symbol | String | nil value) -> (Symbol | String | nil)
    def type=(value)
      @options[:type] = value
    end

    # Sets or gets the chart title.
    #
    # @param value [String, Hash, nil] Chart title text or options hash.
    # @param kwargs [Hash] Keyword arguments for title styling.
    # @return [String, Hash, nil]
    # @api public
    #: (?String | Hash[Symbol, untyped] | nil value, **untyped kwargs) -> (String | Hash[Symbol, untyped] | nil)
    def title(value = nil, **kwargs)
      if !kwargs.empty?
        @options[:title] = kwargs
        nil
      elsif !value.nil?
        warn "[Xlsxrb DEPRECATION] ChartBuilder#title(value) is deprecated; use #title= or #title(**kwargs) instead.", uplevel: 1
        @options[:title] = value
      else
        @options[:title]
      end
    end

    # Sets the chart title.
    #
    # @param value [String, Hash, nil]
    # @return [String, Hash, nil]
    # @api public
    #: (String | Hash[Symbol, untyped] | nil value) -> (String | Hash[Symbol, untyped] | nil)
    def title=(value)
      @options[:title] = value
    end

    # Adds a data series to the chart.
    #
    # @param value [Hash{Symbol => Object}, Array, nil] Series options hash or array.
    # @param kwargs [Hash] Keyword arguments for series configuration.
    # @yield [series_builder]
    # @yieldparam series_builder [SeriesBuilder]
    # @return [Array<Hash{Symbol => Object}>, nil]
    # @api public
    #: (?Hash[Symbol, untyped] | Array[Hash[Symbol, untyped]] | nil value, **untyped kwargs) ?{ (SeriesBuilder) -> void } -> (Array[Hash[Symbol, untyped]] | nil)
    def series(value = nil, **kwargs, &)
      if block_given?
        sb = SeriesBuilder.new
        yield sb
        (@options[:series] ||= []) << sb.options
        nil
      elsif !kwargs.empty?
        (@options[:series] ||= []) << kwargs
        nil
      elsif !value.nil?
        warn "[Xlsxrb DEPRECATION] ChartBuilder#series(value) with positional argument is deprecated; use #series=, #series(**kwargs), or #series(&block) instead.", uplevel: 1
        if value.is_a?(Array)
          (@options[:series] ||= []).concat(value)
        else
          (@options[:series] ||= []) << value
        end
      else
        @options[:series] ||= []
      end
    end

    # Sets series array.
    #
    # @param value [Array<Hash{Symbol => Object}>, Hash{Symbol => Object}]
    # @return [Array<Hash{Symbol => Object}>]
    # @api public
    #: (Array[Hash[Symbol, untyped]] | Hash[Symbol, untyped] value) -> (Array[Hash[Symbol, untyped]] | Hash[Symbol, untyped])
    def series=(value)
      @options[:series] = value.is_a?(Array) ? value : [value]
    end

    # Configures the legend property for this chart.
    #
    # @param args [Array] Positional arguments (e.g. position string).
    # @param kwargs [Hash] Keyword arguments for legend styling/layout.
    # @return [Object] The configured legend property.
    # @api public
    #: (*(Hash[Symbol, untyped] | String | Symbol) args, **untyped kwargs) -> untyped
    def legend(*args, **kwargs)
      if !args.empty?
        warn "[Xlsxrb DEPRECATION] ChartBuilder#legend(options) with positional argument is deprecated; use #legend= or #legend(**kwargs) instead.", uplevel: 1
        @options[:legend] = args.first
      elsif !kwargs.empty?
        @options[:legend] = kwargs
        nil
      else
        @options[:legend]
      end
    end

    # Sets the legend property.
    #
    # @param value [Object]
    # @return [Object]
    # @api public
    #: (untyped value) -> untyped
    def legend=(value)
      @options[:legend] = value
    end

    # Configures the plot_area property for this chart.
    #
    # @param args [Array] Positional arguments.
    # @param kwargs [Hash] Keyword arguments for plot area configuration.
    # @return [Object] The configured plot_area property.
    # @api public
    #: (*(Hash[Symbol, String | Numeric | bool | nil] | String) args, **String | Numeric | bool | nil kwargs) -> untyped
    def plot_area(*args, **kwargs)
      if !args.empty?
        warn "[Xlsxrb DEPRECATION] ChartBuilder#plot_area(options) with positional argument is deprecated; use #plot_area= or #plot_area(**kwargs) instead.", uplevel: 1
        @options[:plot_area] = args.first
      elsif !kwargs.empty?
        @options[:plot_area] = kwargs
        nil
      else
        @options[:plot_area]
      end
    end

    # Sets the plot_area property.
    #
    # @param value [Object]
    # @return [Object]
    # @api public
    #: (untyped value) -> untyped
    def plot_area=(value)
      @options[:plot_area] = value
    end

    # Configures the chart_space property for this chart.
    #
    # @param args [Array] Positional arguments.
    # @param kwargs [Hash] Keyword arguments.
    # @return [Object] The configured chart_space property.
    # @api public
    #: (*(Hash[Symbol, String | Numeric | bool | nil] | String) args, **String | Numeric | bool | nil kwargs) -> untyped
    def chart_space(*args, **kwargs)
      if !args.empty?
        warn "[Xlsxrb DEPRECATION] ChartBuilder#chart_space(options) with positional argument is deprecated; use #chart_space= or #chart_space(**kwargs) instead.", uplevel: 1
        @options[:chart_space] = args.first
      elsif !kwargs.empty?
        @options[:chart_space] = kwargs
        nil
      else
        @options[:chart_space]
      end
    end

    # Sets the chart_space property.
    #
    # @param value [Object]
    # @return [Object]
    # @api public
    #: (untyped value) -> untyped
    def chart_space=(value)
      @options[:chart_space] = value
    end

    # Configures the style index/id for this chart.
    #
    # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#style=} instead.
    # @param value [Object, nil]
    # @return [Object] The configured style property.
    # @api public
    #: (?untyped value) -> untyped
    def style(value = nil)
      if value.nil?
        @options[:style]
      else
        warn "[Xlsxrb DEPRECATION] ChartBuilder#style(value) is deprecated; use #style= instead.", uplevel: 1
        @options[:style] = value
      end
    end

    # Sets the style property.
    #
    # @param value [Object]
    # @return [Object]
    # @api public
    #: (untyped value) -> untyped
    def style=(value)
      @options[:style] = value
    end

    # Configures data labels for this chart.
    #
    # @param args [Array] Positional arguments.
    # @param kwargs [Hash] Keyword arguments.
    # @return [Object] The configured data_labels property.
    # @api public
    #: (*(Hash[Symbol, String | Numeric | bool | nil] | String) args, **String | Numeric | bool | nil kwargs) -> untyped
    def data_labels(*args, **kwargs)
      if !args.empty?
        warn "[Xlsxrb DEPRECATION] ChartBuilder#data_labels(options) with positional argument is deprecated; use #data_labels= or #data_labels(**kwargs) instead.", uplevel: 1
        @options[:data_labels] = args.first
      elsif !kwargs.empty?
        @options[:data_labels] = kwargs
        nil
      else
        @options[:data_labels]
      end
    end

    # Sets data labels.
    #
    # @param value [Object]
    # @return [Object]
    # @api public
    #: (untyped value) -> untyped
    def data_labels=(value)
      @options[:data_labels] = value
    end

    # Configures whether to plot visible cells only.
    #
    # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#plot_visible_only=} instead.
    # @param value [Boolean, String, nil]
    # @return [Boolean, String]
    # @api public
    #: (?untyped value) -> untyped
    def plot_visible_only(value = nil)
      if value.nil?
        @options[:plot_visible_only]
      else
        warn "[Xlsxrb DEPRECATION] ChartBuilder#plot_visible_only(value) is deprecated; use #plot_visible_only= instead.", uplevel: 1
        @options[:plot_visible_only] = value
      end
    end

    # Sets whether to plot visible cells only.
    #
    # @param value [Boolean, nil]
    # @return [Boolean, nil]
    # @api public
    #: (untyped value) -> untyped
    def plot_visible_only=(value)
      @options[:plot_visible_only] = value
    end

    # Configures how blank cells are displayed in the chart.
    #
    # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#display_blanks_as=} instead.
    # @param value [String, nil]
    # @return [String]
    # @api public
    #: (?untyped value) -> untyped
    def display_blanks_as(value = nil)
      if value.nil?
        @options[:display_blanks_as]
      else
        warn "[Xlsxrb DEPRECATION] ChartBuilder#display_blanks_as(value) is deprecated; use #display_blanks_as= instead.", uplevel: 1
        @options[:display_blanks_as] = value
      end
    end

    # Sets display blanks as mode.
    #
    # @param value [String, nil]
    # @return [String, nil]
    # @api public
    #: (untyped value) -> untyped
    def display_blanks_as=(value)
      @options[:display_blanks_as] = value
    end

    # Configures 3D view properties for 3D charts.
    #
    # @param args [Array] Positional arguments.
    # @param kwargs [Hash] Keyword arguments.
    # @return [Object]
    # @api public
    #: (*(Hash[Symbol, String | Numeric | bool | nil] | String) args, **String | Numeric | bool | nil kwargs) -> untyped
    def view3d(*args, **kwargs)
      if !args.empty?
        warn "[Xlsxrb DEPRECATION] ChartBuilder#view3d(options) with positional argument is deprecated; use #view3d= or #view3d(**kwargs) instead.", uplevel: 1
        @options[:view3d] = args.first
      elsif !kwargs.empty?
        @options[:view3d] = kwargs
        nil
      else
        @options[:view3d]
      end
    end

    # Sets 3D view properties.
    #
    # @param value [Object]
    # @return [Object]
    # @api public
    #: (untyped value) -> untyped
    def view3d=(value)
      @options[:view3d] = value
    end

    # Configures the category axis properties for this chart.
    #
    # @param args [Array] Positional arguments.
    # @param kwargs [Hash] Keyword arguments.
    # @return [Object] The configured category axis property.
    # @api public
    #: (*(Hash[Symbol, String | Numeric | bool | nil] | String) args, **String | Numeric | bool | nil kwargs) -> untyped
    def category_axis(*args, **kwargs)
      if !args.empty?
        warn "[Xlsxrb DEPRECATION] ChartBuilder#category_axis(options) with positional argument is deprecated; use #category_axis= or #category_axis(**kwargs) instead.", uplevel: 1
        @options[:category_axis] = args.first
      elsif !kwargs.empty?
        @options[:category_axis] = kwargs
        nil
      else
        @options[:category_axis]
      end
    end

    # Sets category axis properties.
    #
    # @param value [Object]
    # @return [Object]
    # @api public
    #: (untyped value) -> untyped
    def category_axis=(value)
      @options[:category_axis] = value
    end

    # Configures the value axis properties for this chart.
    #
    # @param args [Array] Positional arguments.
    # @param kwargs [Hash] Keyword arguments.
    # @return [Object] The configured value axis property.
    # @api public
    #: (*(Hash[Symbol, String | Numeric | bool | nil] | String) args, **String | Numeric | bool | nil kwargs) -> untyped
    def value_axis(*args, **kwargs)
      if !args.empty?
        warn "[Xlsxrb DEPRECATION] ChartBuilder#value_axis(options) with positional argument is deprecated; use #value_axis= or #value_axis(**kwargs) instead.", uplevel: 1
        @options[:value_axis] = args.first
      elsif !kwargs.empty?
        @options[:value_axis] = kwargs
        nil
      else
        @options[:value_axis]
      end
    end

    # Sets value axis properties.
    #
    # @param value [Object]
    # @return [Object]
    # @api public
    #: (untyped value) -> untyped
    def value_axis=(value)
      @options[:value_axis] = value
    end

    # Configures the secondary value axis properties for combo charts.
    #
    # @param args [Array] Positional arguments.
    # @param kwargs [Hash] Keyword arguments.
    # @return [Object]
    # @api public
    #: (*(Hash[Symbol, untyped] | String | bool) args, **untyped kwargs) -> untyped
    def secondary_value_axis(*args, **kwargs)
      if !args.empty?
        warn "[Xlsxrb DEPRECATION] ChartBuilder#secondary_value_axis(options) with positional argument is deprecated; use #secondary_value_axis= or #secondary_value_axis(**kwargs) instead.", uplevel: 1
        @options[:secondary_value_axis] = args.first
      elsif !kwargs.empty?
        @options[:secondary_value_axis] = kwargs
        nil
      else
        @options[:secondary_value_axis]
      end
    end

    # Sets secondary value axis properties.
    #
    # @param value [Object]
    # @return [Object]
    # @api public
    #: (untyped value) -> untyped
    def secondary_value_axis=(value)
      @options[:secondary_value_axis] = value
    end

    # @deprecated Use {#secondary_value_axis} instead.
    #: (*(Hash[Symbol, untyped] | String | bool) args, **untyped kwargs) -> untyped
    def secondary_axis(*args, **kwargs)
      warn "[DEPRECATION] `secondary_axis` is deprecated; use `secondary_value_axis` instead.", uplevel: 1
      secondary_value_axis(*args, **kwargs)
    end

    # @deprecated Use {#secondary_value_axis=} instead.
    #: (untyped value) -> untyped
    def secondary_axis=(value)
      warn "[DEPRECATION] `secondary_axis=` is deprecated; use `secondary_value_axis=` instead.", uplevel: 1
      self.secondary_value_axis = value
    end

    # Configures whether to show the legend key in data labels.
    #
    # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#show_legend_key=} instead.
    # @param value [Boolean, String, nil]
    # @return [Boolean, String]
    # @api public
    #: (?untyped value) -> untyped
    def show_legend_key(value = nil)
      if value.nil?
        @options[:show_legend_key]
      else
        warn "[Xlsxrb DEPRECATION] ChartBuilder#show_legend_key(value) is deprecated; use #show_legend_key= instead.", uplevel: 1
        @options[:show_legend_key] = value
      end
    end

    # Sets whether to show the legend key in data labels.
    #
    # @param value [Boolean, nil]
    # @return [Boolean, nil]
    # @api public
    #: (untyped value) -> untyped
    def show_legend_key=(value)
      @options[:show_legend_key] = value
    end

    # Builder for a single series entry in block-style chart definitions.
    #
    # @api public
    class SeriesBuilder
      #: () -> void
      def initialize
        @options = {}
      end

      # @return [Hash{Symbol => Object}]
      #: Hash[Symbol, untyped]
      attr_reader :options

      # Configures categories reference for this series (e.g. "Sheet1!$A$2:$A$10").
      #
      # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#categories=} instead.
      # @param value [Object, nil]
      # @return [Object]
      # @api public
      #: (?untyped value) -> untyped
      def categories(value = nil)
        if value.nil?
          @options[:categories]
        else
          warn "[Xlsxrb DEPRECATION] SeriesBuilder#categories(value) is deprecated; use #categories= instead.", uplevel: 1
          @options[:categories] = value
        end
      end

      # Sets categories reference for this series.
      #
      # @param value [Object]
      # @return [Object]
      # @api public
      #: (untyped value) -> untyped
      def categories=(value)
        @options[:categories] = value
      end

      # Configures values reference for this series (e.g. "Sheet1!$B$2:$B$10").
      #
      # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#values=} instead.
      # @param value [Object, nil]
      # @return [Object]
      # @api public
      #: (?untyped value) -> untyped
      def values(value = nil)
        if value.nil?
          @options[:values]
        else
          warn "[Xlsxrb DEPRECATION] SeriesBuilder#values(value) is deprecated; use #values= instead.", uplevel: 1
          @options[:values] = value
        end
      end

      # Sets values reference for this series.
      #
      # @param value [Object]
      # @return [Object]
      # @api public
      #: (untyped value) -> untyped
      def values=(value)
        @options[:values] = value
      end

      # Configures the name for this series (e.g. "Total Sales" or cell reference "Sheet1!$B$1").
      #
      # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#name=} instead.
      # @param value [Object, nil]
      # @return [Object]
      # @api public
      #: (?untyped value) -> untyped
      def name(value = nil)
        if value.nil?
          @options[:name]
        else
          warn "[Xlsxrb DEPRECATION] SeriesBuilder#name(value) is deprecated; use #name= instead.", uplevel: 1
          @options[:name] = value
        end
      end

      # Sets name for this series.
      #
      # @param value [Object]
      # @return [Object]
      # @api public
      #: (untyped value) -> untyped
      def name=(value)
        @options[:name] = value
      end

      # Configures data markers for this series.
      #
      # @param args [Array] Positional arguments.
      # @param kwargs [Hash] Keyword arguments.
      # @return [Object]
      # @api public
      #: (*untyped args, **untyped kwargs) -> untyped
      def marker(*args, **kwargs)
        if !args.empty?
          warn "[Xlsxrb DEPRECATION] SeriesBuilder#marker(options) with positional argument is deprecated; use #marker= or #marker(**kwargs) instead.", uplevel: 1
          @options[:marker] = args.first
        elsif !kwargs.empty?
          @options[:marker] = kwargs
          nil
        else
          @options[:marker]
        end
      end

      # Sets marker for this series.
      #
      # @param value [Object]
      # @return [Object]
      # @api public
      #: (untyped value) -> untyped
      def marker=(value)
        @options[:marker] = value
      end

      # Configures fill color/pattern for this series.
      #
      # @param args [Array] Positional arguments.
      # @param kwargs [Hash] Keyword arguments.
      # @return [Object]
      # @api public
      #: (*(Hash[Symbol, String | Numeric | bool | nil] | String) args, **String | Numeric | bool | nil kwargs) -> untyped
      def fill(*args, **kwargs)
        if !args.empty?
          warn "[Xlsxrb DEPRECATION] SeriesBuilder#fill(options) with positional argument is deprecated; use #fill= or #fill(**kwargs) instead.", uplevel: 1
          @options[:fill] = args.first
        elsif !kwargs.empty?
          @options[:fill] = kwargs
          nil
        else
          @options[:fill]
        end
      end

      # Sets fill for this series.
      #
      # @param value [Object]
      # @return [Object]
      # @api public
      #: (untyped value) -> untyped
      def fill=(value)
        @options[:fill] = value
      end

      # Configures line color/thickness for this series.
      #
      # @param args [Array] Positional arguments.
      # @param kwargs [Hash] Keyword arguments.
      # @return [Object]
      # @api public
      #: (*(Hash[Symbol, String | Numeric | bool | nil] | String) args, **String | Numeric | bool | nil kwargs) -> untyped
      def line(*args, **kwargs)
        if !args.empty?
          warn "[Xlsxrb DEPRECATION] SeriesBuilder#line(options) with positional argument is deprecated; use #line= or #line(**kwargs) instead.", uplevel: 1
          @options[:line] = args.first
        elsif !kwargs.empty?
          @options[:line] = kwargs
          nil
        else
          @options[:line]
        end
      end

      # Sets line for this series.
      #
      # @param value [Object]
      # @return [Object]
      # @api public
      #: (untyped value) -> untyped
      def line=(value)
        @options[:line] = value
      end

      # Configures trendline for this series.
      #
      # @param args [Array] Positional arguments.
      # @param kwargs [Hash] Keyword arguments.
      # @return [Object]
      # @api public
      #: (*(Hash[Symbol, String | Numeric | bool | nil] | String) args, **String | Numeric | bool | nil kwargs) -> untyped
      def trendline(*args, **kwargs)
        if !args.empty?
          warn "[Xlsxrb DEPRECATION] SeriesBuilder#trendline(options) with positional argument is deprecated; use #trendline= or #trendline(**kwargs) instead.", uplevel: 1
          @options[:trendline] = args.first
        elsif !kwargs.empty?
          @options[:trendline] = kwargs
          nil
        else
          @options[:trendline]
        end
      end

      # Sets trendline for this series.
      #
      # @param value [Object]
      # @return [Object]
      # @api public
      #: (untyped value) -> untyped
      def trendline=(value)
        @options[:trendline] = value
      end

      # Configures data labels for this series.
      #
      # @param args [Array] Positional arguments.
      # @param kwargs [Hash] Keyword arguments.
      # @return [Object]
      # @api public
      #: (*(Hash[Symbol, String | Numeric | bool | nil] | String) args, **String | Numeric | bool | nil kwargs) -> untyped
      def data_labels(*args, **kwargs)
        if !args.empty?
          warn "[Xlsxrb DEPRECATION] SeriesBuilder#data_labels(options) with positional argument is deprecated; use #data_labels= or #data_labels(**kwargs) instead.", uplevel: 1
          @options[:data_labels] = args.first
        elsif !kwargs.empty?
          @options[:data_labels] = kwargs
          nil
        else
          @options[:data_labels]
        end
      end

      # Sets data labels for this series.
      #
      # @param value [Object]
      # @return [Object]
      # @api public
      #: (untyped value) -> untyped
      def data_labels=(value)
        @options[:data_labels] = value
      end

      # Configures line smoothing for this series.
      #
      # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#smooth=} instead.
      # @param value [Boolean, nil]
      # @return [Boolean, String]
      # @api public
      #: (?untyped value) -> untyped
      def smooth(value = nil)
        if value.nil?
          @options[:smooth]
        else
          warn "[Xlsxrb DEPRECATION] SeriesBuilder#smooth(value) is deprecated; use #smooth= instead.", uplevel: 1
          @options[:smooth] = value
        end
      end

      # Sets line smoothing for this series.
      #
      # @param value [Boolean, nil]
      # @return [Boolean, nil]
      # @api public
      #: (untyped value) -> untyped
      def smooth=(value)
        @options[:smooth] = value
      end

      # Configures 3D bar/column shape (e.g. "cylinder", "cone", "pyramid").
      #
      # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#shape=} instead.
      # @param value [String, nil]
      # @return [String]
      # @api public
      #: (?untyped value) -> untyped
      def shape(value = nil)
        if value.nil?
          @options[:shape]
        else
          warn "[Xlsxrb DEPRECATION] SeriesBuilder#shape(value) is deprecated; use #shape= instead.", uplevel: 1
          @options[:shape] = value
        end
      end

      # Sets 3D bar/column shape for this series.
      #
      # @param value [String, nil]
      # @return [String, nil]
      # @api public
      #: (untyped value) -> untyped
      def shape=(value)
        @options[:shape] = value
      end

      # Configures individual series chart type override in combo charts.
      #
      # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#type=} instead.
      # @param value [Symbol, String, nil]
      # @return [Symbol, String, Hash, nil]
      # @api public
      #: (?untyped value) -> untyped
      def type(value = nil)
        if value.nil?
          @options[:type]
        else
          warn "[Xlsxrb DEPRECATION] SeriesBuilder#type(value) is deprecated; use #type= instead.", uplevel: 1
          @options[:type] = value
        end
      end

      # Sets individual series chart type override.
      #
      # @param value [Symbol, String, nil]
      # @return [Symbol, String, nil]
      # @api public
      #: (untyped value) -> untyped
      def type=(value)
        @options[:type] = value
      end

      # Configures whether this series uses the secondary value axis.
      #
      # @deprecated Use {#axis} with :secondary instead.
      # @param args [Array] Positional arguments.
      # @param kwargs [Hash] Keyword arguments.
      # @return [Boolean, Hash]
      # @api public
      #: (*bool args, **bool kwargs) -> untyped
      def secondary_axis(*args, **kwargs)
        warn "[DEPRECATION] `secondary_axis` is deprecated; use `axis(:secondary)` instead.", uplevel: 1
        return @options[:secondary_axis] if args.empty? && kwargs.empty?

        val = if kwargs.empty?
                args.empty? || args.first
              else
                kwargs
              end
        self.axis = val ? :secondary : :primary
      end

      # Sets whether this series uses the secondary value axis.
      #
      # @deprecated Use {#axis=} with :secondary instead.
      # @param value [Boolean, nil]
      # @return [Boolean, nil]
      # @api public
      #: (untyped value) -> untyped
      def secondary_axis=(value)
        warn "[DEPRECATION] `secondary_axis=` is deprecated; use `self.axis = :secondary` instead.", uplevel: 1
        self.axis = if value
                      :secondary
                    else
                      :primary
                    end
      end

      # Configures the axis (:primary or :secondary) for this series.
      #
      # @deprecated Calling with arguments is deprecated and will be removed in 0.2.0. Use {#axis=} instead.
      # @param value [Symbol, String, nil]
      # @return [Symbol, String, nil]
      # @api public
      #: (?Symbol | String | nil value) -> (Symbol | String | nil)
      def axis(value = nil)
        return @options[:axis] if value.nil?

        warn "[Xlsxrb DEPRECATION] SeriesBuilder#axis(value) is deprecated; use #axis= instead.", uplevel: 1
        @options[:axis] = value
        @options[:secondary_axis] = (value.to_s == "secondary")
        value
      end

      # Sets the axis for this series.
      #
      # @param value [Symbol, String, nil]
      # @return [Symbol, String, nil]
      # @api public
      #: (Symbol | String | nil value) -> (Symbol | String | nil)
      def axis=(value)
        @options[:axis] = value
        @options[:secondary_axis] = (value.to_s == "secondary")
        value
      end
    end
  end
end
