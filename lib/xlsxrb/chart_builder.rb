# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  # Builder for block-style chart definitions.
  #
  # @example Create a chart definition
  #   builder = ChartBuilder.new
  #   builder.type = :bar
  #   builder.title = "Quarterly Sales"
  #   builder.series do |s|
  #     s.name = "2026"
  #     s.categories = "Sheet1!$A$2:$A$5"
  #     s.values = "Sheet1!$B$2:$B$5"
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

    # Gets the chart type (e.g. :line, :bar, :pie, :area, :radar, :scatter).
    #
    # @return [Symbol, String, nil]
    # @api public
    #: () -> (Symbol | String | nil)
    def type
      @options[:type]
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

    # Gets or configures the chart title.
    #
    # @overload title
    #   Returns the configured title options or text.
    #   @return [String, Hash{Symbol => Object}, nil]
    # @overload title(**kwargs)
    #   Configures title using keyword arguments.
    #   @return [void]
    # @api public
    #: (**untyped kwargs) -> (String | Hash[Symbol, untyped] | nil)
    def title(**kwargs)
      if kwargs.empty?
        @options[:title]
      else
        @options[:title] = kwargs
        nil
      end
    end

    # Sets the chart title.
    #
    # @param value [String, Hash{Symbol => Object}, nil]
    # @return [String, Hash{Symbol => Object}, nil]
    # @api public
    #: (String | Hash[Symbol, untyped] | nil value) -> (String | Hash[Symbol, untyped] | nil)
    def title=(value)
      @options[:title] = value
    end

    # Adds a data series to the chart or returns the configured series list.
    #
    # @overload series
    #   Returns the configured series options list.
    #   @return [Array<Hash{Symbol => Object}>]
    # @overload series(**kwargs)
    #   Adds a series configured via keyword arguments.
    #   @return [void]
    # @overload series(&block)
    #   Configures a series using a block.
    #   @yield [series_builder]
    #   @yieldparam series_builder [SeriesBuilder]
    #   @return [void]
    # @api public
    #: (**untyped kwargs) ?{ (SeriesBuilder) -> void } -> (Array[Hash[Symbol, untyped]] | nil)
    def series(**kwargs, &)
      if block_given?
        sb = SeriesBuilder.new
        yield sb
        (@options[:series] ||= []) << sb.options
        nil
      elsif !kwargs.empty?
        (@options[:series] ||= []) << kwargs
        nil
      else
        @options[:series] || []
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

    # Gets or configures the legend property for this chart.
    #
    # @overload legend
    #   Returns the configured legend property.
    #   @return [Object]
    # @overload legend(**kwargs)
    #   Configures legend using keyword arguments.
    #   @return [void]
    # @api public
    #: (**untyped kwargs) -> untyped
    def legend(**kwargs)
      if kwargs.empty?
        @options[:legend]
      else
        @options[:legend] = kwargs
        nil
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

    # Gets or configures the plot_area property for this chart.
    #
    # @overload plot_area
    #   Returns the configured plot_area property.
    #   @return [Object]
    # @overload plot_area(**kwargs)
    #   Configures plot_area using keyword arguments.
    #   @return [void]
    # @api public
    #: (**untyped kwargs) -> untyped
    def plot_area(**kwargs)
      if kwargs.empty?
        @options[:plot_area]
      else
        @options[:plot_area] = kwargs
        nil
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

    # Gets or configures the chart_space property for this chart.
    #
    # @overload chart_space
    #   Returns the configured chart_space property.
    #   @return [Object]
    # @overload chart_space(**kwargs)
    #   Configures chart_space using keyword arguments.
    #   @return [void]
    # @api public
    #: (**untyped kwargs) -> untyped
    def chart_space(**kwargs)
      if kwargs.empty?
        @options[:chart_space]
      else
        @options[:chart_space] = kwargs
        nil
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

    # Gets the style index/id for this chart.
    #
    # @return [Integer, String, Hash, nil] The configured style property.
    # @api public
    #: () -> untyped
    def style
      @options[:style]
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

    # Gets or configures data labels for this chart.
    #
    # @overload data_labels
    #   Returns the configured data_labels property.
    #   @return [Object]
    # @overload data_labels(**kwargs)
    #   Configures data_labels using keyword arguments.
    #   @return [void]
    # @api public
    #: (**untyped kwargs) -> untyped
    def data_labels(**kwargs)
      if kwargs.empty?
        @options[:data_labels]
      else
        @options[:data_labels] = kwargs
        nil
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

    # Gets whether to plot visible cells only.
    #
    # @return [Boolean, nil]
    # @api public
    #: () -> bool?
    def plot_visible_only
      @options[:plot_visible_only]
    end

    # Sets whether to plot visible cells only.
    #
    # @param value [Boolean, nil]
    # @return [Boolean, nil]
    # @api public
    #: (bool? value) -> bool?
    def plot_visible_only=(value)
      @options[:plot_visible_only] = value
    end

    # Gets how blank cells are displayed in the chart.
    #
    # @return [String, nil]
    # @api public
    #: () -> String?
    def display_blanks_as
      @options[:display_blanks_as]
    end

    # Sets display blanks as mode.
    #
    # @param value [String, nil]
    # @return [String, nil]
    # @api public
    #: (String? value) -> String?
    def display_blanks_as=(value)
      @options[:display_blanks_as] = value
    end

    # Gets or configures 3D view properties for 3D charts.
    #
    # @overload view3d
    #   Returns the configured view3d property.
    #   @return [Object]
    # @overload view3d(**kwargs)
    #   Configures view3d using keyword arguments.
    #   @return [void]
    # @api public
    #: (**untyped kwargs) -> untyped
    def view3d(**kwargs)
      if kwargs.empty?
        @options[:view3d]
      else
        @options[:view3d] = kwargs
        nil
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

    # Gets or configures the category axis properties for this chart.
    #
    # @overload category_axis
    #   Returns the configured category axis property.
    #   @return [Object]
    # @overload category_axis(**kwargs)
    #   Configures category axis using keyword arguments.
    #   @return [void]
    # @api public
    #: (**untyped kwargs) -> untyped
    def category_axis(**kwargs)
      if kwargs.empty?
        @options[:category_axis]
      else
        @options[:category_axis] = kwargs
        nil
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

    # Gets or configures the value axis properties for this chart.
    #
    # @overload value_axis
    #   Returns the configured value axis property.
    #   @return [Object]
    # @overload value_axis(**kwargs)
    #   Configures value axis using keyword arguments.
    #   @return [void]
    # @api public
    #: (**untyped kwargs) -> untyped
    def value_axis(**kwargs)
      if kwargs.empty?
        @options[:value_axis]
      else
        @options[:value_axis] = kwargs
        nil
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

    # Gets or configures the secondary value axis properties for combo charts.
    #
    # @overload secondary_value_axis
    #   Returns the configured secondary value axis property.
    #   @return [Object]
    # @overload secondary_value_axis(**kwargs)
    #   Configures secondary value axis using keyword arguments.
    #   @return [void]
    # @api public
    #: (**untyped kwargs) -> untyped
    def secondary_value_axis(**kwargs)
      if kwargs.empty?
        @options[:secondary_value_axis]
      else
        @options[:secondary_value_axis] = kwargs
        nil
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

    # Gets whether to show the legend key in data labels.
    #
    # @return [Boolean, nil]
    # @api public
    #: () -> bool?
    def show_legend_key
      @options[:show_legend_key]
    end

    # Sets whether to show the legend key in data labels.
    #
    # @param value [Boolean, nil]
    # @return [Boolean, nil]
    # @api public
    #: (bool? value) -> bool?
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

      # Gets categories reference for this series (e.g. "Sheet1!$A$2:$A$10").
      #
      # @return [Object]
      # @api public
      #: () -> untyped
      def categories
        @options[:categories]
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

      # Gets values reference for this series (e.g. "Sheet1!$B$2:$B$10").
      #
      # @return [Object]
      # @api public
      #: () -> untyped
      def values
        @options[:values]
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

      # Gets the name for this series (e.g. "Total Sales" or cell reference "Sheet1!$B$1").
      #
      # @return [Object]
      # @api public
      #: () -> untyped
      def name
        @options[:name]
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

      # Gets or configures data markers for this series.
      #
      # @overload marker
      #   Returns the configured marker property.
      #   @return [Object]
      # @overload marker(**kwargs)
      #   Configures marker using keyword arguments.
      #   @return [void]
      # @api public
      #: (**untyped kwargs) -> untyped
      def marker(**kwargs)
        if kwargs.empty?
          @options[:marker]
        else
          @options[:marker] = kwargs
          nil
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

      # Gets or configures fill color/pattern for this series.
      #
      # @overload fill
      #   Returns the configured fill property.
      #   @return [Object]
      # @overload fill(**kwargs)
      #   Configures fill using keyword arguments.
      #   @return [void]
      # @api public
      #: (**untyped kwargs) -> untyped
      def fill(**kwargs)
        if kwargs.empty?
          @options[:fill]
        else
          @options[:fill] = kwargs
          nil
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

      # Gets or configures line color/thickness for this series.
      #
      # @overload line
      #   Returns the configured line property.
      #   @return [Object]
      # @overload line(**kwargs)
      #   Configures line using keyword arguments.
      #   @return [void]
      # @api public
      #: (**untyped kwargs) -> untyped
      def line(**kwargs)
        if kwargs.empty?
          @options[:line]
        else
          @options[:line] = kwargs
          nil
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

      # Gets or configures trendline for this series.
      #
      # @overload trendline
      #   Returns the configured trendline property.
      #   @return [Object]
      # @overload trendline(**kwargs)
      #   Configures trendline using keyword arguments.
      #   @return [void]
      # @api public
      #: (**untyped kwargs) -> untyped
      def trendline(**kwargs)
        if kwargs.empty?
          @options[:trendline]
        else
          @options[:trendline] = kwargs
          nil
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

      # Gets or configures data labels for this series.
      #
      # @overload data_labels
      #   Returns the configured data_labels property.
      #   @return [Object]
      # @overload data_labels(**kwargs)
      #   Configures data_labels using keyword arguments.
      #   @return [void]
      # @api public
      #: (**untyped kwargs) -> untyped
      def data_labels(**kwargs)
        if kwargs.empty?
          @options[:data_labels]
        else
          @options[:data_labels] = kwargs
          nil
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

      # Gets line smoothing for this series.
      #
      # @return [Boolean, nil]
      # @api public
      #: () -> bool?
      def smooth
        @options[:smooth]
      end

      # Sets line smoothing for this series.
      #
      # @param value [Boolean, nil]
      # @return [Boolean, nil]
      # @api public
      #: (bool? value) -> bool?
      def smooth=(value)
        @options[:smooth] = value
      end

      # Gets 3D bar/column shape (e.g. "cylinder", "cone", "pyramid").
      #
      # @return [String, nil]
      # @api public
      #: () -> String?
      def shape
        @options[:shape]
      end

      # Sets 3D bar/column shape for this series.
      #
      # @param value [String, nil]
      # @return [String, nil]
      # @api public
      #: (String? value) -> String?
      def shape=(value)
        @options[:shape] = value
      end

      # Gets individual series chart type override in combo charts.
      #
      # @return [Symbol, String, nil]
      # @api public
      #: () -> (Symbol | String | nil)
      def type
        @options[:type]
      end

      # Sets individual series chart type override.
      #
      # @param value [Symbol, String, nil]
      # @return [Symbol, String, nil]
      # @api public
      #: (Symbol | String | nil value) -> (Symbol | String | nil)
      def type=(value)
        @options[:type] = value
      end

      # Gets the axis (:primary or :secondary) for this series.
      #
      # @return [Symbol, String, nil]
      # @api public
      #: () -> (Symbol | String | nil)
      def axis
        @options[:axis]
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
