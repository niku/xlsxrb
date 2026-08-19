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

    # Sets the chart type (e.g. :line, :bar, :pie, :area, :radar, :scatter).
    #
    # @param value [Symbol, String] Chart type.
    # @return [Symbol, String]
    # @api public
    #: (Symbol | String value) -> (Symbol | String)
    def type(value)
      @options[:type] = value
    end

    # Sets the chart title.
    #
    # @param value [String, Hash, nil] Chart title text or options hash.
    # @return [String, Hash, nil]
    # @api public
    #: (String | Hash[Symbol, untyped] | nil value) -> (String | Hash[Symbol, untyped] | nil)
    def title(value)
      @options[:title] = value
    end

    # Adds a data series to the chart.
    #
    # @overload series(value)
    #   Adds a pre-built series options hash.
    #   @param value [Hash{Symbol => Object}] Series options hash.
    #   @return [Array<Hash{Symbol => Object}>]
    #
    # @overload series(&block)
    #   Configures a series using a block.
    #   @yield [series_builder]
    #   @yieldparam series_builder [SeriesBuilder]
    #   @return [Array<Hash{Symbol => Object}>]
    #
    # @api public
    #: (?Hash[Symbol, untyped]? value) ?{ (SeriesBuilder) -> void } -> Array[Hash[Symbol, untyped]]
    def series(value = nil)
      @options[:series] ||= []
      if block_given?
        sb = SeriesBuilder.new
        yield sb
        @options[:series] << sb.options
      elsif value
        @options[:series] << value
      end
      @options[:series]
    end

    # Configures the legend property for this chart.
    #
    # @param args [Array] Positional arguments (e.g. position string).
    # @param kwargs [Hash] Keyword arguments for legend styling/layout.
    # @return [Object] The configured legend property.
    # @api public
    #: (*(Hash[Symbol, String | Integer | bool | nil] | String) args, **String | Integer | bool | nil kwargs) -> (Hash[Symbol, String | Integer | bool | nil] | String)
    def legend(*args, **kwargs)
      @options[:legend] = kwargs.empty? ? args.first : kwargs
    end

    # Configures the plot_area property for this chart.
    #
    # @param args [Array] Positional arguments.
    # @param kwargs [Hash] Keyword arguments for plot area configuration.
    # @return [Object] The configured plot_area property.
    # @api public
    #: (*(Hash[Symbol, String | Integer | bool | nil] | String) args, **String | Integer | bool | nil kwargs) -> (Hash[Symbol, String | Integer | bool | nil] | String)
    def plot_area(*args, **kwargs)
      @options[:plot_area] = kwargs.empty? ? args.first : kwargs
    end

    # Configures the chart_space property for this chart.
    #
    # @param args [Array] Positional arguments.
    # @param kwargs [Hash] Keyword arguments.
    # @return [Object] The configured chart_space property.
    # @api public
    #: (*(Hash[Symbol, String | Integer | bool | nil] | String) args, **String | Integer | bool | nil kwargs) -> (Hash[Symbol, String | Integer | bool | nil] | String)
    def chart_space(*args, **kwargs)
      @options[:chart_space] = kwargs.empty? ? args.first : kwargs
    end

    # Configures the style index/id for this chart.
    #
    # @param args [Array] Positional arguments.
    # @param kwargs [Hash] Keyword arguments.
    # @return [Object] The configured style property.
    # @api public
    #: (*(Hash[Symbol, String | Integer | bool | nil] | String | Integer) args, **String | Integer | bool | nil kwargs) -> (Hash[Symbol, String | Integer | bool | nil] | String | Integer)
    def style(*args, **kwargs)
      @options[:style] = kwargs.empty? ? args.first : kwargs
    end

    # Configures data labels for this chart.
    #
    # @param args [Array] Positional arguments.
    # @param kwargs [Hash] Keyword arguments.
    # @return [Object] The configured data_labels property.
    # @api public
    #: (*(Hash[Symbol, String | Integer | bool | nil] | String) args, **String | Integer | bool | nil kwargs) -> (Hash[Symbol, String | Integer | bool | nil] | String)
    def data_labels(*args, **kwargs)
      @options[:data_labels] = kwargs.empty? ? args.first : kwargs
    end

    # Configures whether to plot visible cells only.
    #
    # @param args [Array] Positional arguments.
    # @param kwargs [Hash] Keyword arguments.
    # @return [Boolean, String]
    # @api public
    #: (*(bool | String) args, **String | Integer | bool | nil kwargs) -> (bool | String)
    def plot_visible_only(*args, **kwargs)
      @options[:plot_visible_only] = kwargs.empty? ? args.first : kwargs
    end

    # Configures how blank cells are displayed in the chart.
    #
    # @param args [Array] Positional arguments.
    # @param kwargs [Hash] Keyword arguments.
    # @return [String]
    # @api public
    #: (*(String) args, **String | Integer | bool | nil kwargs) -> String
    def display_blanks_as(*args, **kwargs)
      @options[:display_blanks_as] = kwargs.empty? ? args.first : kwargs
    end

    # Configures 3D view properties for 3D charts.
    #
    # @param args [Array] Positional arguments.
    # @param kwargs [Hash] Keyword arguments.
    # @return [Object]
    # @api public
    #: (*(Hash[Symbol, String | Integer | bool | nil] | String) args, **String | Integer | bool | nil kwargs) -> (Hash[Symbol, String | Integer | bool | nil] | String)
    def view3d(*args, **kwargs)
      @options[:view3d] = kwargs.empty? ? args.first : kwargs
    end

    # Configures the category axis properties for this chart.
    #
    # @param args [Array] Positional arguments.
    # @param kwargs [Hash] Keyword arguments.
    # @return [Object]
    # @api public
    #: (*(Hash[Symbol, String | Integer | bool | nil] | String) args, **String | Integer | bool | nil kwargs) -> (Hash[Symbol, String | Integer | bool | nil] | String)
    def category_axis(*args, **kwargs)
      @options[:category_axis] = kwargs.empty? ? args.first : kwargs
    end

    # Configures the value axis properties for this chart.
    #
    # @param args [Array] Positional arguments.
    # @param kwargs [Hash] Keyword arguments.
    # @return [Object]
    # @api public
    #: (*(Hash[Symbol, String | Integer | bool | nil] | String) args, **String | Integer | bool | nil kwargs) -> (Hash[Symbol, String | Integer | bool | nil] | String)
    def value_axis(*args, **kwargs)
      @options[:value_axis] = kwargs.empty? ? args.first : kwargs
    end

    # Configures whether to show the legend key in data labels.
    #
    # @param args [Array] Positional arguments.
    # @param kwargs [Hash] Keyword arguments.
    # @return [Boolean, String]
    # @api public
    #: (*(bool | String) args, **String | Integer | bool | nil kwargs) -> (bool | String)
    def show_legend_key(*args, **kwargs)
      @options[:show_legend_key] = kwargs.empty? ? args.first : kwargs
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
      # @param args [Array] Positional arguments.
      # @param kwargs [Hash] Keyword arguments.
      # @return [Object]
      # @api public
      #: (*untyped args, **untyped kwargs) -> untyped
      def categories(*args, **kwargs)
        @options[:categories] = kwargs.empty? ? args.first : kwargs
      end

      # Configures values reference for this series (e.g. "Sheet1!$B$2:$B$10").
      #
      # @param args [Array] Positional arguments.
      # @param kwargs [Hash] Keyword arguments.
      # @return [Object]
      # @api public
      #: (*untyped args, **untyped kwargs) -> untyped
      def values(*args, **kwargs)
        @options[:values] = kwargs.empty? ? args.first : kwargs
      end

      # Configures the name for this series (e.g. "Total Sales" or cell reference "Sheet1!$B$1").
      #
      # @param args [Array] Positional arguments.
      # @param kwargs [Hash] Keyword arguments.
      # @return [Object]
      # @api public
      #: (*untyped args, **untyped kwargs) -> untyped
      def name(*args, **kwargs)
        @options[:name] = kwargs.empty? ? args.first : kwargs
      end

      # Configures data markers for this series.
      #
      # @param args [Array] Positional arguments.
      # @param kwargs [Hash] Keyword arguments.
      # @return [Object]
      # @api public
      #: (*untyped args, **untyped kwargs) -> untyped
      def marker(*args, **kwargs)
        @options[:marker] = kwargs.empty? ? args.first : kwargs
      end

      # Configures fill color/pattern for this series.
      #
      # @param args [Array] Positional arguments.
      # @param kwargs [Hash] Keyword arguments.
      # @return [Object]
      # @api public
      #: (*(Hash[Symbol, String | Integer | bool | nil] | String) args, **String | Integer | bool | nil kwargs) -> (Hash[Symbol, String | Integer | bool | nil] | String)
      def fill(*args, **kwargs)
        @options[:fill] = kwargs.empty? ? args.first : kwargs
      end

      # Configures line color/thickness for this series.
      #
      # @param args [Array] Positional arguments.
      # @param kwargs [Hash] Keyword arguments.
      # @return [Object]
      # @api public
      #: (*(Hash[Symbol, String | Integer | bool | nil] | String) args, **String | Integer | bool | nil kwargs) -> (Hash[Symbol, String | Integer | bool | nil] | String)
      def line(*args, **kwargs)
        @options[:line] = kwargs.empty? ? args.first : kwargs
      end

      # Configures trendline for this series.
      #
      # @param args [Array] Positional arguments.
      # @param kwargs [Hash] Keyword arguments.
      # @return [Object]
      # @api public
      #: (*(Hash[Symbol, String | Integer | bool | nil] | String) args, **String | Integer | bool | nil kwargs) -> (Hash[Symbol, String | Integer | bool | nil] | String)
      def trendline(*args, **kwargs)
        @options[:trendline] = kwargs.empty? ? args.first : kwargs
      end

      # Configures data labels for this series.
      #
      # @param args [Array] Positional arguments.
      # @param kwargs [Hash] Keyword arguments.
      # @return [Object]
      # @api public
      #: (*(Hash[Symbol, String | Integer | bool | nil] | String) args, **String | Integer | bool | nil kwargs) -> (Hash[Symbol, String | Integer | bool | nil] | String)
      def data_labels(*args, **kwargs)
        @options[:data_labels] = kwargs.empty? ? args.first : kwargs
      end

      # Configures line smoothing for this series.
      #
      # @param args [Array] Positional arguments.
      # @param kwargs [Hash] Keyword arguments.
      # @return [Boolean, String]
      # @api public
      #: (*(bool | String) args, **String | Integer | bool | nil kwargs) -> (bool | String)
      def smooth(*args, **kwargs)
        @options[:smooth] = kwargs.empty? ? args.first : kwargs
      end

      # Configures 3D bar/column shape (e.g. "cylinder", "cone", "pyramid").
      #
      # @param args [Array] Positional arguments.
      # @param kwargs [Hash] Keyword arguments.
      # @return [String]
      # @api public
      #: (*(String) args, **String | Integer | bool | nil kwargs) -> String
      def shape(*args, **kwargs)
        @options[:shape] = kwargs.empty? ? args.first : kwargs
      end

      # Configures individual series chart type override in combo charts.
      #
      # @param args [Array] Positional arguments.
      # @param kwargs [Hash] Keyword arguments.
      # @return [String]
      # @api public
      #: (*(String) args, **String | Integer | bool | nil kwargs) -> String
      def type(*args, **kwargs)
        @options[:type] = kwargs.empty? ? args.first : kwargs
      end
    end
  end
end
