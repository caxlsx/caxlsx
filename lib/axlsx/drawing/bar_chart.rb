# frozen_string_literal: true

module Axlsx
  # The BarChart is a two dimensional barchart that you can add to your worksheet.
  # @see Worksheet#add_chart
  # @see Chart#add_series
  # @see Package#serialize
  # @see README for an example
  class BarChart < Chart
    # the category axis
    # @return [CatAxis]
    def cat_axis
      axes[:cat_axis]
    end
    alias :catAxis :cat_axis

    # the value axis
    # @return [ValAxis]
    def val_axis
      axes[:val_axis]
    end
    alias :valAxis :val_axis

    # The direction of the bars in the chart
    # must be one of [:bar, :col]
    # @return [Symbol]
    def bar_dir
      @bar_dir ||= :bar
    end
    alias :barDir :bar_dir

    # space between bar or column clusters, as a percentage of the bar or column width.
    # @return [Integer]
    def gap_width
      @gap_width ||= 150
    end
    alias :gapWidth :gap_width

    # grouping for a column, line, or area chart.
    # must be one of  [:percentStacked, :clustered, :standard, :stacked]
    # @return [Symbol]
    def grouping
      @grouping ||= :clustered
    end

    # Overlap between series
    # @return [Integer]
    def overlap
      @overlap ||= 0
    end

    # The shape of the bars or columns
    # must be one of  [:cone, :coneToMax, :box, :cylinder, :pyramid, :pyramidToMax]
    # @return [Symbol]
    def shape
      @shape ||= :box
    end

    # Creates a new bar chart object
    # @param [GraphicFrame] frame The workbook that owns this chart.
    # @option options [Cell, String] title
    # @option options [Boolean] show_legend
    # @option options [Symbol] bar_dir
    # @option options [Symbol] grouping
    # @option options [String] gap_width
    # @option options [Symbol] shape
    # @see Chart
    def initialize(frame, options = {})
      @vary_colors = true
      @gap_width, @overlap, @shape = nil, nil, nil
      super
      @series_type = BarSeries
      @d_lbls = nil
    end

    # The direction of the bars in the chart
    # must be one of [:bar, :col]
    def bar_dir=(v)
      RestrictionValidator.validate "BarChart.bar_dir", [:bar, :col], v
      @bar_dir = v
    end
    alias :barDir= :bar_dir=

    # grouping for a column, line, or area chart.
    # must be one of  [:percentStacked, :clustered, :standard, :stacked]
    def grouping=(v)
      RestrictionValidator.validate "BarChart.grouping", [:percentStacked, :clustered, :standard, :stacked], v
      @grouping = v
    end

    # space between bar or column clusters, as a percentage of the bar or column width.
    def gap_width=(v)
      RangeValidator.validate "BarChart.gap_width", 0, 500, v
      @gap_width = v
    end
    alias :gapWidth= :gap_width=

    def overlap=(v)
      RangeValidator.validate "BarChart.overlap", -100, 100, v
      @overlap = v
    end

    # The shape of the bars or columns
    # must be one of  [:cone, :coneToMax, :box, :cylinder, :pyramid, :pyramidToMax]
    def shape=(v)
      RestrictionValidator.validate "BarChart.shape", [:cone, :coneToMax, :box, :cylinder, :pyramid, :pyramidToMax], v
      @shape = v
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = +'')
      super do
        str << '<c:barChart>'
        str << '<c:barDir val="' << bar_dir.to_s << '"/>'
        str << '<c:grouping val="' << grouping.to_s << '"/>'
        str << '<c:varyColors val="' << vary_colors.to_s << '"/>'
        @series.each { |ser| ser.to_xml_string(str) }
        @d_lbls.to_xml_string(str) if @d_lbls
        str << '<c:overlap val="' << @overlap.to_s << '"/>' unless @overlap.nil?
        str << '<c:gapWidth val="' << @gap_width.to_s << '"/>' unless @gap_width.nil?
        str << '<c:shape val="' << @shape.to_s << '"/>' unless @shape.nil?
        axes.to_xml_string(str, ids: true)
        str << '</c:barChart>'
        axes.to_xml_string(str)
      end
    end

    # A hash of axes used by this chart. Bar charts have a value and
    # category axes specified via axes[:val_axes] and axes[:cat_axis]
    # @return [Axes]
    def axes
      @axes ||= begin
        a = Axes.new(cat_axis: CatAxis, val_axis: ValAxis)

        if bar_dir == :col
          a[:val_axis].ax_pos = :l
        else
          a[:cat_axis].ax_pos = :l
        end

        a
      end
    end
  end
end
