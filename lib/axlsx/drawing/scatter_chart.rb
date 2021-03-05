# encoding: UTF-8

module Axlsx

  # The ScatterChart allows you to insert a scatter chart into your worksheet
  # @see Worksheet#add_chart
  # @see Chart#add_series
  # @see README for an example
  class ScatterChart < Chart

    include Axlsx::OptionsParser

    # The Style for the scatter chart
    # must be one of :none | :line | :lineMarker | :marker | :smooth | :smoothMarker
    # return [Symbol]
    attr_reader :scatter_style
    alias :scatterStyle :scatter_style

    # the x value axis
    # @return [ValAxis]
    def x_val_axis
      axes[:x_val_axis]
    end

    # the y value axis
    # @return [ValAxis]
    def y_val_axis
      axes[:y_val_axis]
    end

    # secondary y value axis
    def secondary_y_val_axis
      secondary_axes[:secondary_y_val_axis]
    end

    # Creates a new scatter chart
    def initialize(frame, options={})
      @vary_colors = 0
      @scatter_style = :lineMarker

           super(frame, options)
      @series_type = ScatterSeries
      @d_lbls = nil
      parse_options options
    end

    # see #scatterStyle
    def scatter_style=(v)
      Axlsx.validate_scatter_style(v)
      @scatter_style = v
    end
    alias :scatterStyle= :scatter_style=

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      axes_set = [axes, secondary_axes]

      super(str) do
        @series.partition(&:on_primary_y_axis).each do |series_set|
          next unless series_set.any?

          str << '<c:scatterChart>'
          str << ('<c:scatterStyle val="' << scatter_style.to_s << '"/>')
          str << ('<c:varyColors val="' << vary_colors.to_s << '"/>')
          series_set.each { |ser| ser.to_xml_string(str) }
          d_lbls.to_xml_string(str) if @d_lbls

          axes_set.shift.to_xml_string(str, ids: true)
          str << '</c:scatterChart>'
        end

        axes.to_xml_string(str)
        secondary_axes.to_xml_string(str) unless @series.all?(&:on_primary_y_axis)
      end
      str
    end

    # The axes for the scatter chart. ScatterChart has an x_val_axis and
    # a y_val_axis
    # @return [Axes]
    def axes
      @axes ||= Axes.new(:x_val_axis => ValAxis, :y_val_axis => ValAxis)
    end

    # Secondary axes for the scatter chart.
    # @return [Axes]
    def secondary_axes
      @secondary_axes ||= Axes.new.tap do |axes|
        axes.add_axis(:secondary_x_val_axis, ValAxis).last.tap do |_, sec_x|
          sec_x.ax_pos = :b
          sec_x.delete = 1
          sec_x.gridlines = false
        end

        axes.add_axis(:secondary_y_val_axis, ValAxis).last.tap do |_, sec_y|
          sec_y.ax_pos = :r
          sec_y.crosses = :max
          sec_y.gridlines = false
        end
      end
    end
  end
end
