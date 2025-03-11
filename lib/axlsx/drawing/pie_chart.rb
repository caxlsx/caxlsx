# frozen_string_literal: true

module Axlsx
  # The PieChart is a pie chart that you can add to your worksheet.
  # @see Worksheet#add_chart
  # @see Chart#add_series
  # @see README for an example
  class PieChart < Chart
    # Creates a new pie chart object
    # @param [GraphicFrame] frame The workbook that owns this chart.
    # @option options [Cell, String] title
    # @option options [Boolean] show_legend
    # @option options [Symbol] grouping
    # @option options [String] gap_depth
    # @see Chart
    def initialize(frame, options = {})
      @vary_colors = true
      super
      @series_type = PieSeries
      @d_lbls = nil
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = +'')
      super do
        str << '<c:pieChart>'
        str << '<c:varyColors val="' << vary_colors.to_s << '"/>'
        @series.each { |ser| ser.to_xml_string(str) }
        d_lbls.to_xml_string(str) if @d_lbls
        str << '</c:pieChart>'
      end
    end
  end
end
