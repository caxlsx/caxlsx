# frozen_string_literal: true

module Axlsx
  require_relative 'd_lbls'
  require_relative 'title'
  require_relative 'series_title'
  require_relative 'series'
  require_relative 'pie_series'
  require_relative 'bar_series'
  require_relative 'line_series'
  require_relative 'scatter_series'
  require_relative 'bubble_series'
  require_relative 'area_series'

  require_relative 'scaling'
  require_relative 'axis'

  require_relative 'str_val'
  require_relative 'num_val'
  require_relative 'str_data'
  require_relative 'num_data'
  require_relative 'num_data_source'
  require_relative 'ax_data_source'

  require_relative 'ser_axis'
  require_relative 'cat_axis'
  require_relative 'val_axis'
  require_relative 'axes'

  require_relative 'marker'

  require_relative 'one_cell_anchor'
  require_relative 'two_cell_anchor'
  require_relative 'graphic_frame'

  require_relative 'view_3D'
  require_relative 'chart'
  require_relative 'pie_3D_chart'
  require_relative 'pie_chart'
  require_relative 'bar_3D_chart'
  require_relative 'bar_chart'
  require_relative 'line_chart'
  require_relative 'line_3D_chart'
  require_relative 'scatter_chart'
  require_relative 'bubble_chart'
  require_relative 'area_chart'

  require_relative 'picture_locking'
  require_relative 'pic'
  require_relative 'hyperlink'

  require_relative 'vml_drawing'
  require_relative 'vml_shape'

  # A Drawing is a canvas for charts and images. Each worksheet has a single drawing that manages anchors.
  # The anchors reference the charts or images via graphical frames. This is not a trivial relationship so please do follow the advice in the note.
  # @note The recommended way to manage drawings is to use the Worksheet.add_chart and Worksheet.add_image methods.
  # @see Worksheet#add_chart
  # @see Worksheet#add_image
  # @see Chart
  # see examples/example.rb for an example of how to create a chart.
  class Drawing
    # The worksheet that owns the drawing
    # @return [Worksheet]
    attr_reader :worksheet

    # A collection of anchors for this drawing
    # only TwoCellAnchors are supported in this version
    # @return [SimpleTypedList]
    attr_reader :anchors

    # Creates a new Drawing object
    # @param [Worksheet] worksheet The worksheet that owns this drawing
    def initialize(worksheet)
      DataTypeValidator.validate "Drawing.worksheet", Worksheet, worksheet
      @worksheet = worksheet
      @worksheet.workbook.drawings << self
      @anchors = SimpleTypedList.new [TwoCellAnchor, OneCellAnchor]
    end

    # Adds an image to the chart If th end_at option is specified we create a two cell anchor. By default we use a one cell anchor.
    # @note The recommended way to manage images is to use Worksheet.add_image. Please refer to that method for documentation.
    # @see Worksheet#add_image
    # @return [Pic]
    def add_image(options = {})
      if options[:end_at]
        TwoCellAnchor.new(self, options).add_pic(options)
      else
        OneCellAnchor.new(self, options)
      end
      @anchors.last.object
    end

    # Adds a chart to the drawing.
    # @note The recommended way to manage charts is to use Worksheet.add_chart. Please refer to that method for documentation.
    # @see Worksheet#add_chart
    def add_chart(chart_type, options = {})
      TwoCellAnchor.new(self, options)
      @anchors.last.add_chart(chart_type, options)
    end

    # An array of charts that are associated with this drawing's anchors
    # @return [Array]
    def charts
      charts = @anchors.select { |a| a.object.is_a?(GraphicFrame) }
      charts.map { |a| a.object.chart }
    end

    # An array of hyperlink objects associated with this drawings images
    # @return [Array]
    def hyperlinks
      links = images.select { |a| a.hyperlink.is_a?(Hyperlink) }
      links.map(&:hyperlink)
    end

    # An array of image objects that are associated with this drawing's anchors
    # @return [Array]
    def images
      images = @anchors.select { |a| a.object.is_a?(Pic) }
      images.map(&:object)
    end

    # The index of this drawing in the owning workbooks's drawings collection.
    # @return [Integer]
    def index
      @worksheet.workbook.drawings.index(self)
    end

    # The part name for this drawing
    # @return [String]
    def pn
      format(DRAWING_PN, index + 1)
    end

    # The relational part name for this drawing
    # #NOTE This should be rewritten to return an Axlsx::Relationship object.
    # @return [String]
    def rels_pn
      format(DRAWING_RELS_PN, index + 1)
    end

    # A list of objects this drawing holds.
    # @return [Array]
    def child_objects
      charts + images + hyperlinks
    end

    # The drawing's relationships.
    # @return [Relationships]
    def relationships
      r = Relationships.new
      child_objects.each { |child| r << child.relationship }
      r
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = +'')
      str << '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      str << '<xdr:wsDr xmlns:xdr="' << XML_NS_XDR << '" xmlns:a="' << XML_NS_A << '">'
      anchors.each { |anchor| anchor.to_xml_string(str) }
      str << '</xdr:wsDr>'
    end
  end
end
