# frozen_string_literal: true

module Axlsx
  autoload :DLbls,            File.expand_path('d_lbls', __dir__)
  autoload :Title,            File.expand_path('title', __dir__)
  autoload :SeriesTitle,      File.expand_path('series_title', __dir__)
  autoload :Series,           File.expand_path('series', __dir__)
  autoload :PieSeries,        File.expand_path('pie_series', __dir__)
  autoload :BarSeries,        File.expand_path('bar_series', __dir__)
  autoload :LineSeries,       File.expand_path('line_series', __dir__)
  autoload :ScatterSeries,    File.expand_path('scatter_series', __dir__)
  autoload :BubbleSeries,     File.expand_path('bubble_series', __dir__)
  autoload :AreaSeries,       File.expand_path('area_series', __dir__)

  autoload :Scaling,          File.expand_path('scaling', __dir__)
  autoload :Axis,             File.expand_path('axis', __dir__)

  autoload :StrVal,           File.expand_path('str_val', __dir__)
  autoload :NumVal,           File.expand_path('num_val', __dir__)
  autoload :StrData,          File.expand_path('str_data', __dir__)
  autoload :NumData,          File.expand_path('num_data', __dir__)
  autoload :NumDataSource,    File.expand_path('num_data_source', __dir__)
  autoload :AxDataSource,     File.expand_path('ax_data_source', __dir__)

  autoload :SerAxis,          File.expand_path('ser_axis', __dir__)
  autoload :CatAxis,          File.expand_path('cat_axis', __dir__)
  autoload :ValAxis,          File.expand_path('val_axis', __dir__)
  autoload :Axes,             File.expand_path('axes', __dir__)

  autoload :Marker,           File.expand_path('marker', __dir__)

  autoload :OneCellAnchor,    File.expand_path('one_cell_anchor', __dir__)
  autoload :TwoCellAnchor,    File.expand_path('two_cell_anchor', __dir__)
  autoload :GraphicFrame,     File.expand_path('graphic_frame', __dir__)

  autoload :View3D,           File.expand_path('view_3D', __dir__)
  autoload :Chart,            File.expand_path('chart', __dir__)
  autoload :Pie3DChart,       File.expand_path('pie_3D_chart', __dir__)
  autoload :PieChart,         File.expand_path('pie_chart', __dir__)
  autoload :Bar3DChart,       File.expand_path('bar_3D_chart', __dir__)
  autoload :BarChart,         File.expand_path('bar_chart', __dir__)
  autoload :LineChart,        File.expand_path('line_chart', __dir__)
  autoload :Line3DChart,      File.expand_path('line_3D_chart', __dir__)
  autoload :ScatterChart,     File.expand_path('scatter_chart', __dir__)
  autoload :BubbleChart,      File.expand_path('bubble_chart', __dir__)
  autoload :AreaChart,        File.expand_path('area_chart', __dir__)

  autoload :PictureLocking,   File.expand_path('picture_locking', __dir__)
  autoload :Pic,              File.expand_path('pic', __dir__)
  autoload :Hyperlink,        File.expand_path('hyperlink', __dir__)

  autoload :VmlDrawing,       File.expand_path('vml_drawing', __dir__)
  autoload :VmlShape,         File.expand_path('vml_shape', __dir__)

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
