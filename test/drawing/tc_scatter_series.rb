# frozen_string_literal: true

require 'tc_helper'

class TestScatterSeries < Test::Unit::TestCase
  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet name: "hmmm"
    @chart = @ws.add_chart Axlsx::ScatterChart, title: "Scatter Chart"
    @series = @chart.add_series xData: [1, 2, 4], yData: [1, 3, 9], title: "exponents", color: 'FF0000', smooth: true
  end

  def test_initialize
    assert_equal("exponents", @series.title.text, "series title has been applied")
  end

  def test_smoothed_chart_default_smoothing
    @chart = @ws.add_chart Axlsx::ScatterChart, title: "Smooth Chart", scatter_style: :smoothMarker
    @series = @chart.add_series xData: [1, 2, 4], yData: [1, 3, 9], title: "smoothed exponents"

    assert(@series.smooth, "series is smooth by default on smooth charts")
  end

  def test_unsmoothed_chart_default_smoothing
    @chart = @ws.add_chart Axlsx::ScatterChart, title: "Unsmooth Chart", scatter_style: :line
    @series = @chart.add_series xData: [1, 2, 4], yData: [1, 3, 9], title: "unsmoothed exponents"

    refute(@series.smooth, "series is not smooth by default on non-smooth charts")
  end

  def test_explicit_smoothing
    @chart = @ws.add_chart Axlsx::ScatterChart, title: "Unsmooth Chart, Smooth Series", scatter_style: :line
    @series = @chart.add_series xData: [1, 2, 4], yData: [1, 3, 9], title: "smoothed exponents", smooth: true

    assert(@series.smooth, "series is smooth when overriding chart default")
  end

  def test_explicit_unsmoothing
    @chart = @ws.add_chart Axlsx::ScatterChart, title: "Smooth Chart, Unsmooth Series", scatter_style: :smoothMarker
    @series = @chart.add_series xData: [1, 2, 4], yData: [1, 3, 9], title: "unsmoothed exponents", smooth: false

    refute(@series.smooth, "series is not smooth when overriding chart default")
  end

  def test_ln_width
    @chart = @ws.add_chart Axlsx::ScatterChart, title: "ln width", scatter_style: :line
    @series = @chart.add_series xData: [1, 2, 4], yData: [1, 3, 9], title: "ln_width"
    @series.ln_width = 12_700

    assert_equal(12_700, @series.ln_width, 'line width assigment is allowed')
  end

  def test_to_xml_string
    @chart.scatter_style = :line
    @series.ln_width = 12_700
    doc = Nokogiri::XML(@chart.to_xml_string)

    assert_equal(12_700, @series.ln_width)
    assert_equal(4, doc.xpath("//a:srgbClr[@val='#{@series.color}']").size)
    assert_equal(1, doc.xpath("//a:ln[@w='#{@series.ln_width}']").length)
  end

  def test_false_show_marker
    @chart = @ws.add_chart Axlsx::ScatterChart, title: 'Smooth Chart', scatter_style: :smoothMarker
    @series = @chart.add_series xData: [1, 2, 4], yData: [1, 3, 9]

    assert(@series.show_marker, 'markers are enabled for marker-related styles')
  end

  def test_true_show_marker
    @chart = @ws.add_chart Axlsx::ScatterChart, title: 'Line chart', scatter_style: :line
    @series = @chart.add_series xData: [1, 2, 4], yData: [1, 3, 9]

    refute(@series.show_marker, 'markers are disabled for markerless scatter styles')
  end

  def test_marker_symbol
    @chart = @ws.add_chart Axlsx::ScatterChart, title: 'Line chart', scatter_style: :line
    @series = @chart.add_series xData: [1, 2, 4], yData: [1, 3, 9], marker_symbol: :diamond

    assert_equal(:diamond, @series.marker_symbol, 'series could have own custom marker symbol')
  end
end
