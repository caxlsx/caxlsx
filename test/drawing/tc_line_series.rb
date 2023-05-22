# frozen_string_literal: true

require 'tc_helper'

class TestLineSeries < Test::Unit::TestCase
  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet name: "hmmm"
    chart = @ws.add_chart Axlsx::Line3DChart, title: "fishery"
    @series = chart.add_series(
      data: [0, 1, 2],
      labels: ["zero", "one", "two"],
      title: "bob",
      color: "#FF0000",
      show_marker: true,
      smooth: true
    )
  end

  def test_initialize
    assert_equal("bob", @series.title.text, "series title has been applied")
    assert_equal(@series.labels.class, Axlsx::AxDataSource)
    assert_equal(@series.data.class, Axlsx::NumDataSource)
  end

  def test_show_marker
    assert(@series.show_marker)
    @series.show_marker = false

    refute(@series.show_marker)
  end

  def test_smooth
    assert(@series.smooth)
    @series.smooth = false

    refute(@series.smooth)
  end

  def test_marker_symbol
    assert_equal(:default, @series.marker_symbol)
    @series.marker_symbol = :circle

    assert_equal(:circle, @series.marker_symbol)
  end

  def test_to_xml_string
    doc = Nokogiri::XML(wrap_with_namespaces(@series))

    assert(doc.xpath("//srgbClr[@val='#{@series.color}']"))
    assert_equal(0, xpath_with_namespaces(doc, "//c:marker").size)
    assert(doc.xpath("//smooth"))

    @series.marker_symbol = :diamond
    doc = Nokogiri::XML(wrap_with_namespaces(@series))

    assert_equal(1, xpath_with_namespaces(doc, "//c:marker/c:symbol[@val='diamond']").size)

    @series.show_marker = false
    doc = Nokogiri::XML(wrap_with_namespaces(@series))

    assert_equal(1, xpath_with_namespaces(doc, "//c:marker/c:symbol[@val='none']").size)
  end

  def wrap_with_namespaces(series)
    +'<c:chartSpace xmlns:c="' <<
      Axlsx::XML_NS_C <<
      '" xmlns:a="' <<
      Axlsx::XML_NS_A <<
      '">' <<
      series.to_xml_string <<
      '</c:chartSpace>'
  end

  def xpath_with_namespaces(doc, xpath)
    doc.xpath(xpath, "a" => Axlsx::XML_NS_A, "c" => Axlsx::XML_NS_C)
  end
end
