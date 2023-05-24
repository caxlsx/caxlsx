# frozen_string_literal: true

require 'tc_helper'

class TestBarSeries < Test::Unit::TestCase
  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet name: "hmmm"
    @chart = @ws.add_chart Axlsx::Bar3DChart, title: "fishery"
    @series = @chart.add_series(
      data: [0, 1, 2],
      labels: ['zero', 'one', 'two'],
      title: 'bob',
      colors: ['FF0000', '00FF00', '0000FF'],
      shape: :cone,
      series_color: '5A5A5A'
    )
  end

  def test_initialize
    assert_equal("bob", @series.title.text, "series title has been applied")
    assert_equal(@series.data.class, Axlsx::NumDataSource, "data option applied")
    assert_equal(:cone, @series.shape, "series shape has been applied")
    assert_equal('5A5A5A', @series.series_color, 'series color has been applied')
    assert(@series.data.is_a?(Axlsx::NumDataSource))
    assert(@series.labels.is_a?(Axlsx::AxDataSource))
  end

  def test_colors
    assert_equal(3, @series.colors.size)
  end

  def test_shape
    assert_raise(ArgumentError, "require valid shape") { @series.shape = :teardropt }
    assert_nothing_raised("allow valid shape") { @series.shape = :box }
    assert_equal(:box, @series.shape)
  end

  def test_to_xml_string
    doc = Nokogiri::XML(@chart.to_xml_string)
    @series.colors.each_with_index do |_color, index|
      assert_equal(1, doc.xpath("//c:dPt/c:idx[@val='#{index}']").size)
      assert_equal(1, doc.xpath("//c:dPt/c:spPr/a:solidFill/a:srgbClr[@val='#{@series.colors[index]}']").size)
    end
    assert_equal('5A5A5A', doc.xpath('//c:spPr[not(ancestor::c:dPt)]/a:solidFill/a:srgbClr').first.get_attribute('val'), 'series color has been applied')
  end
end
