# frozen_string_literal: true

require 'tc_helper'

class TestDrawing < Minitest::Test
  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet
  end

  def test_initialization
    assert_empty(@ws.workbook.drawings)
  end

  def test_add_chart
    chart = @ws.add_chart(Axlsx::Pie3DChart, title: "bob", start_at: [0, 0], end_at: [1, 1])

    assert_kind_of(Axlsx::Pie3DChart, chart, "must create a chart")
    assert_equal(@ws.workbook.charts.last, chart, "must be added to workbook charts collection")
    assert_equal(@ws.drawing.anchors.last.object.chart, chart, "an anchor has been created and holds a reference to this chart")
    anchor = @ws.drawing.anchors.last

    assert_equal([0, 0], [anchor.from.row, anchor.from.col], "options for start at are applied")
    assert_equal([1, 1], [anchor.to.row, anchor.to.col], "options for start at are applied")
    assert_equal("bob", chart.title.text, "option for title is applied")
  end

  def test_add_image
    src = "#{File.dirname(__FILE__)}/../fixtures/image1.jpeg"
    image = @ws.add_image(image_src: src, start_at: [0, 0], width: 600, height: 400)

    assert_kind_of(Axlsx::OneCellAnchor, @ws.drawing.anchors.last)
    assert_kind_of(Axlsx::Pic, image)
    assert_equal(600, image.width)
    assert_equal(400, image.height)
  end

  def test_add_two_cell_anchor_image
    src = "#{File.dirname(__FILE__)}/../fixtures/image1.jpeg"
    image = @ws.add_image(image_src: src, start_at: [0, 0], end_at: [15, 0])

    assert_kind_of(Axlsx::TwoCellAnchor, @ws.drawing.anchors.last)
    assert_kind_of(Axlsx::Pic, image)
  end

  def test_charts
    chart = @ws.add_chart(Axlsx::Pie3DChart, title: "bob", start_at: [0, 0], end_at: [1, 1])

    assert_equal(@ws.drawing.charts.last, chart, "add chart is returned")
    chart = @ws.add_chart(Axlsx::Pie3DChart, title: "nancy", start_at: [1, 5], end_at: [5, 10])

    assert_equal(@ws.drawing.charts.last, chart, "add chart is returned")
  end

  def test_pn
    @ws.add_chart(Axlsx::Pie3DChart)

    assert_equal("drawings/drawing1.xml", @ws.drawing.pn)
  end

  def test_rels_pn
    @ws.add_chart(Axlsx::Pie3DChart)

    assert_equal("drawings/_rels/drawing1.xml.rels", @ws.drawing.rels_pn)
  end

  def test_index
    @ws.add_chart(Axlsx::Pie3DChart)

    assert_equal(@ws.drawing.index, @ws.workbook.drawings.index(@ws.drawing))
  end

  def test_relationships
    @ws.add_chart(Axlsx::Pie3DChart, title: "bob", start_at: [0, 0], end_at: [1, 1])

    assert_equal(1, @ws.drawing.relationships.size, "adding a chart adds a relationship")
    @ws.add_chart(Axlsx::Pie3DChart, title: "nancy", start_at: [1, 5], end_at: [5, 10])

    assert_equal(2, @ws.drawing.relationships.size, "adding a chart adds a relationship")
  end

  def test_to_xml
    schema = Nokogiri::XML::Schema(File.open(Axlsx::DRAWING_XSD))
    @ws.add_chart(Axlsx::Pie3DChart)
    doc = Nokogiri::XML(@ws.drawing.to_xml_string)
    errors = schema.validate(doc)

    assert_empty(errors)
  end
end
