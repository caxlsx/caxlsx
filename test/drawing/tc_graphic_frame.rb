# frozen_string_literal: true

require 'tc_helper'

class TestGraphicFrame < Minitest::Test
  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet
    @chart = @ws.add_chart Axlsx::Chart
    @frame = @chart.graphic_frame
  end

  def teardown; end

  def test_initialization
    assert_kind_of(Axlsx::TwoCellAnchor, @frame.anchor)
    assert_equal(@frame.chart, @chart)
  end

  def test_rId
    assert_equal @ws.drawing.relationships.for(@chart).Id, @frame.rId
  end

  def test_to_xml_has_correct_rId
    doc = Nokogiri::XML(@frame.to_xml_string)

    assert_equal @frame.rId, doc.xpath("//c:chart", doc.collect_namespaces).first["r:id"]
  end
end
