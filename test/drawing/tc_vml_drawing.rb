# frozen_string_literal: true

require 'tc_helper'

class TestVmlDrawing < Minitest::Test
  def setup
    p = Axlsx::Package.new
    wb = p.workbook
    @ws = wb.add_worksheet
    @ws.add_comment ref: 'A1', text: 'penut machine', author: 'crank'
    @ws.add_comment ref: 'C3', text: 'rust bucket', author: 'PO'
    @vml_drawing = @ws.comments.vml_drawing
  end

  def test_initialize
    assert_raises(ArgumentError) { Axlsx::VmlDrawing.new }
  end

  def test_pn
    str = @vml_drawing.pn

    assert_equal("drawings/vmlDrawing1.vml", str)
  end

  def test_to_xml_string
    str = @vml_drawing.to_xml_string
    doc = Nokogiri::XML(str)

    assert_equal(2, doc.xpath("//v:shape").size)
    assert(doc.xpath("//o:idmap[@o:data='#{@ws.index + 1}']"))
  end
end
