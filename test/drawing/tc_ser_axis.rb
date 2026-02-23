# frozen_string_literal: true

require 'tc_helper'

class TestSerAxis < Minitest::Test
  def setup
    @axis = Axlsx::SerAxis.new
  end

  def teardown; end

  def test_initialization
    assert_nil(@axis.tick_lbl_skip, "tick_lbl_skip default incorrect")
    assert_nil(@axis.tick_mark_skip, "tick_mark_skip default incorrect")
  end

  def test_options
    a = Axlsx::SerAxis.new(tick_lbl_skip: 9, tick_mark_skip: 7)

    assert_equal(9, a.tick_lbl_skip)
    assert_equal(7, a.tick_mark_skip)
  end

  def test_tick_lbl_skip
    assert_raises(ArgumentError, "requires valid tick_lbl_skip") { @axis.tick_lbl_skip = -1 }
    refute_raises { @axis.tick_lbl_skip = 1 }
    assert_equal(1, @axis.tick_lbl_skip)
    refute_raises { @axis.tick_lbl_skip = nil }
    assert_nil(@axis.tick_lbl_skip)
  end

  def test_tick_mark_skip
    assert_raises(ArgumentError, "requires valid tick_mark_skip") { @axis.tick_mark_skip = :my_eyes }
    refute_raises { @axis.tick_mark_skip = 2 }
    assert_equal(2, @axis.tick_mark_skip)
    refute_raises { @axis.tick_mark_skip = nil }
    assert_nil(@axis.tick_mark_skip)
  end

  def test_to_xml_string_omits_tick_lbl_skip_when_nil
    @axis.cross_axis = Axlsx::ValAxis.new
    str = +'<?xml version="1.0" encoding="UTF-8"?>'
    str << '<c:chartSpace xmlns:c="' << Axlsx::XML_NS_C << '">'
    doc = Nokogiri::XML(@axis.to_xml_string(str))

    assert_empty(doc.xpath('//c:tickLblSkip'))
    assert_empty(doc.xpath('//c:tickMarkSkip'))
  end

  def test_to_xml_string_includes_tick_lbl_skip_when_set
    @axis.tick_lbl_skip = 5
    @axis.tick_mark_skip = 3
    @axis.cross_axis = Axlsx::ValAxis.new
    str = +'<?xml version="1.0" encoding="UTF-8"?>'
    str << '<c:chartSpace xmlns:c="' << Axlsx::XML_NS_C << '">'
    doc = Nokogiri::XML(@axis.to_xml_string(str))

    assert_equal(1, doc.xpath("//c:tickLblSkip[@val='5']").size)
    assert_equal(1, doc.xpath("//c:tickMarkSkip[@val='3']").size)
  end
end
