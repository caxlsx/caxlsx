# frozen_string_literal: true

require 'tc_helper'

class TestCatAxis < Minitest::Test
  def setup
    @axis = Axlsx::CatAxis.new
  end

  def teardown; end

  def test_initialization
    assert_equal(1, @axis.auto, "axis auto default incorrect")
    assert_equal(:ctr, @axis.lbl_algn, "label align default incorrect")
    assert_equal("100", @axis.lbl_offset, "label offset default incorrect")
    assert_nil(@axis.tick_lbl_skip, "tick_lbl_skip default incorrect")
    assert_nil(@axis.tick_mark_skip, "tick_mark_skip default incorrect")
  end

  def test_options
    a = Axlsx::CatAxis.new(tick_lbl_skip: 9, tick_mark_skip: 7)

    assert_equal(9, a.tick_lbl_skip)
    assert_equal(7, a.tick_mark_skip)
  end

  def test_auto
    assert_raises(ArgumentError, "requires valid auto") { @axis.auto = :nowhere }
    refute_raises { @axis.auto = false }
  end

  def test_lbl_algn
    assert_raises(ArgumentError, "requires valid label alignment") { @axis.lbl_algn = :nowhere }
    refute_raises { @axis.lbl_algn = :r }
  end

  def test_lbl_offset
    assert_raises(ArgumentError, "requires valid label offset") { @axis.lbl_offset = 'foo' }
    refute_raises { @axis.lbl_offset = "20" }
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
