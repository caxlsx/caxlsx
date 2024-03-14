# frozen_string_literal: true

require 'tc_helper'

class TestPatternFill < Minitest::Test
  def setup
    @item = Axlsx::PatternFill.new
  end

  def teardown; end

  def test_initialiation
    assert_equal(:none, @item.patternType)
    assert_nil(@item.bgColor)
    assert_nil(@item.fgColor)
  end

  def test_bgColor
    assert_raises(ArgumentError) { @item.bgColor = -1.1 }
    refute_raises { @item.bgColor = Axlsx::Color.new }
    assert_equal("FF000000", @item.bgColor.rgb)
  end

  def test_fgColor
    assert_raises(ArgumentError) { @item.fgColor = -1.1 }
    refute_raises { @item.fgColor = Axlsx::Color.new }
    assert_equal("FF000000", @item.fgColor.rgb)
  end

  def test_pattern_type
    assert_raises(ArgumentError) { @item.patternType = -1.1 }
    refute_raises { @item.patternType = :lightUp }
    assert_equal(:lightUp, @item.patternType)
  end

  def test_to_xml_string
    @item = Axlsx::PatternFill.new bgColor: Axlsx::Color.new(rgb: "FF0000"), fgColor: Axlsx::Color.new(rgb: "00FF00")
    doc = Nokogiri::XML(@item.to_xml_string)

    assert(doc.xpath('//color[@rgb="FFFF0000"]'))
    assert(doc.xpath('//color[@rgb="FF00FF00"]'))
  end
end
