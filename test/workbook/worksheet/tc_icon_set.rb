# frozen_string_literal: true

require 'tc_helper'

class TestIconSet < Minitest::Test
  def setup
    @icon_set = Axlsx::IconSet.new
  end

  def test_defaults
    assert_equal("3TrafficLights1", @icon_set.iconSet)
    assert(@icon_set.percent)
    assert_false(@icon_set.reverse)
    assert(@icon_set.showValue)
    assert_equal([0, 33, 67], @icon_set.interpolationPoints)
  end

  def test_icon_set
    assert_raises(ArgumentError) { @icon_set.iconSet = "invalid_value" }
    refute_raises { @icon_set.iconSet = "5Rating" }
    assert_equal("5Rating", @icon_set.iconSet)
  end

  def test_interpolation_points
    assert_raises(ArgumentError) { @icon_set.interpolationPoints = ["invalid_value"] }
    refute_raises { @icon_set.interpolationPoints = [0, 60, 80] }
    assert_equal([0, 60, 80], @icon_set.interpolationPoints)
  end

  def test_percent
    assert_raises(ArgumentError) { @icon_set.percent = :invalid_type }
    refute_raises { @icon_set.percent = false }
    assert_false(@icon_set.percent)
  end

  def test_showValue
    assert_raises(ArgumentError) { @icon_set.showValue = :invalid_type }
    refute_raises { @icon_set.showValue = false }
    assert_false(@icon_set.showValue)
  end

  def test_reverse
    assert_raises(ArgumentError) { @icon_set.reverse = :invalid_type }
    refute_raises { @icon_set.reverse = false }
    assert_false(@icon_set.reverse)
  end

  def test_to_xml_string
    doc = Nokogiri::XML.parse(@icon_set.to_xml_string)

    assert_equal(1, doc.xpath(".//iconSet[@iconSet='3TrafficLights1'][@percent=1][@reverse=0][@showValue=1]").size)
    assert_equal(3, doc.xpath(".//iconSet//cfvo").size)
  end

  def test_to_xml_string_with_interpolation_points
    @icon_set.iconSet = '5Arrows'
    @icon_set.interpolationPoints = [0, 20, 40, 60, 80]
    doc = Nokogiri::XML.parse(@icon_set.to_xml_string)

    assert_equal(1, doc.xpath(".//iconSet[@iconSet='5Arrows'][@percent=1][@reverse=0][@showValue=1]").size)
    assert_equal(5, doc.xpath(".//iconSet//cfvo").size)
  end
end
