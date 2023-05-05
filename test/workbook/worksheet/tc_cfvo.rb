# frozen_string_literal: true

require 'tc_helper'

class TestCfvo < Test::Unit::TestCase
  def setup
    @cfvo = Axlsx::Cfvo.new(:val => "0", :type => :min)
  end

  def test_val
    assert_nothing_raised { @cfvo.val = "abc" }
    assert_equal("abc", @cfvo.val)
  end

  def test_type
    assert_raise(ArgumentError) { @cfvo.type = :invalid_type }
    assert_nothing_raised { @cfvo.type = :max }
    assert_equal(:max, @cfvo.type)
  end

  def test_gte
    assert_raise(ArgumentError) { @cfvo.gte = :bob }
    assert(@cfvo.gte)
    assert_nothing_raised { @cfvo.gte = false }
    refute(@cfvo.gte)
  end

  def test_to_xml_string
    doc = Nokogiri::XML.parse(@cfvo.to_xml_string)

    assert doc.xpath(".//cfvo[@type='min'][@val=0][@gte=true]")
  end
end
