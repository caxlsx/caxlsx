# frozen_string_literal: true

require 'tc_helper'

class TestCfvo < Minitest::Test
  def setup
    @cfvo = Axlsx::Cfvo.new(val: "0", type: :min)
  end

  def test_val
    refute_raises { @cfvo.val = "abc" }
    assert_equal("abc", @cfvo.val)
  end

  def test_type
    assert_raises(ArgumentError) { @cfvo.type = :invalid_type }
    refute_raises { @cfvo.type = :max }
    assert_equal(:max, @cfvo.type)
  end

  def test_gte
    assert_raises(ArgumentError) { @cfvo.gte = :bob }
    assert(@cfvo.gte)
    refute_raises { @cfvo.gte = false }
    assert_false(@cfvo.gte)
  end

  def test_to_xml_string
    doc = Nokogiri::XML.parse(@cfvo.to_xml_string)

    assert doc.xpath(".//cfvo[@type='min'][@val=0][@gte=true]")
  end
end
