# frozen_string_literal: true

require 'tc_helper'

class TestBreak < Minitest::Test
  def setup
    @break = Axlsx::Break.new(id: 1, min: 1, max: 10, man: true, pt: false)
  end

  def test_id
    assert_equal(1, @break.id)
    assert_raises ArgumentError do
      Axlsx::Break.new(:hoge, { id: -1 })
    end
  end

  def test_min
    assert_equal(1, @break.min)
    assert_raises ArgumentError do
      Axlsx::Break.new(:hoge, min: -1)
    end
  end

  def test_max
    assert_equal(10, @break.max)
    assert_raises ArgumentError do
      Axlsx::Break.new(:hoge, max: -1)
    end
  end

  def test_man
    assert(@break.man)
    assert_raises ArgumentError do
      Axlsx::Break.new(man: -1)
    end
  end

  def test_pt
    assert_false(@break.pt)
    assert_raises ArgumentError do
      Axlsx::Break.new(pt: -1)
    end
  end

  def test_to_xml_string
    doc = Nokogiri::XML(@break.to_xml_string)

    assert_equal(1, doc.xpath('//brk[@id="1"][@min="1"][@max="10"][@pt=0][@man=1]').size)
  end
end
