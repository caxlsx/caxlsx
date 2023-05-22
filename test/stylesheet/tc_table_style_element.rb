# frozen_string_literal: true

require 'tc_helper'

class TestTableStyleElement < Test::Unit::TestCase
  def setup
    @item = Axlsx::TableStyleElement.new
  end

  def teardown; end

  def test_initialiation
    assert_nil(@item.type)
    assert_nil(@item.size)
    assert_nil(@item.dxfId)
    options = { type: :headerRow, size: 10, dxfId: 1 }

    tse = Axlsx::TableStyleElement.new options

    options.each { |key, value| assert_equal(tse.send(key.to_sym), value) }
  end

  def test_type
    assert_raise(ArgumentError) { @item.type = -1.1 }
    assert_nothing_raised { @item.type = :blankRow }
    assert_equal(:blankRow, @item.type)
  end

  def test_size
    assert_raise(ArgumentError) { @item.size = -1.1 }
    assert_nothing_raised { @item.size = 2 }
    assert_equal(2, @item.size)
  end

  def test_dxfId
    assert_raise(ArgumentError) { @item.dxfId = -1.1 }
    assert_nothing_raised { @item.dxfId = 7 }
    assert_equal(7, @item.dxfId)
  end

  def test_to_xml_string
    doc = Nokogiri::XML(@item.to_xml_string)
    @item.type = :headerRow

    assert(doc.xpath("//tableStyleElement[@type='#{@item.type}']"))
  end
end
