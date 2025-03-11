# frozen_string_literal: true

require 'tc_helper'

class TestTableStyle < Minitest::Test
  def setup
    @item = Axlsx::TableStyle.new "fisher"
  end

  def teardown; end

  def test_initialiation
    assert_equal("fisher", @item.name)
    assert_nil(@item.pivot)
    assert_nil(@item.table)
    ts = Axlsx::TableStyle.new 'price', pivot: true, table: true

    assert_equal('price', ts.name)
    assert(ts.pivot)
    assert(ts.table)
  end

  def test_name
    assert_raises(ArgumentError) { @item.name = -1.1 }
    refute_raises { @item.name = "lovely table style" }
    assert_equal("lovely table style", @item.name)
  end

  def test_pivot
    assert_raises(ArgumentError) { @item.pivot = -1.1 }
    refute_raises { @item.pivot = true }
    assert(@item.pivot)
  end

  def test_table
    assert_raises(ArgumentError) { @item.table = -1.1 }
    refute_raises { @item.table = true }
    assert(@item.table)
  end

  def test_to_xml_string
    doc = Nokogiri::XML(@item.to_xml_string)

    assert(doc.xpath("//tableStyle[@name='#{@item.name}']"))
  end
end
