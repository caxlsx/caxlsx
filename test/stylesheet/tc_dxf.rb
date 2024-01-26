# frozen_string_literal: true

require 'tc_helper'

class TestDxf < Test::Unit::TestCase
  def setup
    @item = Axlsx::Dxf.new
    @styles = Axlsx::Styles.new
  end

  def teardown; end

  def test_initialiation
    assert_nil(@item.alignment)
    assert_nil(@item.protection)
    assert_nil(@item.numFmt)
    assert_nil(@item.font)
    assert_nil(@item.fill)
    assert_nil(@item.border)
  end

  def test_alignment
    assert_raise(ArgumentError) { @item.alignment = -1.1 }
    assert_nothing_raised { @item.alignment = Axlsx::CellAlignment.new }
    assert_kind_of(Axlsx::CellAlignment, @item.alignment)
  end

  def test_protection
    assert_raise(ArgumentError) { @item.protection = -1.1 }
    assert_nothing_raised { @item.protection = Axlsx::CellProtection.new }
    assert_kind_of(Axlsx::CellProtection, @item.protection)
  end

  def test_numFmt
    assert_raise(ArgumentError) { @item.numFmt = 1 }
    assert_nothing_raised { @item.numFmt = Axlsx::NumFmt.new }
    assert_kind_of Axlsx::NumFmt, @item.numFmt
  end

  def test_fill
    assert_raise(ArgumentError) { @item.fill = 1 }
    assert_nothing_raised { @item.fill = Axlsx::Fill.new(Axlsx::PatternFill.new(patternType: :solid, fgColor: Axlsx::Color.new(rgb: "FF000000"))) }
    assert_kind_of Axlsx::Fill, @item.fill
  end

  def test_font
    assert_raise(ArgumentError) { @item.font = 1 }
    assert_nothing_raised { @item.font = Axlsx::Font.new }
    assert_kind_of Axlsx::Font, @item.font
  end

  def test_border
    assert_raise(ArgumentError) { @item.border = 1 }
    assert_nothing_raised { @item.border = Axlsx::Border.new }
    assert_kind_of Axlsx::Border, @item.border
  end

  def test_to_xml
    @item.border = Axlsx::Border.new
    doc = Nokogiri::XML.parse(@item.to_xml_string)

    assert_equal(1, doc.xpath(".//dxf//border").size)
    assert_equal(0, doc.xpath(".//dxf//font").size)
  end

  def test_many_options_xml
    @item.border = Axlsx::Border.new
    @item.alignment = Axlsx::CellAlignment.new
    @item.fill = Axlsx::Fill.new(Axlsx::PatternFill.new(patternType: :solid, fgColor: Axlsx::Color.new(rgb: "FF000000")))
    @item.font = Axlsx::Font.new
    @item.protection = Axlsx::CellProtection.new
    @item.numFmt = Axlsx::NumFmt.new

    doc = Nokogiri::XML.parse(@item.to_xml_string)

    assert_equal(1, doc.xpath(".//dxf//fill//patternFill[@patternType='solid']//fgColor[@rgb='FF000000']").size)
    assert_equal(1, doc.xpath(".//dxf//font").size)
    assert_equal(1, doc.xpath(".//dxf//protection").size)
    assert_equal(1, doc.xpath(".//dxf//numFmt[@numFmtId='0'][@formatCode='']").size)
    assert_equal(1, doc.xpath(".//dxf//alignment").size)
    assert_equal(1, doc.xpath(".//dxf//border").size)
  end
end
