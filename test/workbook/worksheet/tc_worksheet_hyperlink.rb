# frozen_string_literal: true

require 'tc_helper'

class TestWorksheetHyperlink < Minitest::Test
  def setup
    p = Axlsx::Package.new
    wb = p.workbook
    @ws = wb.add_worksheet
    @options = { location: 'https://github.com/randym/axlsx?foo=1&bar=2', tooltip: 'axlsx', ref: 'A1', display: 'AXSLX', target: :internal }
    @a = @ws.add_hyperlink @options
  end

  def test_initailize
    assert_raises(ArgumentError) { Axlsx::WorksheetHyperlink.new }
  end

  def test_location
    assert_equal(@options[:location], @a.location)
  end

  def test_tooltip
    assert_equal(@options[:tooltip], @a.tooltip)
  end

  def test_target
    assert_equal(@options[:target], Axlsx.instance_values_for(@a)['target'])
  end

  def test_display
    assert_equal(@options[:display], @a.display)
  end

  def test_ref
    assert_equal(@options[:ref], @a.ref)
  end

  def test_to_xml_string_with_non_external
    doc = Nokogiri::XML(@ws.to_xml_string)

    assert_equal(1, doc.xpath("//xmlns:hyperlink[@ref='#{@a.ref}']").size)
    assert_equal(1, doc.xpath("//xmlns:hyperlink[@tooltip='#{@a.tooltip}']").size)
    assert_equal(1, doc.xpath("//xmlns:hyperlink[@location='#{@a.location}']").size)
    assert_equal(1, doc.xpath("//xmlns:hyperlink[@display='#{@a.display}']").size)
    assert_equal(0, doc.xpath("//xmlns:hyperlink[@r:id]").size)
  end

  def test_to_xml_stirng_with_external
    @a.target = :external
    doc = Nokogiri::XML(@ws.to_xml_string)

    assert_equal(1, doc.xpath("//xmlns:hyperlink[@ref='#{@a.ref}']").size)
    assert_equal(1, doc.xpath("//xmlns:hyperlink[@tooltip='#{@a.tooltip}']").size)
    assert_equal(1, doc.xpath("//xmlns:hyperlink[@display='#{@a.display}']").size)
    assert_equal(0, doc.xpath("//xmlns:hyperlink[@location='#{@a.location}']").size)
    assert_equal(1, doc.xpath("//xmlns:hyperlink[@r:id='#{@a.relationship.Id}']").size)
  end

  def test_to_xml_string_with_underscores
    @a.location = "'dummy_sheet_2'!A1"
    doc = Nokogiri::XML(@ws.to_xml_string)

    assert_equal(1, doc.xpath("//xmlns:hyperlink/@location").size)
    assert_equal(@a.location, doc.xpath("//xmlns:hyperlink/@location").first.value)
  end
end
