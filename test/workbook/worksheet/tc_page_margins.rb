# frozen_string_literal: true

require 'tc_helper'

class TestPageMargins < Minitest::Test
  def setup
    p = Axlsx::Package.new
    ws = p.workbook.add_worksheet name: "hmmm"
    @pm = ws.page_margins
  end

  def test_initialize
    assert_equal(Axlsx::PageMargins::DEFAULT_LEFT_RIGHT, @pm.left)
    assert_equal(Axlsx::PageMargins::DEFAULT_LEFT_RIGHT, @pm.right)
    assert_equal(Axlsx::PageMargins::DEFAULT_TOP_BOTTOM, @pm.top)
    assert_equal(Axlsx::PageMargins::DEFAULT_TOP_BOTTOM, @pm.bottom)
    assert_equal(Axlsx::PageMargins::DEFAULT_HEADER_FOOTER, @pm.header)
    assert_equal(Axlsx::PageMargins::DEFAULT_HEADER_FOOTER, @pm.footer)
  end

  def test_initialize_with_options
    optioned = Axlsx::PageMargins.new(left: 2, right: 3, top: 2, bottom: 1, header: 0.1, footer: 0.1)

    assert_equal(2, optioned.left)
    assert_equal(3, optioned.right)
    assert_equal(2, optioned.top)
    assert_equal(1, optioned.bottom)
    assert_in_delta(0.1, optioned.header)
    assert_in_delta(0.1, optioned.footer)
  end

  def test_set_all_values
    @pm.set(left: 1.1, right: 1.2, top: 1.3, bottom: 1.4, header: 0.8, footer: 0.9)

    assert_in_delta(1.1, @pm.left)
    assert_in_delta(1.2, @pm.right)
    assert_in_delta(1.3, @pm.top)
    assert_in_delta(1.4, @pm.bottom)
    assert_in_delta(0.8, @pm.header)
    assert_in_delta(0.9, @pm.footer)
  end

  def test_set_some_values
    @pm.set(left: 1.1, right: 1.2)

    assert_in_delta(1.1, @pm.left)
    assert_in_delta(1.2, @pm.right)
    assert_equal(Axlsx::PageMargins::DEFAULT_TOP_BOTTOM, @pm.top)
    assert_equal(Axlsx::PageMargins::DEFAULT_TOP_BOTTOM, @pm.bottom)
    assert_equal(Axlsx::PageMargins::DEFAULT_HEADER_FOOTER, @pm.header)
    assert_equal(Axlsx::PageMargins::DEFAULT_HEADER_FOOTER, @pm.footer)
  end

  def test_to_xml
    @pm.left = 1.1
    @pm.right = 1.2
    @pm.top = 1.3
    @pm.bottom = 1.4
    @pm.header = 0.8
    @pm.footer = 0.9
    doc = Nokogiri::XML.parse(@pm.to_xml_string)

    assert_equal(1, doc.xpath(".//pageMargins[@left=1.1][@right=1.2][@top=1.3][@bottom=1.4][@header=0.8][@footer=0.9]").size)
  end

  def test_left
    assert_raises(ArgumentError) { @pm.left = -1.2 }
    refute_raises { @pm.left = 1.5 }
    assert_in_delta(@pm.left, 1.5)
  end

  def test_right
    assert_raises(ArgumentError) { @pm.right = -1.2 }
    refute_raises { @pm.right = 1.5 }
    assert_in_delta(@pm.right, 1.5)
  end

  def test_top
    assert_raises(ArgumentError) { @pm.top = -1.2 }
    refute_raises { @pm.top = 1.5 }
    assert_in_delta(@pm.top, 1.5)
  end

  def test_bottom
    assert_raises(ArgumentError) { @pm.bottom = -1.2 }
    refute_raises { @pm.bottom = 1.5 }
    assert_in_delta(@pm.bottom, 1.5)
  end

  def test_header
    assert_raises(ArgumentError) { @pm.header = -1.2 }
    refute_raises { @pm.header = 1.5 }
    assert_in_delta(@pm.header, 1.5)
  end

  def test_footer
    assert_raises(ArgumentError) { @pm.footer = -1.2 }
    refute_raises { @pm.footer = 1.5 }
    assert_in_delta(@pm.footer, 1.5)
  end
end
