# frozen_string_literal: true

require 'tc_helper'

class TestHyperlink < Minitest::Test
  def setup
    @p = Axlsx::Package.new
    ws = @p.workbook.add_worksheet
    @test_img = "#{File.dirname(__FILE__)}/../fixtures/image1.jpeg"
    @image = ws.add_image image_src: @test_img, hyperlink: "http://axlsx.blogspot.com"
    @hyperlink = @image.hyperlink
  end

  def teardown; end

  def test_href
    refute_raises { @hyperlink.href = "http://axlsx.blogspot.com" }
    assert_equal("http://axlsx.blogspot.com", @hyperlink.href)
  end

  def test_tgtFrame
    refute_raises { @hyperlink.tgtFrame = "http://axlsx.blogspot.com" }
    assert_equal("http://axlsx.blogspot.com", @hyperlink.tgtFrame)
  end

  def test_tooltip
    refute_raises { @hyperlink.tooltip = "http://axlsx.blogspot.com" }
    assert_equal("http://axlsx.blogspot.com", @hyperlink.tooltip)
  end

  def test_invalidUrl
    refute_raises { @hyperlink.invalidUrl = "http://axlsx.blogspot.com" }
    assert_equal("http://axlsx.blogspot.com", @hyperlink.invalidUrl)
  end

  def test_action
    refute_raises { @hyperlink.action = "flee" }
    assert_equal("flee", @hyperlink.action)
  end

  def test_endSnd
    refute_raises { @hyperlink.endSnd = "true" }
    assert_raises(ArgumentError) { @hyperlink.endSnd = "bob" }
    assert_equal("true", @hyperlink.endSnd)
  end

  def test_highlightClick
    refute_raises { @hyperlink.highlightClick = false }
    assert_raises(ArgumentError) { @hyperlink.highlightClick = "bob" }
    assert_false(@hyperlink.highlightClick)
  end

  def test_history
    refute_raises { @hyperlink.history = false }
    assert_raises(ArgumentError) { @hyperlink.history = "bob" }
    assert_false(@hyperlink.history)
  end

  def test_to_xml_string
    doc = Nokogiri::XML(@p.workbook.worksheets.first.drawing.to_xml_string)

    assert(doc.xpath("//a:hlinkClick"))
  end
end
