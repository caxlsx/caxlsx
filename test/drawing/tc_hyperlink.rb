require 'tc_helper.rb'

class TestHyperlink < Test::Unit::TestCase
  def setup
    @p = Axlsx::Package.new
    ws = @p.workbook.add_worksheet
    @test_img = File.dirname(__FILE__) + "/../fixtures/image1.jpeg"
    @image = ws.add_image :image_src => @test_img, :hyperlink => "http://axlsx.blogspot.com"
    @hyperlink = @image.hyperlink
  end

  def teardown; end

  def test_href
    assert_nothing_raised { @hyperlink.href = "http://axlsx.blogspot.com" }
    assert_equal("http://axlsx.blogspot.com", @hyperlink.href)
  end

  def test_tgtFrame
    assert_nothing_raised { @hyperlink.tgtFrame = "http://axlsx.blogspot.com" }
    assert_equal("http://axlsx.blogspot.com", @hyperlink.tgtFrame)
  end

  def test_tooltip
    assert_nothing_raised { @hyperlink.tooltip = "http://axlsx.blogspot.com" }
    assert_equal("http://axlsx.blogspot.com", @hyperlink.tooltip)
  end

  def test_invalidUrl
    assert_nothing_raised { @hyperlink.invalidUrl = "http://axlsx.blogspot.com" }
    assert_equal("http://axlsx.blogspot.com", @hyperlink.invalidUrl)
  end

  def test_action
    assert_nothing_raised { @hyperlink.action = "flee" }
    assert_equal("flee", @hyperlink.action)
  end

  def test_endSnd
    assert_nothing_raised { @hyperlink.endSnd = "true" }
    assert_raise(ArgumentError) { @hyperlink.endSnd = "bob" }
    assert_equal("true", @hyperlink.endSnd)
  end

  def test_highlightClick
    assert_nothing_raised { @hyperlink.highlightClick = false }
    assert_raise(ArgumentError) { @hyperlink.highlightClick = "bob" }
    refute(@hyperlink.highlightClick)
  end

  def test_history
    assert_nothing_raised { @hyperlink.history = false }
    assert_raise(ArgumentError) { @hyperlink.history = "bob" }
    refute(@hyperlink.history)
  end

  def test_to_xml_string
    doc = Nokogiri::XML(@p.workbook.worksheets.first.drawing.to_xml_string)

    assert(doc.xpath("//a:hlinkClick"))
  end
end
