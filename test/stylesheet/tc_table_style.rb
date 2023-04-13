require 'tc_helper.rb'

class TestTableStyle < Test::Unit::TestCase
  def setup
    @item = Axlsx::TableStyle.new "fisher"
  end

  def teardown; end

  def test_initialiation
    assert_equal("fisher", @item.name)
    assert_nil(@item.pivot)
    assert_nil(@item.table)
    ts = Axlsx::TableStyle.new 'price', :pivot => true, :table => true

    assert_equal('price', ts.name)
    assert(ts.pivot)
    assert(ts.table)
  end

  def test_name
    assert_raise(ArgumentError) { @item.name = -1.1 }
    assert_nothing_raised { @item.name = "lovely table style" }
    assert_equal("lovely table style", @item.name)
  end

  def test_pivot
    assert_raise(ArgumentError) { @item.pivot = -1.1 }
    assert_nothing_raised { @item.pivot = true }
    assert(@item.pivot)
  end

  def test_table
    assert_raise(ArgumentError) { @item.table = -1.1 }
    assert_nothing_raised { @item.table = true }
    assert(@item.table)
  end

  def test_to_xml_string
    doc = Nokogiri::XML(@item.to_xml_string)

    assert(doc.xpath("//tableStyle[@name='#{@item.name}']"))
  end
end
