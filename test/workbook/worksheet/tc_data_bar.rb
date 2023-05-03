require 'tc_helper'

class TestDataBar < Test::Unit::TestCase
  def setup
    @data_bar = Axlsx::DataBar.new :color => "FF638EC6"
  end

  def test_defaults
    assert_equal(10, @data_bar.minLength)
    assert_equal(90, @data_bar.maxLength)
    assert(@data_bar.showValue)
  end

  def test_override_default_cfvos
    data_bar = Axlsx::DataBar.new({ :color => 'FF00FF00' }, { :type => :min, :val => "20" })

    assert_equal("20", data_bar.value_objects.first.val)
    assert_equal("0", data_bar.value_objects.last.val)
  end

  def test_minLength
    assert_raise(ArgumentError) { @data_bar.minLength = :invalid_type }
    assert_nothing_raised { @data_bar.minLength = 0 }
    assert_equal(0, @data_bar.minLength)
  end

  def test_maxLength
    assert_raise(ArgumentError) { @data_bar.maxLength = :invalid_type }
    assert_nothing_raised { @data_bar.maxLength = 0 }
    assert_equal(0, @data_bar.maxLength)
  end

  def test_showValue
    assert_raise(ArgumentError) { @data_bar.showValue = :invalid_type }
    assert_nothing_raised { @data_bar.showValue = false }
    refute(@data_bar.showValue)
  end

  def test_to_xml_string
    doc = Nokogiri::XML.parse(@data_bar.to_xml_string)

    assert_equal(1, doc.xpath(".//dataBar[@minLength=10][@maxLength=90][@showValue=1]").size)
    assert_equal(2, doc.xpath(".//dataBar//cfvo").size)
    assert_equal(1, doc.xpath(".//dataBar//color").size)
  end
end
