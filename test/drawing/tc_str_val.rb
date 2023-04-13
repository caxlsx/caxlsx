require 'tc_helper.rb'

class TestStrVal < Test::Unit::TestCase
  def setup
    @str_val = Axlsx::StrVal.new :v => "1"
    @str_val_with_special_characters = Axlsx::StrVal.new :v => "a & b <c>"
  end

  def test_initialize
    assert_equal("1", @str_val.v)
  end

  def test_to_xml_string
    str = '<?xml version="1.0" encoding="UTF-8"?>'
    str << '<c:chartSpace xmlns:c="' << Axlsx::XML_NS_C << '">'
    str << @str_val.to_xml_string(0)
    doc = Nokogiri::XML(str)

    assert_equal(1, doc.xpath("//c:pt/c:v[text()='1']").size)
  end

  def test_to_xml_string_special_characters
    str = '<?xml version="1.0" encoding="UTF-8"?>'
    str << '<c:chartSpace xmlns:c="' << Axlsx::XML_NS_C << '">'
    str << @str_val_with_special_characters.to_xml_string(0)
    doc = Nokogiri::XML(str)

    assert_equal(1, doc.xpath("//c:pt/c:v[text()='a & b <c>']").size)
  end
end
