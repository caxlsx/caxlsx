# frozen_string_literal: true

require 'tc_helper'

class TestCore < Test::Unit::TestCase
  def setup
    @core = Axlsx::Core.new
    # could still see some false positives if the second changes between the next two calls
    @time = Time.now.strftime('%Y-%m-%dT%H:%M:%SZ')
    @doc = Nokogiri::XML(@core.to_xml_string)
  end

  def test_valid_document
    schema = Nokogiri::XML::Schema(File.open(Axlsx::CORE_XSD))
    errors = []
    schema.validate(@doc).each do |error|
      puts error.message
      errors << error
    end

    assert_equal(0, errors.size, "core.xml Invalid" + errors.map(&:message).to_s)
  end

  def test_populates_created
    assert_equal(@doc.xpath('//dcterms:created').text, @time, "dcterms:created incorrect")
  end

  def test_created_as_option
    time = Time.utc(2013, 1, 1, 12, 0)
    c = Axlsx::Core.new :created => time
    doc = Nokogiri::XML(c.to_xml_string)

    assert_equal(doc.xpath('//dcterms:created').text, time.xmlschema, "dcterms:created incorrect")
  end

  def test_populates_default_name
    assert_equal("axlsx", @doc.xpath('//dc:creator').text, "Default name not populated")
  end

  def test_creator_as_option
    c = Axlsx::Core.new :creator => "some guy"
    doc = Nokogiri::XML(c.to_xml_string)

    assert_equal("some guy", doc.xpath('//dc:creator').text)
  end
end
