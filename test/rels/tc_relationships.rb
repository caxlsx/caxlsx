# frozen_string_literal: true

require 'tc_helper'

class TestRelationships < Minitest::Test
  def test_for
    source_obj_1 = Object.new
    source_obj_2 = Object.new
    rel_1 = Axlsx::Relationship.new(source_obj_1, Axlsx::WORKSHEET_R, "bar")
    rel_2 = Axlsx::Relationship.new(source_obj_2, Axlsx::WORKSHEET_R, "bar")
    rels = Axlsx::Relationships.new
    rels << rel_1
    rels << rel_2

    assert_equal rel_1, rels.for(source_obj_1)
    assert_equal rel_2, rels.for(source_obj_2)
  end

  def test_valid_document
    @rels = Axlsx::Relationships.new
    schema = Nokogiri::XML::Schema(File.open(Axlsx::RELS_XSD))
    doc = Nokogiri::XML(@rels.to_xml_string)
    errors = schema.validate(doc)

    assert_empty(errors)

    @rels << Axlsx::Relationship.new(nil, Axlsx::WORKSHEET_R, "bar")
    doc = Nokogiri::XML(@rels.to_xml_string)
    errors = schema.validate(doc)

    assert_empty(errors)
  end
end
