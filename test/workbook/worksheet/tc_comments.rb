# frozen_string_literal: true

require 'tc_helper'

class TestComments < Minitest::Test
  def setup
    p = Axlsx::Package.new
    wb = p.workbook
    @ws = wb.add_worksheet
    @c1 = @ws.add_comment ref: 'A1', text: 'penut machine', author: 'crank'
    @c2 = @ws.add_comment ref: 'C3', text: 'rust bucket', author: 'PO'
  end

  def test_initialize
    assert_raises(ArgumentError) { Axlsx::Comments.new }
    assert_kind_of(Axlsx::VmlDrawing, @ws.comments.vml_drawing)
  end

  def test_add_comment
    assert_equal(2, @ws.comments.size)
    assert_raises(ArgumentError) { @ws.comments.add_comment }
    assert_raises(ArgumentError) { @ws.comments.add_comment(text: 'Yes We Can', ref: 'A1') }
    assert_raises(ArgumentError) { @ws.comments.add_comment(author: 'bob', ref: 'A1') }
    assert_raises(ArgumentError) { @ws.comments.add_comment(author: 'bob', text: 'Yes We Can') }
    refute_raises { @ws.comments.add_comment(author: 'bob', text: 'Yes We Can', ref: 'A1') }
    assert_equal(3, @ws.comments.size)
  end

  def test_authors
    assert_equal(@ws.comments.authors.size, @ws.comments.size)
    @ws.add_comment(text: 'Yes We Can!', author: 'bob', ref: 'F1')

    assert_equal(3, @ws.comments.authors.size)
    @ws.add_comment(text: 'Yes We Can!', author: 'bob', ref: 'F1')

    assert_equal(3, @ws.comments.authors.size, 'only unique authors are returned')
  end

  def test_pn
    assert_equal(@ws.comments.pn, format(Axlsx::COMMENT_PN, @ws.index + 1))
  end

  def test_index
    assert_equal(@ws.index, @ws.comments.index)
  end

  def test_to_xml_string
    doc = Nokogiri::XML(@ws.comments.to_xml_string)
    schema = Nokogiri::XML::Schema(File.open(Axlsx::SML_XSD))

    assert_equal(0, schema.validate(doc).length)

    # TODO: figure out why these xpath expressions dont work!
    # assert(doc.xpath("//comments"))
    # assert_equal(doc.xpath("//xmlns:author").size, @ws.comments.authors.size)
    # assert_equal(doc.xpath("//comment").size, @ws.comments.size)
  end
end
