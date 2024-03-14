# frozen_string_literal: true

require 'tc_helper'

class TestComment < Minitest::Test
  def setup
    p = Axlsx::Package.new
    wb = p.workbook
    @ws = wb.add_worksheet
    @c1 = @ws.add_comment ref: 'A1', text: 'text with special char <', author: 'author with special char <', visible: false
    @c2 = @ws.add_comment ref: 'C3', text: 'rust bucket', author: 'PO'
  end

  def test_initailize
    assert_raises(ArgumentError) { Axlsx::Comment.new }
  end

  def test_author
    assert_equal('author with special char <', @c1.author)
    assert_equal('PO', @c2.author)
  end

  def test_text
    assert_equal('text with special char <', @c1.text)
    assert_equal('rust bucket', @c2.text)
  end

  def test_author_index
    assert_equal(1, @c1.author_index)
    assert_equal(0, @c2.author_index)
  end

  def test_visible
    assert_false(@c1.visible)
    assert(@c2.visible)
  end

  def test_ref
    assert_equal('A1', @c1.ref)
    assert_equal('C3', @c2.ref)
  end

  def test_vml_shape
    pos = Axlsx.name_to_indices(@c1.ref)

    assert_kind_of(Axlsx::VmlShape, @c1.vml_shape)
    assert_equal(@c1.vml_shape.column, pos[0])
    assert_equal(@c1.vml_shape.row, pos[1])
    assert_equal(@c1.vml_shape.row, pos[1])
    assert_equal(pos[0], @c1.vml_shape.left_column)
    assert_equal(@c1.vml_shape.top_row, pos[1])
    assert_equal(pos[0] + 2, @c1.vml_shape.right_column)
    assert_equal(@c1.vml_shape.bottom_row, pos[1] + 4)
  end

  def test_to_xml_string
    doc = Nokogiri::XML(@c1.to_xml_string)

    assert_equal(1, doc.xpath("//comment[@ref='#{@c1.ref}']").size)
    assert_equal(1, doc.xpath("//comment[@authorId='#{@c1.author_index}']").size)
    assert_equal(1, doc.xpath("//t[text()='#{@c1.author}:\n']").size)
    assert_equal(1, doc.xpath("//t[text()='#{@c1.text}']").size)
  end

  def test_comment_text_contain_author_and_text
    comment = @ws.add_comment ref: 'C4', text: 'some text', author: 'Bob'
    doc = Nokogiri::XML(comment.to_xml_string)

    assert_equal("Bob:\nsome text", doc.xpath("//comment/text").text)
  end

  def test_comment_text_does_not_contain_stray_colon_if_author_blank
    comment = @ws.add_comment ref: 'C5', text: 'some text', author: ''
    doc = Nokogiri::XML(comment.to_xml_string)

    assert_equal("some text", doc.xpath("//comment/text").text)
  end
end
