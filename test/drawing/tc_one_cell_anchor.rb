# frozen_string_literal: true

require 'tc_helper'

class TestOneCellAnchor < Minitest::Test
  def setup
    @p = Axlsx::Package.new
    @ws = @p.workbook.add_worksheet
    @test_img = "#{File.dirname(__FILE__)}/../fixtures/image1.jpeg"
    @image = @ws.add_image image_src: @test_img
    @anchor = @image.anchor
  end

  def teardown; end

  def test_initialization
    assert_equal(0, @anchor.from.col)
    assert_equal(0, @anchor.from.row)
    assert_equal(0, @anchor.width)
    assert_equal(0, @anchor.height)
  end

  def test_from
    assert_kind_of(Axlsx::Marker, @anchor.from)
  end

  def test_object
    assert_kind_of(Axlsx::Pic, @anchor.object)
  end

  def test_index
    assert_equal(@anchor.index, @anchor.drawing.anchors.index(@anchor))
  end

  def test_width
    assert_raises(ArgumentError) { @anchor.width = "a" }
    refute_raises { @anchor.width = 600 }
    assert_equal(600, @anchor.width)
  end

  def test_height
    assert_raises(ArgumentError) { @anchor.height = "a" }
    refute_raises { @anchor.height = 400 }
    assert_equal(400, @anchor.height)
  end

  def test_ext
    ext = @anchor.send(:ext)

    assert_equal(ext[:cx], (@anchor.width * 914_400 / 96))
    assert_equal(ext[:cy], (@anchor.height * 914_400 / 96))
  end

  def test_options
    assert_raises(ArgumentError, 'invalid start_at') { @ws.add_image image_src: @test_img, start_at: [1] }
    i = @ws.add_image image_src: @test_img, start_at: [1, 2], width: 100, height: 200, name: "someimage", descr: "a neat image"

    assert_equal("a neat image", i.descr)
    assert_equal("someimage", i.name)
    assert_equal(200, i.height)
    assert_equal(100, i.width)
    assert_equal(1, i.anchor.from.col)
    assert_equal(2, i.anchor.from.row)
    assert_equal(@test_img, i.image_src)
  end
end
