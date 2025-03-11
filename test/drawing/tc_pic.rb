# frozen_string_literal: true

require 'tc_helper'

class TestPic < Minitest::Test
  def setup
    stub_request(:get, 'https://example.com/sample-image.png')
      .to_return(body: File.new('examples/sample.png'), status: 200)

    @p = Axlsx::Package.new
    ws = @p.workbook.add_worksheet
    @test_img = @test_img_jpg = "#{File.dirname(__FILE__)}/../fixtures/image1.jpeg"
    @test_img_png =  "#{File.dirname(__FILE__)}/../fixtures/image1.png"
    @test_img_gif =  "#{File.dirname(__FILE__)}/../fixtures/image1.gif"
    @test_img_fake = "#{File.dirname(__FILE__)}/../fixtures/image1_fake.jpg"
    @test_img_remote_png = "https://example.com/sample-image.png"
    @test_img_remote_fake = "invalid_URI"
    @image = ws.add_image image_src: @test_img, hyperlink: 'https://github.com/randym', tooltip: "What's up doc?", opacity: 5
    @image_remote = ws.add_image image_src: @test_img_remote_png, remote: true, hyperlink: 'https://github.com/randym', tooltip: "What's up doc?", opacity: 5
  end

  def test_initialization
    assert_equal(@p.workbook.images.first, @image)
    assert_equal('image1.jpeg', @image.file_name)
    assert_equal(@image.image_src, @test_img)
  end

  def test_remote_img_initialization
    assert_equal(@p.workbook.images[1], @image_remote)
    assert_nil(@image_remote.file_name)
    assert_equal(@image_remote.image_src, @test_img_remote_png)
    assert_predicate(@image_remote, :remote?)
  end

  def test_anchor_swapping
    # swap from one cell to two cell when end_at is specified
    assert_kind_of(Axlsx::OneCellAnchor, @image.anchor)
    start_at = @image.anchor.from
    @image.end_at 10, 5

    assert_kind_of(Axlsx::TwoCellAnchor, @image.anchor)
    assert_equal(start_at.col, @image.anchor.from.col)
    assert_equal(start_at.row, @image.anchor.from.row)
    assert_equal(10, @image.anchor.to.col)
    assert_equal(5, @image.anchor.to.row)

    # swap from two cell to one cell when width or height are specified
    @image.width = 200

    assert_kind_of(Axlsx::OneCellAnchor, @image.anchor)
    assert_equal(start_at.col, @image.anchor.from.col)
    assert_equal(start_at.row, @image.anchor.from.row)
    assert_equal(200, @image.width)
  end

  def test_hyperlink
    assert_equal("https://github.com/randym", @image.hyperlink.href)
    @image.hyperlink = "http://axlsx.blogspot.com"

    assert_equal("http://axlsx.blogspot.com", @image.hyperlink.href)
  end

  def test_name
    assert_raises(ArgumentError) { @image.name = 49 }
    refute_raises { @image.name = "unknown" }
    assert_equal("unknown", @image.name)
  end

  def test_start_at
    assert_raises(ArgumentError) { @image.start_at "a", 1 }
    refute_raises { @image.start_at 6, 7 }
    assert_equal(6, @image.anchor.from.col)
    assert_equal(7, @image.anchor.from.row)
  end

  def test_width
    assert_raises(ArgumentError) { @image.width = "a" }
    refute_raises { @image.width = 600 }
    assert_equal(600, @image.width)
  end

  def test_height
    assert_raises(ArgumentError) { @image.height = "a" }
    refute_raises { @image.height = 600 }
    assert_equal(600, @image.height)
  end

  def test_image_src
    assert_raises(ArgumentError) { @image.image_src = __FILE__ }
    assert_raises(ArgumentError) { @image.image_src = @test_img_fake }
    refute_raises { @image.image_src = @test_img_gif }
    refute_raises { @image.image_src = @test_img_png }
    refute_raises { @image.image_src = @test_img_jpg }
    assert_equal(@image.image_src, @test_img_jpg)
  end

  def test_remote_image_src
    assert_raises(ArgumentError) { @image_remote.image_src = @test_img_fake }
    assert_raises(ArgumentError) { @image_remote.image_src = @test_img_remote_fake }
    refute_raises { @image_remote.image_src = @test_img_remote_png }
    assert_equal(@image_remote.image_src, @test_img_remote_png)
  end

  def test_descr
    assert_raises(ArgumentError) { @image.descr = 49 }
    refute_raises { @image.descr = "test" }
    assert_equal("test", @image.descr)
  end

  def test_to_xml
    schema = Nokogiri::XML::Schema(File.open(Axlsx::DRAWING_XSD))
    doc = Nokogiri::XML(@image.anchor.drawing.to_xml_string)
    errors = schema.validate(doc)

    assert_empty(errors)
  end

  def test_to_xml_has_correct_r_id
    r_id = @image.anchor.drawing.relationships.for(@image).Id
    doc = Nokogiri::XML(@image.anchor.drawing.to_xml_string)

    assert_equal r_id, doc.xpath("//a:blip").first["r:embed"]
  end
end
