# frozen_string_literal: true

require 'tc_helper.rb'
class TestMimeTypeUtils < Test::Unit::TestCase
  def setup
    stub_request(:get, 'https://example.com/sample-image.png')
      .to_return(body: File.new('examples/sample.png'), status: 200)

    @test_img = File.dirname(__FILE__) + "/../fixtures/image1.jpeg"
    @test_img_url = "https://example.com/sample-image.png"
  end

  def teardown
  end

  def test_mime_type_utils
    assert_equal(Axlsx::MimeTypeUtils::get_mime_type(@test_img), 'image/jpeg')
    assert_equal(Axlsx::MimeTypeUtils::get_mime_type_from_uri(@test_img_url), 'image/png')
  end
end
