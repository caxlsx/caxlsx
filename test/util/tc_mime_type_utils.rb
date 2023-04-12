require 'tc_helper.rb'
class TestMimeTypeUtils < Test::Unit::TestCase
  def setup
    @test_img = File.dirname(__FILE__) + "/../fixtures/image1.jpeg"
    @test_img_url = "https://via.placeholder.com/150.png"
  end

  def teardown
  end

  def test_mime_type_utils
    assert_equal(Axlsx::MimeTypeUtils::get_mime_type(@test_img), 'image/jpeg')
    assert_equal(Axlsx::MimeTypeUtils::get_mime_type_from_uri(@test_img_url), 'image/png')
  end
end
