# frozen_string_literal: true

require 'tc_helper'

class TestMimeTypeUtils < Minitest::Test
  def setup
    stub_request(:head, 'https://example.com/sample-image.png')
      .to_return(status: 501)

    stub_request(:get, 'https://example.com/sample-image.png')
      .to_return(body: File.new('examples/sample.png'), status: 200, headers: { 'Content-Type' => 'image/png' })

    @test_img = "#{File.dirname(__FILE__)}/../fixtures/image1.jpeg"
    @test_img_url = "https://example.com/sample-image.png"
  end

  def teardown; end

  def test_mime_type_utils
    assert_equal('image/jpeg', Axlsx::MimeTypeUtils.get_mime_type(@test_img))
    assert_equal('image/png', Axlsx::MimeTypeUtils.get_mime_type_from_uri(@test_img_url))
  end

  def test_escape_uri
    assert_raises(URI::InvalidURIError) { Axlsx::MimeTypeUtils.get_mime_type_from_uri('| ls') }
  end
end
