# frozen_string_literal: true

require 'tc_helper'

class TestUriUtils < Minitest::Test
  def setup
    @test_url = 'https://example.com/test-resource'
    @redirect_url = 'https://example.com/redirect-me'
    @final_url = 'https://example.com/final-destination'

    # Stub successful response (using HEAD since uri_utils uses HEAD requests)
    stub_request(:head, @test_url)
      .to_return(status: 200, headers: { 'Content-Type' => 'image/png' })

    # Stub redirect chain
    stub_request(:head, @redirect_url)
      .to_return(status: 302, headers: { 'Location' => @final_url })

    stub_request(:head, @final_url)
      .to_return(status: 200, headers: { 'Content-Type' => 'image/jpeg' })

    # Stub too many redirects
    stub_request(:head, 'https://example.com/infinite-redirect')
      .to_return(status: 301, headers: { 'Location' => 'https://example.com/infinite-redirect' })

    # Stub error response
    stub_request(:head, 'https://example.com/not-found')
      .to_return(status: 404, body: 'Not Found')
  end

  def test_fetch_headers_success
    uri = URI.parse(@test_url)
    response = Axlsx::UriUtils.fetch_headers(uri)

    assert_kind_of Net::HTTPSuccess, response
    assert_equal '200', response.code
    assert_equal 'image/png', response['Content-Type']
  end

  def test_fetch_headers_with_redirect
    uri = URI.parse(@redirect_url)
    response = Axlsx::UriUtils.fetch_headers(uri)

    assert_kind_of Net::HTTPSuccess, response
    assert_equal '200', response.code
    assert_equal 'image/jpeg', response['Content-Type']
  end

  def test_fetch_headers_too_many_redirects
    uri = URI.parse('https://example.com/infinite-redirect')

    error = assert_raises(ArgumentError) do
      Axlsx::UriUtils.fetch_headers(uri)
    end

    assert_includes error.message, 'Too many redirects'
    assert_includes error.message, 'exceeded 5'
  end

  def test_fetch_headers_custom_redirect_limit
    uri = URI.parse(@redirect_url)
    response = Axlsx::UriUtils.fetch_headers(uri, 2)

    # Should work with limit of 2 (since we have 1 redirect, need limit > redirect count)
    assert_kind_of Net::HTTPSuccess, response
    assert_equal '200', response.code
  end

  def test_fetch_headers_custom_redirect_limit_exceeded
    # Create a chain with 3 redirects
    stub_request(:head, 'https://example.com/chain1')
      .to_return(status: 302, headers: { 'Location' => 'https://example.com/chain2' })

    stub_request(:head, 'https://example.com/chain2')
      .to_return(status: 302, headers: { 'Location' => 'https://example.com/chain3' })

    stub_request(:head, 'https://example.com/chain3')
      .to_return(status: 200, headers: { 'Content-Type' => 'text/html' })

    uri = URI.parse('https://example.com/chain1')

    # Should fail with limit of 1 (needs 2 redirects)
    error = assert_raises(ArgumentError) do
      Axlsx::UriUtils.fetch_headers(uri, 1)
    end

    assert_includes error.message, 'Too many redirects'
  end

  def test_fetch_headers_http_error
    uri = URI.parse('https://example.com/not-found')

    error = assert_raises(ArgumentError) do
      Axlsx::UriUtils.fetch_headers(uri)
    end

    assert_includes error.message, 'Failed to fetch resource'
    assert_includes error.message, '404'
  end

  def test_fetch_headers_with_query_parameters
    test_url_with_query = 'https://example.com/test?param1=value1&param2=value2'

    stub_request(:head, test_url_with_query)
      .to_return(status: 200, headers: { 'Content-Type' => 'application/json' })

    uri = URI.parse(test_url_with_query)
    response = Axlsx::UriUtils.fetch_headers(uri)

    assert_kind_of Net::HTTPSuccess, response
    assert_equal 'application/json', response['Content-Type']
  end

  def test_fetch_headers_relative_redirect
    # Test redirect with relative URL
    stub_request(:head, 'https://example.com/relative-redirect')
      .to_return(status: 302, headers: { 'Location' => '/relative-target' })

    stub_request(:head, 'https://example.com/relative-target')
      .to_return(status: 200, headers: { 'Content-Type' => 'text/plain' })

    uri = URI.parse('https://example.com/relative-redirect')
    response = Axlsx::UriUtils.fetch_headers(uri)

    assert_kind_of Net::HTTPSuccess, response
    assert_equal 'text/plain', response['Content-Type']
  end

  def test_fetch_headers_https_scheme
    https_url = 'https://secure.example.com/secure-resource'

    stub_request(:head, https_url)
      .to_return(status: 200, headers: { 'Content-Type' => 'application/pdf' })

    uri = URI.parse(https_url)
    response = Axlsx::UriUtils.fetch_headers(uri)

    assert_kind_of Net::HTTPSuccess, response
    assert_equal 'application/pdf', response['Content-Type']
  end

  def test_fetch_headers_with_default_redirect_limit
    # Create 4 redirects (should succeed with default limit of 5)
    (1..4).each do |i|
      next_url = i == 4 ? 'https://example.com/final' : "https://example.com/redirect#{i + 1}"
      stub_request(:head, "https://example.com/redirect#{i}")
        .to_return(status: 302, headers: { 'Location' => next_url })
    end

    stub_request(:head, 'https://example.com/final')
      .to_return(status: 200, headers: { 'Content-Type' => 'text/html' })

    uri = URI.parse('https://example.com/redirect1')
    response = Axlsx::UriUtils.fetch_headers(uri)

    # Should succeed with 4 redirects (within default limit of 5)
    assert_kind_of Net::HTTPSuccess, response
    assert_equal 'text/html', response['Content-Type']
  end

  def test_fetch_headers_head_to_get_fallback
    # Test the HEAD -> GET fallback when server doesn't support HEAD
    stub_request(:head, 'https://example.com/no-head-support')
      .to_return(status: 405, body: 'Method Not Allowed')

    stub_request(:get, 'https://example.com/no-head-support')
      .to_return(status: 200, headers: { 'Content-Type' => 'image/gif' })

    uri = URI.parse('https://example.com/no-head-support')
    response = Axlsx::UriUtils.fetch_headers(uri)

    assert_kind_of Net::HTTPSuccess, response
    assert_equal '200', response.code
    assert_equal 'image/gif', response['Content-Type']
  end

  def test_fetch_headers_head_not_implemented_fallback
    # Test fallback when server returns 501 Not Implemented
    stub_request(:head, 'https://example.com/head-not-implemented')
      .to_return(status: 501, body: 'Not Implemented')

    stub_request(:get, 'https://example.com/head-not-implemented')
      .to_return(status: 200, headers: { 'Content-Type' => 'application/xml' })

    uri = URI.parse('https://example.com/head-not-implemented')
    response = Axlsx::UriUtils.fetch_headers(uri)

    assert_kind_of Net::HTTPSuccess, response
    assert_equal 'application/xml', response['Content-Type']
  end

  def test_fetch_headers_both_head_and_get_fail
    # Test case where both HEAD and GET fail with 405
    stub_request(:head, 'https://example.com/both-fail')
      .to_return(status: 405, body: 'Method Not Allowed')

    stub_request(:get, 'https://example.com/both-fail')
      .to_return(status: 405, body: 'Method Not Allowed')

    uri = URI.parse('https://example.com/both-fail')

    error = assert_raises(ArgumentError) do
      Axlsx::UriUtils.fetch_headers(uri)
    end

    assert_includes error.message, 'Failed to fetch resource'
    assert_includes error.message, '405'
  end

  def test_fetch_headers_http_scheme
    # Test regular HTTP (non-HTTPS) URLs
    http_url = 'http://example.com/http-resource'

    stub_request(:head, http_url)
      .to_return(status: 200, headers: { 'Content-Type' => 'text/css' })

    uri = URI.parse(http_url)
    response = Axlsx::UriUtils.fetch_headers(uri)

    assert_kind_of Net::HTTPSuccess, response
    assert_equal 'text/css', response['Content-Type']
  end

  def test_fetch_headers_different_redirect_codes
    # Test different redirect status codes (301, 307, 308)
    stub_request(:head, 'https://example.com/redirect-301')
      .to_return(status: 301, headers: { 'Location' => 'https://example.com/moved-permanently' })

    stub_request(:head, 'https://example.com/moved-permanently')
      .to_return(status: 200, headers: { 'Content-Type' => 'text/javascript' })

    uri = URI.parse('https://example.com/redirect-301')
    response = Axlsx::UriUtils.fetch_headers(uri)

    assert_kind_of Net::HTTPSuccess, response
    assert_equal 'text/javascript', response['Content-Type']
  end

  def test_fetch_headers_zero_redirect_limit
    # Test with redirect_limit = 0 (no redirects allowed)
    uri = URI.parse(@redirect_url)

    error = assert_raises(ArgumentError) do
      Axlsx::UriUtils.fetch_headers(uri, 0)
    end

    assert_includes error.message, 'Too many redirects'
    assert_includes error.message, 'exceeded 0'
  end

  def test_fetch_headers_server_error
    # Test server errors (5xx)
    stub_request(:head, 'https://example.com/server-error')
      .to_return(status: 500, body: 'Internal Server Error')

    uri = URI.parse('https://example.com/server-error')

    error = assert_raises(ArgumentError) do
      Axlsx::UriUtils.fetch_headers(uri)
    end

    assert_includes error.message, 'Failed to fetch resource'
    assert_includes error.message, '500'
  end

  def test_fetch_headers_redirect_missing_location
    # Test redirect without Location header
    stub_request(:head, 'https://example.com/bad-redirect')
      .to_return(status: 300, headers: {})

    uri = URI.parse('https://example.com/bad-redirect')

    error = assert_raises(ArgumentError) do
      Axlsx::UriUtils.fetch_headers(uri)
    end

    assert_includes error.message, 'Redirect response missing Location header'
  end
end
