# frozen_string_literal: true

require 'net/http'
require 'uri'

module Axlsx
  # This module defines some utils related to mime type detection
  module UriUtils
    class << self
      # Performs an HTTP GeT request fetching only headers
      def fetch_headers(uri, redirect_limit = 5)
        redirect_count = 0
        use_get = false

        while redirect_count < redirect_limit
          case (response = fetch_headers_request(uri, use_get: use_get))
          when Net::HTTPSuccess
            return response
          when Net::HTTPMethodNotAllowed, Net::HTTPNotImplemented
            fail_response(response) if use_get
            use_get = true
            next # Retry current URL with GET
          when Net::HTTPRedirection
            uri = follow_redirect(uri, response)
            redirect_count += 1
          else
            fail_response(response)
          end
        end

        raise ArgumentError, "Too many redirects (exceeded #{redirect_limit})"
      end

      private

      def fetch_headers_request(uri, use_get: false)
        Net::HTTP.start(uri.host, uri.port, use_ssl: uri.scheme == 'https') do |http|
          path = build_path(uri)
          if use_get
            http.request_get(path) { |response| break(response) } # Exit early with just headers
          else
            http.head(path)
          end
        end
      end

      def build_path(uri)
        "#{uri.path.empty? ? '/' : uri.path}#{"?#{uri.query}" if uri.query}"
      end

      def follow_redirect(original_uri, response)
        location = response['location']

        if location.nil? || location.empty?
          raise ArgumentError, "Redirect response missing Location header"
        end

        if location.start_with?('http://', 'https://')
          URI.parse(location)
        else
          URI.join(original_uri, location)
        end
      end

      def fail_response(response)
        raise ArgumentError, "Failed to fetch resource: #{response.code} #{response.message}"
      end
    end
  end
end
