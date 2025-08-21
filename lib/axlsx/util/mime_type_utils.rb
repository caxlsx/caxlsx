# frozen_string_literal: true

module Axlsx
  # This module defines some utils related with mime type detection
  module MimeTypeUtils
    # Extension to MIME type mapping for Windows fallback
    EXTENSION_FALLBACKS = {
      '.jpg' => 'image/jpeg',
      '.jpeg' => 'image/jpeg',
      '.png' => 'image/png',
      '.gif' => 'image/gif'
    }.freeze

    class << self
      # Detect a file mime type
      # @param [String] v File path
      # @return [String] File mime type
      def get_mime_type(v)
        mime_type = Marcel::MimeType.for(Pathname.new(v))

        # Windows fallback: Marcel sometimes returns application/octet-stream for valid image files
        if mime_type == 'application/octet-stream' && windows_platform?
          extension = File.extname(v).downcase
          # Verify it's actually an image by checking the file header
          if EXTENSION_FALLBACKS.key?(extension) && File.exist?(v) && valid_image_file?(v, extension)
            mime_type = EXTENSION_FALLBACKS[extension]
          end
        end

        mime_type
      end

      # Detect a file mime type from URI
      # @param [String] v URI
      # @return [String] File mime type
      def get_mime_type_from_uri(v)
        uri = URI.parse(v)

        unless uri.is_a?(URI::HTTP) || uri.is_a?(URI::HTTPS)
          raise URI::InvalidURIError, "Only HTTP/HTTPS URIs are supported"
        end

        response = UriUtils.fetch_headers(uri)
        mime_type = response&.[]('content-type')&.split(';')&.first&.strip

        if mime_type.nil? || mime_type.empty?
          raise ArgumentError, "Unable to determine MIME type from response"
        end

        mime_type
      end

      private

      def windows_platform?
        RUBY_PLATFORM =~ /mswin|mingw|cygwin/
      end

      def valid_image_file?(file_path, extension)
        return false unless File.exist?(file_path)

        # Check magic bytes for common image formats
        begin
          header = File.binread(file_path, 10)
          case extension
          when '.jpg', '.jpeg'
            header[0..2].bytes == [0xFF, 0xD8, 0xFF] # JPEG magic bytes
          when '.png'
            header[0..7].bytes == [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A] # PNG magic bytes
          when '.gif'
            header[0..2] == 'GIF' # GIF magic bytes
          else
            false
          end
        rescue StandardError
          false
        end
      end
    end
  end
end
