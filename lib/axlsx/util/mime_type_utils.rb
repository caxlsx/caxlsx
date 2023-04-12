require 'open-uri'

module Axlsx
  # This module defines some utils related with mime type detection
  module MimeTypeUtils
    # Detect a file mime type
    # @param [String] v File path
    # @return [String] File mime type
    def self.get_mime_type(v)
      Marcel::MimeType.for(Pathname.new(v))
    end

    # Detect a file mime type from URI
    # @param [String] v URI
    # @return [String] File mime type
    def self.get_mime_type_from_uri(v)
      if URI.respond_to?(:open)
        Marcel::MimeType.for(URI.open(v))
      else
        Marcel::MimeType.for(URI.parse(v).open)
      end
    end
  end
end
