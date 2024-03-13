# frozen_string_literal: true

module Axlsx
  # An default content part. These parts are automatically created by for you based on the content of your package.
  class Default < AbstractContentType
    include Axlsx::Accessors

    # The serialization node name for this class
    NODE_NAME = 'Default'

    # The extension of the content type.
    # @return [String]
    # @!attribute
    string_attr_accessor :extension
    alias :Extension :extension
    alias :Extension= :extension=

    # Serializes this object to xml
    def to_xml_string(str = +'')
      super(NODE_NAME, str)
    end
  end
end
