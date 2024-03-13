# frozen_string_literal: true

module Axlsx
  # An override content part. These parts are automatically created by for you based on the content of your package.
  class Override < AbstractContentType
    include Axlsx::Accessors

    # Serialization node name for this object
    NODE_NAME = 'Override'

    # The name and location of the part.
    # @return [String]
    # @!attribute
    string_attr_accessor :part_name
    alias :PartName :part_name
    alias :PartName= :part_name=

    # Serializes this object to xml
    def to_xml_string(str = +'')
      super(NODE_NAME, str)
    end
  end
end
