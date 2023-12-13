# frozen_string_literal: true

module Axlsx
  # This class extracts the common parts from Default and Override
  class AbstractContentType
    include Axlsx::Accessors
    include Axlsx::OptionsParser

    # Initializes an abstract content type
    # @see Default, Override
    def initialize(options = {})
      parse_options options
    end

    # The type of content.
    # @return [String]
    # @!attribute
    validated_attr_accessor :content_type, :validate_content_type
    alias :ContentType :content_type
    alias :ContentType= :content_type=

    # Serialize the contenty type to xml
    def to_xml_string(node_name = '', str = +'')
      str << '<' << node_name << ' '
      Axlsx.instance_values_for(self).each_with_index do |key_value, index|
        str << ' ' unless index.zero?
        str << Axlsx.camel(key_value.first) << '="' << key_value.last.to_s << '"'
      end
      str << '/>'
    end
  end
end
