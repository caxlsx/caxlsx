# frozen_string_literal: true

module Axlsx
  # This class extracts the common parts from Default and Override
  class AbstractContentType
    include Axlsx::OptionsParser

    # Initializes an abstract content type
    # @see Default, Override
    def initialize(options = {})
      parse_options options
    end

    # The type of content.
    # @return [String]
    attr_reader :content_type
    alias :ContentType :content_type

    # The content type.
    # @see Axlsx#validate_content_type
    def content_type=(v)
      Axlsx.validate_content_type v
      @content_type = v
    end
    alias :ContentType= :content_type=

    # Serialize the content type to xml
    def to_xml_string(node_name = '', str = +'')
      str << '<' << node_name << ' '
      Axlsx.instance_values_for(self).each_with_index do |key_value, index|
        str << ' ' unless index == 0
        str << Axlsx.camel(key_value.first) << '="' << key_value.last.to_s << '"'
      end
      str << '/>'
    end
  end
end
