# frozen_string_literal: true

module Axlsx
  # A collection of Cfvo objects that initializes with the required
  # first two items
  class Cfvos < SimpleTypedList
    def initialize
      super(Cfvo)
    end

    # Serialize the Cfvo object
    # @param [#<<] str A String, buffer or IO to append the serialization to.
    # @return [void]
    def to_xml_string(str = +'')
      each { |cfvo| cfvo.to_xml_string(str) }
    end
  end
end
