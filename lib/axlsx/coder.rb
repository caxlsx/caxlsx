require 'builder'

module Axlsx
  class Coder < Builder::XmlBase
    def encode(str)
      _escape_attribute(str).tr(::Builder::XChar::REPLACEMENT_CHAR, '')
    end
  end
end
