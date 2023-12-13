# frozen_string_literal: true

module Axlsx
  # This class specifies data for a particular data point.
  class NumVal < StrVal
    include Axlsx::Accessors

    # A string representing the format code to apply.
    # For more information see see the SpreadsheetML numFmt element's (§18.8.30) formatCode attribute.
    # @return [String]
    # @!attribute
    string_attr_accessor :format_code

    # creates a new NumVal object
    # @option options [String] formatCode
    # @option options [Integer] v
    def initialize(options = {})
      @format_code = "General"
      super(options)
    end

    # serialize the object
    def to_xml_string(idx, str = +'')
      Axlsx.validate_unsigned_int(idx)
      unless v.to_s.empty?
        str << '<c:pt idx="' << idx.to_s << '" formatCode="' << format_code << '"><c:v>' << v.to_s << '</c:v></c:pt>'
      end
    end
  end
end
