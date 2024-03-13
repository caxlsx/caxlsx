# frozen_string_literal: true

module Axlsx
  # CellStyle defines named styles that reference defined formatting records and can be used in your worksheet.
  # @note Using Styles#add_style is the recommended way to manage cell styling.
  # @see Styles#add_style
  class CellStyle
    include Axlsx::Accessors
    include Axlsx::OptionsParser
    include Axlsx::SerializedAttributes

    # Creats a new CellStyle object
    # @option options [String] name
    # @option options [Integer] xfId
    # @option options [Integer] buildinId
    # @option options [Integer] iLevel
    # @option options [Boolean] hidden
    # @option options [Boolean] customBuiltIn
    def initialize(options = {})
      parse_options options
    end

    serializable_attributes :name, :xfId, :buildinId, :iLevel, :hidden, :customBuilin

    # The name of this cell style
    # @return [String]
    # @!attribute
    string_attr_accessor :name

    # The formatting record id this named style utilizes
    # @return [Integer]
    # @see Axlsx::Xf
    # @!attribute
    unsigned_int_attr_accessor :xfId

    # The buildinId to use when this named style is applied
    # @return [Integer]
    # @see Axlsx::NumFmt
    # @!attribute
    unsigned_int_attr_accessor :builtinId

    # Determines if this formatting is for an outline style, and what level of the outline it is to be applied to.
    # @return [Integer]
    # @!attribute
    unsigned_int_attr_accessor :iLevel

    # Determines if this named style should show in the list of styles when using Excel
    # @return [Boolean]
    # @!attribute
    boolean_attr_accessor :hidden

    # Indicates that the build in style reference has been customized.
    # @return [Boolean]
    # @!attribute
    boolean_attr_accessor :customBuiltin

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = +'')
      serialized_tag('cellStyle', str)
    end
  end
end
