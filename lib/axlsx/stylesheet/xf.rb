# frozen_string_literal: true

module Axlsx
  # The Xf class defines a formatting record for use in Styles. The recommended way to manage styles for your workbook is with Styles#add_style
  # @see Styles#add_style
  class Xf
    # does not support extList (ExtensionList)

    include Axlsx::Accessors
    include Axlsx::SerializedAttributes
    include Axlsx::OptionsParser
    # Creates a new Xf object
    # @option options [Integer] numFmtId
    # @option options [Integer] fontId
    # @option options [Integer] fillId
    # @option options [Integer] borderId
    # @option options [Integer] xfId
    # @option options [Boolean] quotePrefix
    # @option options [Boolean] pivotButton
    # @option options [Boolean] applyNumberFormat
    # @option options [Boolean] applyFont
    # @option options [Boolean] applyFill
    # @option options [Boolean] applyBorder
    # @option options [Boolean] applyAlignment
    # @option options [Boolean] applyProtection
    # @option options [CellAlignment] alignment
    # @option options [CellProtection] protection
    def initialize(options = {})
      parse_options options
    end

    serializable_attributes :numFmtId, :fontId, :fillId, :borderId, :xfId, :quotePrefix,
                            :pivotButton, :applyNumberFormat, :applyFont, :applyFill, :applyBorder, :applyAlignment,
                            :applyProtection

    # The cell alignment for this style
    # @return [CellAlignment]
    # @see CellAlignment
    attr_reader :alignment

    # The cell protection for this style
    # @return [CellProtection]
    # @see CellProtection
    attr_reader :protection

    # id of the numFmt to apply to this style
    # @return [Integer]
    # @!attribute
    unsigned_int_attr_accessor :numFmtId

    # index (0 based) of the font to be used in this style
    # @return [Integer]
    # @!attribute
    unsigned_int_attr_accessor :fontId

    # index (0 based) of the fill to be used in this style
    # @return [Integer]
    # @!attribute
    unsigned_int_attr_accessor :fillId

    # index (0 based) of the border to be used in this style
    # @return [Integer]
    # @!attribute
    unsigned_int_attr_accessor :borderId

    # index (0 based) of cellStylesXfs item to be used in this style. Only applies to cellXfs items
    # @return [Integer]
    # @!attribute
    unsigned_int_attr_accessor :xfId

    # indecates if text should be prefixed by a single quote in the cell
    # @return [Boolean]
    # @!attribute
    boolean_attr_accessor :quotePrefix

    # indicates if the cell has a pivot table drop down button
    # @return [Boolean]
    # @!attribute
    boolean_attr_accessor :pivotButton

    # indicates if the numFmtId should be applied
    # @return [Boolean]
    # @!attribute
    boolean_attr_accessor :applyNumberFormat

    # indicates if the fontId should be applied
    # @return [Boolean]
    # @!attribute
    boolean_attr_accessor :applyFont

    # indicates if the fillId should be applied
    # @return [Boolean]
    # @!attribute
    boolean_attr_accessor :applyFill

    # indicates if the borderId should be applied
    # @return [Boolean]
    # @!attribute
    boolean_attr_accessor :applyBorder

    # Indicates if the alignment options should be applied
    # @return [Boolean]
    # @!attribute
    boolean_attr_accessor :applyAlignment

    # Indicates if the protection options should be applied
    # @return [Boolean]
    # @!attribute
    boolean_attr_accessor :applyProtection

    # @see Xf#alignment
    def alignment=(v)
      DataTypeValidator.validate "Xf.alignment", CellAlignment, v
      @alignment = v
    end

    # @see protection
    def protection=(v)
      DataTypeValidator.validate "Xf.protection", CellProtection, v
      @protection = v
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = +'')
      str << '<xf '
      serialized_attributes str
      str << '>'
      alignment.to_xml_string(str) if alignment
      protection.to_xml_string(str) if protection
      str << '</xf>'
    end
  end
end
