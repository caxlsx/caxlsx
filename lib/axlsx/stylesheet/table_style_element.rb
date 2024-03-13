# frozen_string_literal: true

module Axlsx
  # an element of style that belongs to a table style.
  # @note tables and table styles are not supported in this version. This class exists in preparation for that support.
  class TableStyleElement
    include Axlsx::Accessors
    include Axlsx::OptionsParser
    include Axlsx::SerializedAttributes

    # creates a new TableStyleElement object
    # @option options [Symbol] type
    # @option options [Integer] size
    # @option options [Integer] dxfId
    def initialize(options = {})
      parse_options options
    end

    serializable_attributes :type, :size, :dxfId

    # The type of style element. The following type are allowed
    #   :wholeTable
    #   :headerRow
    #   :totalRow
    #   :firstColumn
    #   :lastColumn
    #   :firstRowStripe
    #   :secondRowStripe
    #   :firstColumnStripe
    #   :secondColumnStripe
    #   :firstHeaderCell
    #   :lastHeaderCell
    #   :firstTotalCell
    #   :lastTotalCell
    #   :firstSubtotalColumn
    #   :secondSubtotalColumn
    #   :thirdSubtotalColumn
    #   :firstSubtotalRow
    #   :secondSubtotalRow
    #   :thirdSubtotalRow
    #   :blankRow
    #   :firstColumnSubheading
    #   :secondColumnSubheading
    #   :thirdColumnSubheading
    #   :firstRowSubheading
    #   :secondRowSubheading
    #   :thirdRowSubheading
    #   :pageFieldLabels
    #   :pageFieldValues
    # @return [Symbol]
    # @!attribute
    validated_attr_accessor :type, :validate_table_element_type

    # Number of rows or columns used in striping when the type is firstRowStripe, secondRowStripe, firstColumnStripe, or secondColumnStripe.
    # @return [Integer]
    # @!attribute
    unsigned_int_attr_accessor :size

    # The dxfId this style element points to
    # @return [Integer]
    # @!attribute
    unsigned_int_attr_accessor :dxfId

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = +'')
      serialized_tag('tableStyleElement', str)
    end
  end
end
