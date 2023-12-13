# frozen_string_literal: true

module Axlsx
  # A single table style definition and is a collection for tableStyleElements
  # @note Table are not supported in this version and only the defaults required for a valid workbook are created.
  class TableStyle < SimpleTypedList
    include Axlsx::Accessors
    include Axlsx::OptionsParser
    include Axlsx::SerializedAttributes

    # creates a new TableStyle object
    # @raise [ArgumentError] if name option is not provided.
    # @param [String] name
    # @option options [Boolean] pivot
    # @option options [Boolean] table
    def initialize(name, options = {})
      self.name = name
      parse_options options
      super(TableStyleElement)
    end

    serializable_attributes :name, :pivot, :table

    # The name of this table style
    # @return [string]
    # @!attribute
    string_attr_accessor :name

    # indicates if this style should be applied to pivot tables
    # @return [Boolean]
    # @!attribute
    boolean_attr_accessor :pivot

    # indicates if this style should be applied to tables
    # @return [Boolean]
    # @!attribute
    boolean_attr_accessor :table

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = +'')
      str << '<tableStyle '
      serialized_attributes str, { count: size }
      str << '>'
      each { |table_style_el| table_style_el.to_xml_string(str) }
      str << '</tableStyle>'
    end
  end
end
