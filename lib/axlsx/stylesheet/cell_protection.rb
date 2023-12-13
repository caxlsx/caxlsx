# frozen_string_literal: true

module Axlsx
  # CellProtection stores information about locking or hiding cells in spreadsheet.
  # @note Using Styles#add_style is the recommended way to manage cell protection.
  # @see Styles#add_style
  class CellProtection
    include Axlsx::Accessors
    include Axlsx::OptionsParser
    include Axlsx::SerializedAttributes

    serializable_attributes :hidden, :locked

    # specifies locking for cells that have the style containing this protection
    # @return [Boolean]
    # @!attribute
    boolean_attr_accessor :hidden

    # specifies if the cells that have the style containing this protection
    # @return [Boolean]
    # @!attribute
    boolean_attr_accessor :locked

    # Creates a new CellProtection
    # @option options [Boolean] hidden value for hidden protection
    # @option options [Boolean] locked value for locked protection
    def initialize(options = {})
      parse_options options
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = +'')
      serialized_tag('protection', str)
    end
  end
end
