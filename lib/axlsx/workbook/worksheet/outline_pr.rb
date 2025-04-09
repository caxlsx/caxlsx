# frozen_string_literal: true

module Axlsx
  # The OutlinePr class manages serialization of a worksheet's outlinePr element, which provides various
  # options to control outlining.
  class OutlinePr
    include Axlsx::OptionsParser
    include Axlsx::Accessors
    include Axlsx::SerializedAttributes

    serializable_attributes :summary_below,
                            :summary_right,
                            :apply_styles

    # These attributes are all boolean so I'm doing a bit of a hand
    # waving magic show to set up the attribute accessors
    boolean_attr_accessor :summary_below,
                          :summary_right,
                          :apply_styles

    # Creates a new OutlinePr object
    # @param [Hash] options used to create the outline_pr
    def initialize(options = {})
      parse_options options
    end

    # Serialize the object
    # @param [#<<] str A String, buffer or IO to append the serialization to. serialized output will be appended to this object if provided.
    # @return [void]
    def to_xml_string(str = +'')
      str << '<outlinePr '
      serialized_attributes(str)
      str << '/>'
    end
  end
end
