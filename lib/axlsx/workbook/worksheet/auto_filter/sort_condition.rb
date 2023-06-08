module Axlsx

  class SortCondition

    include Axlsx::OptionsParser
    include Axlsx::SerializedAttributes

    def initialize(col_id)
      self.col_id = col_id
      # @descending = descending
      # @options = options
    end

    attr_reader :col_id

    def to_xml_string(str = +'')
      # str << '<sortCondition '
      # serialized_attributes(str)
      # str << '>'
      # @filter.to_xml_string(str)
      # str << "</sortCondition>"
      '<sortCondition ref=''></sortCondition>'
    end


  end
end
