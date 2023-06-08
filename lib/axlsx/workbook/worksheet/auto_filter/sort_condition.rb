module Axlsx

  class SortCondition

    include Axlsx::OptionsParser
    include Axlsx::SerializedAttributes

    def initialize(col_id, descending = false, options = {})
      @col_id = col_id
      @descending = descending
      @options = options
    end

    attr_reader :col_id
    attr_accessor :sort_conditions_array

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
