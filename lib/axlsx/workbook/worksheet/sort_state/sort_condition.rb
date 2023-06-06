module Axlsx

  class SortCondition
    def initialize(col_id, ascending, options)
    end

    def ref
      column_to_sort
    end






    def to_xml_string(str = +'')
      # str << '<sortCondition '
      # serialized_attributes(str)
      # str << '>'
      # @filter.to_xml_string(str)
      # str << "</sortCondition>"
      '<sortCondition ref="B2:B277"></sortCondition>'
    end


  end
end
