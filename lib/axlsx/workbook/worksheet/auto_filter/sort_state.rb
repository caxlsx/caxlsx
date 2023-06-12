require 'axlsx/workbook/worksheet/auto_filter/sort_condition.rb'

module Axlsx
  # This class performs sorting on a range in a worksheet
  class SortState
   # creates a new SortState object
    # @param [Worksheet] worksheet
    def initialize(auto_filter)
      @auto_filter = auto_filter

    end

    attr_reader :sort_conditions

    # The range the sorting should be applied to.
    # This should be a string like 'A2:B8'.
    # Watch out! The header row should not be included in the range.
    # @return [String]
    # attr_accessor :range
    # attr_accessor :sort_conditions_array
    # attr_reader :sort_conditions
    # attr_reader :worksheet

    def defined_name
      return unless range

      Axlsx.cell_range(@range.split(':').collect { |name| worksheet.name_to_cell(name) })
    end

    def sort_conditions
      @sort_conditions ||= SimpleTypedList.new SortCondition
    end

    def add_sort_condition(col_id, descending = false, options = {})
      sort_conditions << SortCondition.new(col_id, descending, options)
      sort_conditions.last
    end

    def to_xml_string(str = +'')
      str << "<sortState xmlns:xlrd2='http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2' ref='A2:D31'>"
      str << "<sortCondition descending='1' ref='C2:C31' />"
      str << "</sortState>"
    end
  end
end
