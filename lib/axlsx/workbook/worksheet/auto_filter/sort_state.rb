require 'axlsx/workbook/worksheet/auto_filter/sort_condition.rb'

module Axlsx
  # This class performs sorting on a range in a worksheet
  class SortState
   # creates a new SortState object
    # @param [Worksheet] worksheet
    def initialize(range)
      @range = range

    end

    # The range the sorting should be applied to.
    # This should be a string like 'A2:B8'.
    # Watch out! The header row should not be included in the range.
    # @return [String]
    attr_accessor :range
    attr_accessor :sort_conditions_array
    attr_reader :sort_conditions
    attr_reader :worksheet

    def defined_name
      return unless range

      Axlsx.cell_range(@range.split(':').collect { |name| worksheet.name_to_cell(name) })
    end

    def sort_conditions
      @sort_conditions = []
      byebug
      sort_conditions_array.each do |sort_condition_array|
        sort_condition ||= SortCondition.new sort_condition_array[0], sort_condition_array[1], sort_condition_array[2]
        @sort_conditions.append(sort_condition)
      end
      @sort_conditions
    end

    def sort_conditions=(v)
      # DataTypeValidator.validate :sort_state_sort_conditions, Array, v
      sort_conditions.sort_conditions_array = v
    end

    def to_xml_string(str = +'')
      str << "<sortState xmlns:xlrd2='http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2' ref='A2:D31'>"
      str << "<sortCondition descending='1' ref='C2:C31' />"
      str << "</sortState>"
    end
  end
end
