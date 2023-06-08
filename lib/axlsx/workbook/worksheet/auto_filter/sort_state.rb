

module Axlsx
  # This class performs sorting on a range in a worksheet
  class SortState
   # creates a new SortState object
    # @param [Worksheet] worksheet
    def initialize(worksheet, range)
      raise ArgumentError, 'you must provide a worksheet' unless worksheet.is_a?(Worksheet)

      @worksheet = worksheet
    end

    # The range the sorting should be applied to.
    # This should be a string like 'A2:B8'.
    # Watch out! The header row should not be included in the range.
    # @return [String]
    attr_accessor :range
    attr_reader :sort_conditions
    attr_reader :worksheet

    def defined_name
      return unless range

      Axlsx.cell_range(range.split(':').collect { |name| worksheet.name_to_cell(name) })
    end

    def sort_condition(col_id)
      @sort_condition << SortCondition.new(col_id)
      # @sort_conditions.append(@sort_condition)
    end

    def to_xml_string(str = +'')
      str << "<sortState xmlns:xlrd2='http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2' ref='A2:D31'>"
      str << "<sortCondition descending='1' ref='C2:C31' />"
      str << "</sortState>"
    end
  end
end
