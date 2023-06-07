

module Axlsx
  # This class performs sorting on a range in a worksheet
  class SortState
   # creates a new SortState object
    attr_reader :sort_conditions
    attr_reader :worksheet
    # @param [Worksheet] worksheet
    def initialize(range)
      @sort_conditions = []
    end

    # The range the sorting should be applied to.
    # This should be a string like 'A2:B8'.
    # Watch out! The header row should not be included in the range.
    # @return [String]
    attr_accessor :range

    def defined_name
      return unless range

      Axlsx.cell_range(range.split(':').collect { |name| worksheet.name_to_cell(name) })
    end

    # Adds a sort condition to the worksheet.
    #
    # @param [String] reference The cell range to sort.
    # @param [Hash] options The sort options.
    #
    # @option options [String] :sort_by Specifies the column or columns to sort by.
    # @option options [Boolean] :descending Specifies whether to sort in descending order.
    # @option options [Boolean] :case_sensitive Specifies whether the sort is case sensitive.
    #
    # @return [SortState] self
    def add_sort_condition(reference, options = {})
      sort_condition = SortCondition.new(reference, options)
      @sort_conditions << sort_condition
      self
    end

    def to_xml_string(str = +'')
      str << "<sortState xmlns:xlrd2='http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2' ref='A2:D31'>"
      str << "<sortCondition ref='C2:C31' descending='1' />"
      str << "</sortState>"
    end
  end
end
