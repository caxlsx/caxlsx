

module Axlsx
  # This class performs sorting on a range in a worksheet
  class SortState
   # creates a new SortState object
    # @param [Worksheet] worksheet
    def initialize(worksheet, range)
      raise ArgumentError, 'you must provide a worksheet' unless worksheet.is_a?(Worksheet)

      @worksheet = worksheet
      @sort_conditions = []
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

    def apply
      first_cell, last_cell = range.split(':')
      start_point = Axlsx::name_to_indices(first_cell)
      end_point = Axlsx::name_to_indices(last_cell)
      # The +1 is so we skip the header row with the filter drop downs
      rows = worksheet.rows[(start_point.last + 1)..end_point.last] || []
      sorted_rows = rows.sort_by { |row| row.cells[2].value }
      sorted_rows
    end

    def to_xml_string(str = +'')
      str << "<sortState xmlns:xlrd2='http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2' ref='A2:D31'>"
      str << "<sortCondition descending='1' ref='C2:C31' />"
      str << "</sortState>"
    end
  end
end
