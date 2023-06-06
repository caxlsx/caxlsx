

module Axlsx
  # This class performs sorting on a range in a worksheet
  class SortState
   # creates a new SortState object
    # @param [Worksheet] worksheet
    def initialize(worksheet)
      raise ArgumentError, 'you must provide a worksheet' unless worksheet.is_a?(Worksheet)

      @worksheet = worksheet
    end

    attr_reader :worksheet

    # The range the sorting should be applied to.
    # This should be a string like 'A2:B8'.
    # Watch out! The header row should not be included in the range.
    # @return [String]
    attr_accessor :range

    def defined_name
      return unless range

      Axlsx.cell_range(range.split(':').collect { |name| worksheet.name_to_cell(name) })
    end

    def to_xml_string
      str << "<sortState xmlns:xlrd2='http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2' ref='A2:D31'>"
      str << "<sortCondition ref='C2:C31' descending='1' />"
      str << "</sortState>"
    end
  end
end
