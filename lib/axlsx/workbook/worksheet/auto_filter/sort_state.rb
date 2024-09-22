# frozen_string_literal: true

require_relative 'sort_condition'

module Axlsx
  # This class performs sorting on a range in a worksheet
  class SortState
    # creates a new SortState object
    # @param [AutoFilter] auto_filter the auto_filter that this sort_state belongs to
    def initialize(auto_filter)
      @auto_filter = auto_filter
    end

    # A collection of SortConditions for this sort_state
    # @return [SimpleTypedList]
    def sort_conditions
      @sort_conditions ||= SimpleTypedList.new SortCondition
    end

    # Adds a SortCondition to the sort_state. This is the recommended way to add conditions to it.
    # It requires a column_index for the sorting, descending and the custom order are optional.
    # @param [Integer] column_index Zero-based index indicating the AutoFilter column to which the sorting should be applied to
    # @param [Symbol] order The order the column should be sorted on, can only be :asc or :desc
    # @param [Array] custom_list An array containing a custom sorting list in order.
    # @return [SortCondition]
    def add_sort_condition(column_index:, order: :asc, custom_list: [])
      sort_conditions << SortCondition.new(column_index: column_index, order: order, custom_list: custom_list)
      sort_conditions.last
    end

    # method to increment the String representing the first cell of the range of the autofilter by 1 row for the sortCondition
    # xml string
    def increment_cell_value(str)
      letter = str[/[A-Za-z]+/]
      number = str[/\d+/].to_i

      incremented_number = number + 1

      "#{letter}#{incremented_number}"
    end

    # serialize the object
    # @return [String]
    def to_xml_string(str = +'')
      return if sort_conditions.empty?

      ref = @auto_filter.range
      first_cell, last_cell = ref.split(':')
      ref = "#{increment_cell_value(first_cell)}:#{last_cell}"

      str << "<sortState xmlns:xlrd2='http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2' ref='#{ref}'>"
      sort_conditions.each { |sort_condition| sort_condition.to_xml_string(str, ref) }
      str << "</sortState>"
    end
  end
end
