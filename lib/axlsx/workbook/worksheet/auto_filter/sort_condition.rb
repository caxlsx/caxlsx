# frozen_string_literal: true

module Axlsx
  # This class represents a individual sort condition belonging to the sort state of an auto filter
  class SortCondition
    # Creates a new SortCondition object
    # @param [Integer] column_index Zero-based index indicating the AutoFilter column to which the sorting should be applied to
    # @param [Symbol] order The order the column should be sorted on, can only be :asc or :desc
    # @param [Array] custom_list An array containing a custom sorting list in order.
    def initialize(column_index:, order:, custom_list:)
      Axlsx.validate_int column_index
      @column_index = column_index

      RestrictionValidator.validate 'SortCondition.order', [:asc, :desc], order
      @order = order

      DataTypeValidator.validate :sort_condition_custom_list, Array, custom_list
      @custom_list = custom_list
    end

    attr_reader :column_index, :order, :custom_list

    # converts the ref String from the sort_state to a string representing the ref of a single column
    # for the xml string to be returned.
    def ref_to_single_column(ref, column_index)
      first_cell, last_cell = ref.split(':')

      start_point = Axlsx.name_to_indices(first_cell)

      first_row = first_cell[/\d+/]
      last_row = last_cell[/\d+/]

      first_column = Axlsx.col_ref(column_index + start_point.first)
      last_column = first_column

      "#{first_column}#{first_row}:#{last_column}#{last_row}"
    end

    # serialize the object
    # @return [String]
    def to_xml_string(str, ref)
      ref = ref_to_single_column(ref, column_index)

      str << "<sortCondition "
      str << "descending='1' " if order == :desc
      str << "ref='#{ref}' "
      str << "customList='#{custom_list.join(',')}' " unless custom_list.empty?
      str << "/>"
    end
  end
end
