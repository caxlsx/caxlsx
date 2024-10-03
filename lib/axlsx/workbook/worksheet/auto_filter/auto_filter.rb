# frozen_string_literal: true

require_relative 'filter_column'
require_relative 'filters'
require_relative 'sort_state'

module Axlsx
  # This class represents an auto filter range in a worksheet
  class AutoFilter
    # creates a new Autofilter object
    # @param [Worksheet] worksheet
    def initialize(worksheet)
      raise ArgumentError, 'you must provide a worksheet' unless worksheet.is_a?(Worksheet)

      @worksheet = worksheet
      @sort_on_generate = true
    end

    attr_reader :worksheet, :sort_on_generate

    # The range the autofilter should be applied to.
    # This should be a string like 'A1:B8'
    # @return [String]
    attr_accessor :range

    # the formula for the defined name required for this auto filter
    # This prepends the worksheet name to the absolute cell reference
    # e.g. A1:B2 -> 'Sheet1'!$A$1:$B$2
    # @return [String]
    def defined_name
      return unless range

      Axlsx.cell_range(range.split(':').collect { |name| worksheet.name_to_cell(name) })
    end

    # A collection of filterColumns for this auto_filter
    # @return [SimpleTypedList]
    def columns
      @columns ||= SimpleTypedList.new FilterColumn
    end

    # Adds a filter column. This is the recommended way to create and manage filter columns for your autofilter.
    # In addition to the require id and type parameters, options will be passed to the filter column during instantiation.
    # @param [String] col_id Zero-based index indicating the AutoFilter column to which this filter information applies.
    # @param [Symbol] filter_type A symbol representing one of the supported filter types.
    # @param [Hash] options a hash of options to pass into the generated filter
    # @return [FilterColumn]
    def add_column(col_id, filter_type, options = {})
      columns << FilterColumn.new(col_id, filter_type, options)
      columns.last
    end

    # Performs the sorting of the rows based on the sort_state conditions. Then it actually performs
    # the filtering of rows who's cells do not match the filter.
    def apply
      first_cell, last_cell = range.split(':')
      start_point = Axlsx.name_to_indices(first_cell)
      end_point = Axlsx.name_to_indices(last_cell)
      # The +1 is so we skip the header row with the filter drop downs
      rows = worksheet.rows[(start_point.last + 1)..end_point.last] || []

      # the sorting of the rows if sort_conditions are available.
      if !sort_state.sort_conditions.empty? && sort_on_generate
        sort_conditions = sort_state.sort_conditions
        sorted_rows = rows.sort do |row1, row2|
          comparison = 0

          sort_conditions.each do |condition|
            cell_value_row1 = row1.cells[condition.column_index + start_point.first].value
            cell_value_row2 = row2.cells[condition.column_index + start_point.first].value
            custom_list = condition.custom_list
            comparison = if cell_value_row1.nil? || cell_value_row2.nil?
                           cell_value_row1.nil? ? 1 : -1
                         elsif custom_list.empty?
                           condition.order == :asc ? cell_value_row1 <=> cell_value_row2 : cell_value_row2 <=> cell_value_row1
                         else
                           index1 = custom_list.index(cell_value_row1) || custom_list.size
                           index2 = custom_list.index(cell_value_row2) || custom_list.size

                           condition.order == :asc ? index1 <=> index2 : index2 <=> index1
                         end

            break unless comparison == 0
          end

          comparison
        end
        insert_index = start_point.last + 1

        sorted_rows.each do |row|
          # Insert the row at the specified index
          worksheet.rows[insert_index] = row
          insert_index += 1
        end
      end

      column_offset = start_point.first
      columns.each do |column|
        rows.each do |row|
          next if row.hidden

          column.apply(row, column_offset)
        end
      end
    end

    # the SortState object for this AutoFilter
    # @return [SortState]
    def sort_state
      @sort_state ||= SortState.new self
    end

    # @param [Boolean] v Flag indicating whether the AutoFilter should sort the rows when generating the
    # file. If false, the sorting rules will need to be applied manually after generating to alter
    # the order of the rows.
    # @return [Boolean]
    def sort_on_generate=(v)
      Axlsx.validate_boolean v
      @sort_on_generate = v
    end

    # serialize the object
    # @return [String]
    def to_xml_string(str = +'')
      return unless range

      str << "<autoFilter ref='#{range}'>"
      columns.each { |filter_column| filter_column.to_xml_string(str) }
      unless @sort_state.nil?
        @sort_state.to_xml_string(str)
      end
      str << "</autoFilter>"
    end
  end
end
