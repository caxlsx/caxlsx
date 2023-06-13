module Axlsx

  class SortCondition

    # Creates a new SortCondition object
    # @param [Integer] col_id Zero-based index indicating the AutoFilter column to which the sorting should be applied to
    # @param [Boolean] a Boolean value if the order should be descending = true or false
    # @param [Array] An array containg a custom sorting list in order.
    def initialize(col_id, descending, custom_order)
      @col_id = col_id
      Axlsx::validate_boolean(descending)
      @descending = descending
      DataTypeValidator.validate :sort_condition_custom_order, Array, custom_order
      @custom_order = custom_order
    end

    attr_reader :col_id
    attr_reader :descending
    attr_reader :custom_order
    attr_accessor :sort_conditions_array

    # converts the ref String from the sort_state to a string representing the ref of a single column
    # for the xml string to be returned.
    def ref_to_single_column(ref, col_id)
      first_cell, last_cell = ref.split(':')

      first_row = first_cell[/\d+/]
      last_row = last_cell[/\d+/]

      first_column = get_column_letter(col_id)
      last_column = first_column

      modified_range = "#{first_column}#{first_row}:#{last_column}#{last_row}"
    end

    # Get the right letter representing the column from the col_id
    def get_column_letter(col_id)
      letters = []
      while col_id >= 0
        letters.unshift((col_id % 26) + 65)
        col_id = (col_id / 26).to_i - 1
      end
      letters.pack('C*')
    end

    # serialize the object
    # @return [String]
    def to_xml_string(str = +'', ref)
      ref = ref_to_single_column(ref, @col_id)

      str << '<sortCondition '
      str << 'descending="1" ' if @descending
      str << "ref='#{ref}' "
      str << "customList='#{@custom_order.join(',')}' " if @custom_order != []
      str << "/>"
    end
  end
end
