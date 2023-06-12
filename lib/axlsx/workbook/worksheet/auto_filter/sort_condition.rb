module Axlsx

  class SortCondition

    include Axlsx::OptionsParser

    def initialize(col_id, descending = false, options = {})
      @col_id = col_id
      Axlsx::validate_boolean(descending)
      @descending = descending
      @options = options
    end

    attr_reader :col_id
    attr_reader :descending
    attr_reader :options
    attr_accessor :sort_conditions_array

    def ref_to_single_column(ref, col_id)
      first_cell, last_cell = ref.split(':')

      first_row = first_cell[/\d+/]
      last_row = last_cell[/\d+/]

      first_column = get_column_letter(col_id)
      last_column = first_column

      modified_range = "#{first_column}#{first_row}:#{last_column}#{last_row}"
    end

    def get_column_letter(col_id)
      letters = []
      while col_id >= 0
        letters.unshift((col_id % 26) + 65)
        col_id = (col_id / 26).to_i - 1
      end
      letters.pack('C*')
    end

    def to_xml_string(str = +'', ref)
      ref = ref_to_single_column(ref, @col_id)

      str << '<sortCondition '
      str << 'descending="1" ' if @descending
      str << "ref='#{ref}' />"
      # str << "customList='#{@options}' />"
    end


  end
end
