# frozen_string_literal: true

module Axlsx
  # A Row is a single row in a worksheet.
  # @note The recommended way to manage rows and cells is to use Worksheet#add_row
  # @see Worksheet#add_row
  class Row < SimpleTypedList
    include SerializedAttributes
    include Accessors

    # No support is provided for the following attributes
    # spans
    # thickTop
    # thickBottom

    # Creates a new row. New Cell objects are created based on the values, types and style options.
    # A new cell is created for each item in the values array. style and types options are applied as follows:
    #   If the types option is defined and is a symbol it is applied to all the cells created.
    #   If the types option is an array, cell types are applied by index for each cell
    #   If the types option is not set, the cell will automatically determine its type.
    #   If the style option is defined and is an Integer, it is applied to all cells created.
    #   If the style option is an array, style is applied by index for each cell.
    #   If the style option is not defined, the default style (0) is applied to each cell.
    # @param [Worksheet] worksheet
    # @option options [Array] values
    # @option options [Array, Symbol] types
    # @option options [Array, Integer] style
    # @option options [Array, Boolean] escape_formulas
    # @option options [Float] height the row's height (in points)
    # @option options [Integer] offset - add empty columns before values
    # @see Row#array_to_cells
    # @see Cell
    def initialize(worksheet, values = [], options = {})
      self.worksheet = worksheet
      super(Cell, nil, values.size + options[:offset].to_i)
      self.height = options.delete(:height)
      worksheet.rows << self
      array_to_cells(values, options)
    end

    # A list of serializable attributes.
    serializable_attributes :hidden, :outline_level, :collapsed, :custom_format, :s, :ph, :custom_height, :ht

    # Boolean row attribute accessors
    boolean_attr_accessor :hidden, :collapsed, :custom_format, :ph, :custom_height

    # The worksheet this row belongs to
    # @return [Worksheet]
    attr_reader :worksheet

    # Row height measured in point size. There is no margin padding on row height.
    # @return [Float]
    def height
      defined?(@ht) ? @ht : nil
    end

    # Outlining level of the row, when outlining is on
    # @return [Integer]
    attr_reader :outline_level
    alias :outlineLevel :outline_level

    # The style applied to the row. This affects the entire row.
    # @return [Integer]
    attr_reader :s

    # @see Row#s
    def s=(v)
      Axlsx.validate_unsigned_numeric(v)
      @custom_format = true
      @s = v
    end

    # @see Row#outline
    def outline_level=(v)
      Axlsx.validate_unsigned_numeric(v)
      @outline_level = v
    end

    alias :outlineLevel= :outline_level=

    # The index of this row in the worksheet
    # @return [Integer]
    def row_index
      worksheet.rows.index(self)
    end

    # Serializes the row
    # @param [Integer] r_index The row index, 0 based.
    # @param [#<<] str A String, buffer or IO to append the serialization to.
    # @return [void]
    def to_xml_string(r_index, str = +'')
      serialized_tag('row', str, r: Axlsx.row_ref(r_index)) do
        each_with_index { |cell, c_index| cell.to_xml_string(r_index, c_index, str) }
      end
    end

    # Adds a single cell to the row based on the data provided and updates the worksheet's autofit data.
    # @return [Cell]
    def add_cell(value = '', options = {})
      c = Cell.new(self, value, options)
      self << c
      worksheet.send(:update_column_info, self, [])
      c
    end

    # Sets the color for every cell in this row.
    def color=(color)
      each_with_index do |cell, index|
        cell.color = color.is_a?(Array) ? color[index] : color
      end
    end

    # Sets the style for every cell in this row.
    def style=(style)
      each_with_index do |cell, index|
        cell.style = style.is_a?(Array) ? style[index] : style
      end
    end

    # Sets escape_formulas for every cell in this row. This determines whether to treat
    # values starting with an equals sign as formulas or as literal strings.
    # @param [Array, Boolean] value The value to set.
    def escape_formulas=(value)
      each_with_index do |cell, index|
        cell.escape_formulas = value.is_a?(Array) ? value[index] : value
      end
    end

    # @see height
    def height=(v)
      unless v.nil?
        Axlsx.validate_unsigned_numeric(v)
        @custom_height = true
        @ht = v
      end
    end

    # return cells
    def cells
      self
    end

    private

    # assigns the owning worksheet for this row
    def worksheet=(v)
      DataTypeValidator.validate :row_worksheet, Worksheet, v
      @worksheet = v
    end

    # Converts values, types, and style options into cells and associates them with this row.
    # A new cell is created for each item in the values array.
    # If value option is defined and is a symbol it is applied to all the cells created.
    # If the value option is an array, cell types are applied by index for each cell
    # If the style option is defined and is an Integer, it is applied to all cells created.
    # If the style option is an array, style is applied by index for each cell.
    # @option options [Array] values
    # @option options [Array, Symbol] types
    # @option options [Array, Integer] style
    # @option options [Array, Boolean] escape_formulas
    def array_to_cells(values, options = {})
      DataTypeValidator.validate :array_to_cells, Array, values
      types, style, formula_values, escape_formulas, offset = options.delete(:types), options.delete(:style), options.delete(:formula_values), options.delete(:escape_formulas), options.delete(:offset)
      offset.to_i.times { |index| self[index] = Cell.new(self) } if offset
      values.each_with_index do |value, index|
        options[:style] = (style.is_a?(Array) ? style[index] : style) || worksheet.column_info[index]&.style
        options[:type] = types.is_a?(Array) ? types[index] : types if types
        options[:escape_formulas] = escape_formulas.is_a?(Array) ? escape_formulas[index] : escape_formulas unless escape_formulas.nil?
        options[:formula_value] = formula_values[index] if formula_values.is_a?(Array)

        self[index + offset.to_i] = Cell.new(self, value, options)
      end
    end
  end
end
