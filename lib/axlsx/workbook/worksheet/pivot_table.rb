# frozen_string_literal: true

module Axlsx
  # Table
  # @note Worksheet#add_pivot_table is the recommended way to create tables for your worksheets.
  # @see README for examples
  class PivotTable
    include Axlsx::OptionsParser

    # Creates a new PivotTable object
    # @param [String] ref The reference to where the pivot table lives like 'G4:L17'.
    # @param [String] range The reference to the pivot table data like 'A1:D31'.
    # @param [Worksheet] sheet The sheet containing the table data.
    # @option options [Cell, String] name
    # @option options [TableStyle] style
    def initialize(ref, range, sheet, options = {})
      @ref = ref
      self.range = range
      @sheet = sheet
      @sheet.workbook.pivot_tables << self
      @name = "PivotTable#{index + 1}"
      @data_sheet = nil
      @rows = []
      @columns = []
      @data = []
      @pages = []
      @subtotal = nil
      @no_subtotals_on_headers = []
      @grand_totals = :both
      @sort_on_headers = {}
      @style_info = {}
      parse_options options
      yield self if block_given?
    end

    # Defines the headers in which subtotals are not to be included.
    # @return [Array]
    attr_accessor :no_subtotals_on_headers

    # Defines the headers in which sort is applied.
    # Can be an array of headers to sort ascending by default, or a hash for specific control
    # (with headers as keys, `:ascending` or `:descending` as values).
    #
    # Examples: `["year", "month"]` or `{"year" => :descending, "month" => :descending}`
    # @return [Hash]
    attr_reader :sort_on_headers

    # (see #sort_on_headers)
    def sort_on_headers=(headers)
      headers ||= {}
      headers = Hash[*headers.map { |h| [h, :ascending] }.flatten] if headers.is_a?(Array)
      @sort_on_headers = headers
    end

    # Defines which Grand Totals are to be shown.
    # @return [Symbol] The row and/or column Grand Totals that are to be shown.
    # Defaults to `:both` to show both row & column grand totals.
    # Set to `:row_only`, `:col_only`, or `:none` to hide one or both Grand Totals.
    attr_reader :grand_totals

    # (see #grand_totals)
    def grand_totals=(value)
      RestrictionValidator.validate "PivotTable.grand_totals", [:both, :row_only, :col_only, :none], value

      @grand_totals = value
    end

    # Style info for the pivot table
    # @return [Hash]
    attr_accessor :style_info

    # The reference to the table data
    # @return [String]
    attr_reader :ref

    # The name of the table.
    # @return [String]
    attr_reader :name

    # The name of the sheet.
    # @return [String]
    attr_reader :sheet

    # The sheet used as data source for the pivot table
    # @return [Worksheet]
    attr_writer :data_sheet

    # @see #data_sheet
    def data_sheet
      @data_sheet || @sheet
    end

    # The range where the data for this pivot table lives.
    # @return [String]
    attr_reader :range

    # (see #range)
    def range=(v)
      DataTypeValidator.validate "#{self.class}.range", [String], v
      if v.is_a?(String)
        @range = v
      end
    end

    # The rows
    # @return [Array]
    attr_reader :rows

    # (see #rows)
    def rows=(v)
      DataTypeValidator.validate "#{self.class}.rows", [Array], v
      v.each do |ref|
        DataTypeValidator.validate "#{self.class}.rows[]", [String], ref
      end
      @rows = v
    end

    # The columns
    # @return [Array]
    attr_reader :columns

    # (see #columns)
    def columns=(v)
      DataTypeValidator.validate "#{self.class}.columns", [Array], v
      v.each do |ref|
        DataTypeValidator.validate "#{self.class}.columns[]", [String], ref
      end
      @columns = v
    end

    # The data as an array of either headers (String) or hashes or mix of the two.
    # Hash in format of { ref: header, num_fmt: numFmts, subtotal: subtotal }, where header is String, numFmts is Integer, and subtotal one of %w[sum count average max min product countNums stdDev stdDevp var varp]; leave subtotal blank to sum values
    # @return [Array]
    attr_reader :data

    # (see #data)
    def data=(v)
      DataTypeValidator.validate "#{self.class}.data", [Array], v
      @data = []
      v.each do |data_field|
        if data_field.is_a? String
          data_field = { ref: data_field }
        end
        data_field.each do |key, value|
          if key == :num_fmt
            DataTypeValidator.validate "#{self.class}.data[]", [Integer], value
          else
            DataTypeValidator.validate "#{self.class}.data[]", [String], value
          end
        end
        @data << data_field
      end
    end

    # The pages
    # @return [String]
    attr_reader :pages

    # (see #pages)
    def pages=(v)
      DataTypeValidator.validate "#{self.class}.pages", [Array], v
      v.each do |ref|
        DataTypeValidator.validate "#{self.class}.pages[]", [String], ref
      end
      @pages = v
    end

    # The index of this chart in the workbooks charts collection
    # @return [Integer]
    def index
      @sheet.workbook.pivot_tables.index(self)
    end

    # The part name for this table
    # @return [String]
    def pn
      format(PIVOT_TABLE_PN, index + 1)
    end

    # The relationship part name of this pivot table
    # @return [String]
    def rels_pn
      format(PIVOT_TABLE_RELS_PN, index + 1)
    end

    # The cache_definition for this pivot table
    # @return [PivotTableCacheDefinition]
    def cache_definition
      @cache_definition ||= PivotTableCacheDefinition.new(self)
    end

    # The relationships for this pivot table.
    # @return [Relationships]
    def relationships
      r = Relationships.new
      r << Relationship.new(cache_definition, PIVOT_TABLE_CACHE_DEFINITION_R, "../#{cache_definition.pn}")
      r
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = +'')
      str << '<?xml version="1.0" encoding="UTF-8"?>'

      str << '<pivotTableDefinition xmlns="' << XML_NS << '" name="' << name << '" cacheId="' << cache_definition.cache_id.to_s << '"'
      str << ' dataOnRows="1"' if data.size <= 1
      str << ' rowGrandTotals="0"' if grand_totals == :col_only || grand_totals == :none
      str << ' colGrandTotals="0"' if grand_totals == :row_only || grand_totals == :none
      str << ' applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Data" showMultipleLabel="0" showMemberPropertyTips="0" useAutoFormatting="1" indent="0" compact="0" compactData="0" gridDropZones="1" multipleFieldFilters="0">'

      str << '<location firstDataCol="1" firstDataRow="1" firstHeaderRow="1" ref="' << ref << '"/>'
      str << '<pivotFields count="' << header_cells_count.to_s << '">'

      header_cell_values.each do |cell_value|
        subtotal = !no_subtotals_on_headers.include?(cell_value)
        sorttype = sort_on_headers[cell_value]
        str << pivot_field_for(cell_value, subtotal, sorttype)
      end

      str << '</pivotFields>'
      if rows.empty?
        str << '<rowFields count="1"><field x="-2"/></rowFields>'
        str << '<rowItems count="2"><i><x/></i> <i i="1"><x v="1"/></i></rowItems>'
      else
        str << '<rowFields count="' << rows.size.to_s << '">'
        rows.each do |row_value|
          str << '<field x="' << header_index_of(row_value).to_s << '"/>'
        end
        str << '</rowFields>'
        str << '<rowItems count="' << rows.size.to_s << '">'
        rows.size.times do
          str << '<i/>'
        end
        str << '</rowItems>'
      end
      if columns.empty?
        if data.size > 1
          str << '<colFields count="1"><field x="-2"/></colFields>'
          str << "<colItems count=\"#{data.size}\">"
          str << '<i><x/></i>'
          (data.size - 1).times do |i|
            str << "<i i=\"#{i + 1}\"><x v=\"#{i + 1}\"/></i>"
          end
          str << '</colItems>'
        else
          str << '<colItems count="1"><i/></colItems>'
        end
      else
        str << '<colFields count="' << columns.size.to_s << '">'
        columns.each do |column_value|
          str << '<field x="' << header_index_of(column_value).to_s << '"/>'
        end
        str << '</colFields>'
      end
      unless pages.empty?
        str << '<pageFields count="' << pages.size.to_s << '">'
        pages.each do |page_value|
          str << '<pageField fld="' << header_index_of(page_value).to_s << '"/>'
        end
        str << '</pageFields>'
      end
      unless data.empty?
        str << "<dataFields count=\"#{data.size}\">"
        data.each do |datum_value|
          subtotal_name = datum_value[:subtotal] || 'sum'
          subtotal_name = 'count' if name == 'countNums' # both count & countNums are labelled as count
          str << "<dataField name='#{subtotal_name.capitalize} of #{datum_value[:ref]}' fld='#{header_index_of(datum_value[:ref])}' baseField='0' baseItem='0'"
          str << " numFmtId='#{datum_value[:num_fmt]}'" if datum_value[:num_fmt]
          str << " subtotal='#{datum_value[:subtotal]}' " if datum_value[:subtotal]
          str << "/>"
        end
        str << '</dataFields>'
      end
      # custom pivot table style
      unless style_info.empty?
        str << '<pivotTableStyleInfo'
        style_info.each do |k, v|
          str << ' ' << k.to_s << '="' << v.to_s << '"'
        end
        str << ' />'
      end
      str << '</pivotTableDefinition>'
    end

    # References for header cells
    # @return [Array]
    def header_cell_refs
      Axlsx.range_to_a(header_range).first
    end

    # The header cells for the pivot table
    # @return [Array]
    def header_cells
      data_sheet[header_range]
    end

    # The values in the header cells collection
    # @return [Array]
    def header_cell_values
      header_cells.map(&:value)
    end

    # The number of cells in the header_cells collection
    # @return [Integer]
    def header_cells_count
      header_cells.count
    end

    # The index of a given value in the header cells
    # @return [Integer]
    def header_index_of(value)
      header_cell_values.index(value)
    end

    private

    def pivot_field_for(cell_ref, subtotal, sorttype)
      attributes = %w[compact="0" outline="0" subtotalTop="0" showAll="0" includeNewItemsInFilter="1"]
      items_tag = '<items count="1"><item t="default"/></items>'
      include_items_tag = false

      if rows.include? cell_ref
        attributes << 'axis="axisRow"'
        attributes << "sortType=\"#{sorttype == :descending ? 'descending' : 'ascending'}\"" if sorttype
        if subtotal
          include_items_tag = true
        else
          attributes << 'defaultSubtotal="0"'
        end
      elsif columns.include? cell_ref
        attributes << 'axis="axisCol"'
        attributes << "sortType=\"#{sorttype == :descending ? 'descending' : 'ascending'}\"" if sorttype
        if subtotal
          include_items_tag = true
        else
          attributes << 'defaultSubtotal="0"'
        end
      elsif pages.include? cell_ref
        attributes << 'axis="axisPage"'
        include_items_tag = true
      elsif data_refs.include? cell_ref
        attributes << 'dataField="1"'
      end

      "<pivotField #{attributes.join(' ')}>#{include_items_tag ? items_tag : nil}</pivotField>"
    end

    def data_refs
      data.map { |hash| hash[:ref] }
    end

    def header_range
      range.gsub(/^(\w+?)(\d+)\:(\w+?)\d+$/, '\1\2:\3\2')
    end
  end
end
