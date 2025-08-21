# frozen_string_literal: true

module Axlsx
  require_relative 'worksheet/sheet_calc_pr'
  require_relative 'worksheet/auto_filter/auto_filter'
  require_relative 'worksheet/date_time_converter'
  require_relative 'worksheet/protected_range'
  require_relative 'worksheet/protected_ranges'
  require_relative 'worksheet/rich_text_run'
  require_relative 'worksheet/rich_text'
  require_relative 'worksheet/cell_serializer'
  require_relative 'worksheet/cell'
  require_relative 'worksheet/page_margins'
  require_relative 'worksheet/page_set_up_pr'
  require_relative 'worksheet/outline_pr'
  require_relative 'worksheet/page_setup'
  require_relative 'worksheet/header_footer'
  require_relative 'worksheet/print_options'
  require_relative 'worksheet/cfvo'
  require_relative 'worksheet/cfvos'
  require_relative 'worksheet/color_scale'
  require_relative 'worksheet/data_bar'
  require_relative 'worksheet/icon_set'
  require_relative 'worksheet/conditional_formatting'
  require_relative 'worksheet/conditional_formatting_rule'
  require_relative 'worksheet/conditional_formattings'
  require_relative 'worksheet/row'
  require_relative 'worksheet/col'
  require_relative 'worksheet/cols'
  require_relative 'worksheet/comments'
  require_relative 'worksheet/comment'
  require_relative 'worksheet/merged_cells'
  require_relative 'worksheet/sheet_protection'
  require_relative 'worksheet/sheet_pr'
  require_relative 'worksheet/dimension'
  require_relative 'worksheet/sheet_data'
  require_relative 'worksheet/worksheet_drawing'
  require_relative 'worksheet/worksheet_comments'
  require_relative 'worksheet/worksheet_hyperlink'
  require_relative 'worksheet/worksheet_hyperlinks'
  require_relative 'worksheet/break'
  require_relative 'worksheet/row_breaks'
  require_relative 'worksheet/col_breaks'
  require_relative 'workbook_view'
  require_relative 'workbook_views'
  require_relative 'worksheet/worksheet'
  require_relative 'shared_strings_table'
  require_relative 'defined_name'
  require_relative 'defined_names'
  require_relative 'worksheet/table_style_info'
  require_relative 'worksheet/table'
  require_relative 'worksheet/tables'
  require_relative 'worksheet/pivot_table_cache_definition'
  require_relative 'worksheet/pivot_table'
  require_relative 'worksheet/pivot_tables'
  require_relative 'worksheet/data_validation'
  require_relative 'worksheet/data_validations'
  require_relative 'worksheet/sheet_view'
  require_relative 'worksheet/sheet_format_pr'
  require_relative 'worksheet/pane'
  require_relative 'worksheet/selection'

  # The Workbook class is an xlsx workbook that manages worksheets, charts, drawings and styles.
  # The following parts of the Office Open XML spreadsheet specification are not implemented in this version.
  #
  #   bookViews
  #   calcPr
  #   customWorkbookViews
  #   definedNames
  #   externalReferences
  #   extLst
  #   fileRecoveryPr
  #   fileSharing
  #   fileVersion
  #   functionGroups
  #   oleSize
  #   pivotCaches
  #   smartTagPr
  #   smartTagTypes
  #   webPublishing
  #   webPublishObjects
  #   workbookProtection
  #   workbookPr*
  #
  #   *workbookPr is only supported to the extend of date1904
  class Workbook
    BOLD_FONT_MULTIPLIER = 1.5
    FONT_SCALE_DIVISOR = 10.0

    # When true, the Package will be generated with a shared string table. This may be required by some OOXML processors that do not
    # adhere to the ECMA specification that dictates string may be inline in the sheet.
    # Using this option will increase the time required to serialize the document as every string in every cell must be analzed and referenced.
    # @return [Boolean]
    attr_reader :use_shared_strings

    # @see use_shared_strings
    def use_shared_strings=(v)
      Axlsx.validate_boolean(v)
      @use_shared_strings = v
    end

    # If true reverse the order in which the workbook is serialized
    # @return [Boolean]
    attr_reader :is_reversed

    def is_reversed=(v)
      Axlsx.validate_boolean(v)
      @is_reversed = v
    end

    # A collection of worksheets associated with this workbook.
    # @note The recommended way to manage worksheets is add_worksheet
    # @see Workbook#add_worksheet
    # @see Worksheet
    # @return [SimpleTypedList]
    attr_reader :worksheets

    # A collection of charts associated with this workbook
    # @note The recommended way to manage charts is Worksheet#add_chart
    # @see Worksheet#add_chart
    # @see Chart
    # @return [SimpleTypedList]
    attr_reader :charts

    # A collection of images associated with this workbook
    # @note The recommended way to manage images is Worksheet#add_image
    # @see Worksheet#add_image
    # @see Pic
    # @return [SimpleTypedList]
    attr_reader :images

    # A collection of drawings associated with this workbook
    # @note The recommended way to manage drawings is Worksheet#add_chart
    # @see Worksheet#add_chart
    # @see Drawing
    # @return [SimpleTypedList]
    attr_reader :drawings

    # pretty sure this two are always empty and can be removed.

    # A collection of tables associated with this workbook
    # @note The recommended way to manage drawings is Worksheet#add_table
    # @see Worksheet#add_table
    # @see Table
    # @return [SimpleTypedList]
    attr_reader :tables

    # A collection of pivot tables associated with this workbook
    # @note The recommended way to manage drawings is Worksheet#add_table
    # @see Worksheet#add_table
    # @see Table
    # @return [SimpleTypedList]
    attr_reader :pivot_tables

    # A collection of views for this workbook
    def views
      @views ||= WorkbookViews.new
    end

    # A collection of defined names for this workbook
    # @note The recommended way to manage defined names is Workbook#add_defined_name
    # @see DefinedName
    # @return [DefinedNames]
    def defined_names
      @defined_names ||= DefinedNames.new
    end

    # A collection of comments associated with this workbook
    # @note The recommended way to manage comments is WOrksheet#add_comment
    # @see Worksheet#add_comment
    # @see Comment
    # @return [Comments]
    def comments
      worksheets.map(&:comments).compact
    end

    # The styles associated with this workbook
    # @note The recommended way to manage styles is Styles#add_style
    # @see Style#add_style
    # @see Style
    # @return [Styles]
    def styles
      yield @styles if block_given?
      @styles
    end

    # The theme associated with this workbook
    # @return [Theme]
    def theme
      @theme ||= Theme.new
    end

    # An array that holds all cells with styles
    # @return Set
    def styled_cells
      @styled_cells ||= Set.new
    end

    # Are the styles added with workbook.add_styles applied yet
    # @return Boolean
    attr_accessor :styles_applied

    # A helper to apply styles that were added using `worksheet.add_style`
    # @return [Boolean]
    def apply_styles
      return false unless styled_cells

      styled_cells.each do |cell|
        current_style = styles.style_index[cell.style]

        new_style = if current_style
                      Axlsx.hash_deep_merge(current_style, cell.raw_style)
                    else
                      cell.raw_style
                    end

        cell.style = styles.add_style(new_style)
      end

      self.styles_applied = true
    end

    # Indicates if the epoc date for serialization should be 1904. If false, 1900 is used.
    @@date1904 = false

    # A quick helper to retrieve a worksheet by name
    # @param [String] name The name of the sheet you are looking for
    # @return [Worksheet] The sheet found, or nil
    def sheet_by_name(name)
      encoded_name = Axlsx.coder.encode(name)
      @worksheets.find { |sheet| sheet.name == encoded_name }
    end

    # Creates a new Workbook.
    # The recommended way to work with workbooks is via Package#workbook.
    # @option options [Boolean] date1904 If this is not specified, date1904 is set to false. Office 2011 for Mac defaults to false.
    def initialize(options = {})
      @styles = Styles.new
      @worksheets = SimpleTypedList.new Worksheet
      @drawings = SimpleTypedList.new Drawing
      @charts = SimpleTypedList.new Chart
      @images = SimpleTypedList.new Pic
      # Are these even used????? Check package serialization parts
      @tables = SimpleTypedList.new Table
      @pivot_tables = SimpleTypedList.new PivotTable
      @comments = SimpleTypedList.new Comments
      @use_autowidth = true
      @bold_font_multiplier = BOLD_FONT_MULTIPLIER
      @font_scale_divisor = FONT_SCALE_DIVISOR

      self.escape_formulas = options[:escape_formulas].nil? ? Axlsx.escape_formulas : options[:escape_formulas]
      self.date1904 = !options[:date1904].nil? && options[:date1904]
      yield self if block_given?
    end

    # Instance level access to the class variable 1904
    # @return [Boolean]
    def date1904
      @@date1904
    end

    # see @date1904
    def date1904=(v)
      Axlsx.validate_boolean v
      @@date1904 = v
    end

    # Sets the date1904 attribute to the provided boolean
    # @return [Boolean]
    def self.date1904=(v)
      Axlsx.validate_boolean v
      @@date1904 = v
    end

    # retrieves the date1904 attribute
    # @return [Boolean]
    def self.date1904
      @@date1904
    end

    # Whether to treat values starting with an equals sign as formulas or as literal strings.
    # Allowing user-generated data to be interpreted as formulas is a security risk.
    # See https://www.owasp.org/index.php/CSV_Injection for details.
    # @return [Boolean]
    attr_reader :escape_formulas

    # Sets whether to treat values starting with an equals sign as formulas or as literal strings.
    # @param [Boolean] value The value to set.
    def escape_formulas=(value)
      Axlsx.validate_boolean(value)
      @escape_formulas = value
    end

    # Indicates if the workbook should use autowidths or not.
    # @note This gem no longer depends on RMagick for autowidth
    #     calculation. Thus the performance benefits of turning this off are
    #     marginal unless you are creating a very large sheet.
    # @return [Boolean]
    attr_reader :use_autowidth

    # see @use_autowidth
    def use_autowidth=(v = true)
      Axlsx.validate_boolean v
      @use_autowidth = v
    end

    # Font size of bold fonts is multiplied with this
    # Used for automatic calculation of cell widths with bold text
    # @return [Float]
    attr_reader :bold_font_multiplier

    def bold_font_multiplier=(v)
      Axlsx.validate_float v
      @bold_font_multiplier = v
    end

    # Font scale is calculated with this value (font_size / font_scale_divisor)
    # Used for automatic calculation of cell widths
    # @return [Float]
    attr_reader :font_scale_divisor

    def font_scale_divisor=(v)
      Axlsx.validate_float v
      @font_scale_divisor = v
    end

    # inserts a worksheet into this workbook at the position specified.
    # It the index specified is out of range, the worksheet will be added to the end of the
    # worksheets collection
    # @return [Worksheet]
    # @param index The zero based position to insert the newly created worksheet
    # @param [Hash] options Options to pass into the worksheed during initialization.
    # @option options [String] name The name of the worksheet
    # @option options [Hash] page_margins The page margins for the worksheet
    def insert_worksheet(index = 0, options = {})
      worksheet = Worksheet.new(self, options)
      @worksheets.delete_at(@worksheets.size - 1)
      @worksheets.insert(index, worksheet)
      yield worksheet if block_given?
      worksheet
    end

    #
    # Adds a worksheet to this workbook
    # @return [Worksheet]
    # @option options [String] name The name of the worksheet.
    # @option options [Hash] page_margins The page margins for the worksheet.
    # @see Worksheet#initialize
    def add_worksheet(options = {})
      worksheet = Worksheet.new(self, options)
      yield worksheet if block_given?
      worksheet
    end

    # Adds a new WorkbookView
    # @return WorkbookViews
    # @option options [Hash] options passed into the added WorkbookView
    # @see WorkbookView#initialize
    def add_view(options = {})
      views << WorkbookView.new(options)
    end

    # Adds a defined name to this workbook
    # @return [DefinedName]
    # @param [String] formula @see DefinedName
    # @param [Hash] options @see DefinedName
    def add_defined_name(formula, options)
      defined_names << DefinedName.new(formula, options)
    end

    # The workbook relationships. This is managed automatically by the workbook
    # @return [Relationships]
    def relationships
      r = Relationships.new
      @worksheets.each do |sheet|
        r << Relationship.new(sheet, WORKSHEET_R, format(WORKSHEET_PN, r.size + 1))
      end
      pivot_tables.each_with_index do |pivot_table, index|
        r << Relationship.new(pivot_table.cache_definition, PIVOT_TABLE_CACHE_DEFINITION_R, format(PIVOT_TABLE_CACHE_DEFINITION_PN, index + 1))
      end
      r << Relationship.new(self, STYLES_R, STYLES_PN)
      r << Relationship.new(self, THEME_R, THEME_PN)
      if use_shared_strings
        r << Relationship.new(self, SHARED_STRINGS_R, SHARED_STRINGS_PN)
      end
      r
    end

    # generates a shared string object against all cells in all worksheets.
    # @return [SharedStringTable]
    def shared_strings
      SharedStringsTable.new(worksheets.collect(&:cells), xml_space)
    end

    # The xml:space attribute for the worksheet.
    # This determines how whitespace is handled within the document.
    # The most relevant part being whitespace in the cell text.
    # allowed values are :preserve and :default. Axlsx uses :preserve unless
    # you explicily set this to :default.
    # @return Symbol
    def xml_space
      @xml_space ||= :preserve
    end

    # Sets the xml:space attribute for the worksheet
    # @see Worksheet#xml_space
    # @param [Symbol] space must be one of :preserve or :default
    def xml_space=(space)
      Axlsx::RestrictionValidator.validate(:xml_space, [:preserve, :default], space)
      @xml_space = space
    end

    # returns a range of cells in a worksheet
    # @param [String] cell_def The Excel style reference defining the worksheet and cells. The range must specify the sheet to
    # retrieve the cells from. e.g. range('Sheet1!A1:B2') will return an array of four cells [A1, A2, B1, B2] while range('Sheet1!A1') will return a single Cell.
    # @return [Cell, Array]
    def [](cell_def)
      sheet_name = cell_def.split('!')[0] if cell_def.include?('!')
      worksheet =  worksheets.find { |s| s.name == sheet_name }
      raise ArgumentError, 'Unknown Sheet' unless sheet_name && worksheet.is_a?(Worksheet)

      worksheet[cell_def.gsub(/.+!/, "")]
    end

    # Serialize the workbook
    # @param [String] str
    # @return [String]
    def to_xml_string(str = +'')
      add_worksheet(name: 'Sheet1') if worksheets.empty?
      str << '<?xml version="1.0" encoding="UTF-8"?>'
      str << '<workbook xmlns="' << XML_NS << '" xmlns:r="' << XML_NS_R << '">'
      str << '<workbookPr date1904="' << @@date1904.to_s << '"/>'
      views.to_xml_string(str)
      str << '<sheets>'
      if is_reversed
        worksheets.reverse_each { |sheet| sheet.to_sheet_node_xml_string(str) }
      else
        worksheets.each { |sheet| sheet.to_sheet_node_xml_string(str) }
      end
      str << '</sheets>'
      defined_names.to_xml_string(str)
      unless pivot_tables.empty?
        str << '<pivotCaches>'
        pivot_tables.each do |pivot_table|
          str << '<pivotCache cacheId="' << pivot_table.cache_definition.cache_id.to_s << '" r:id="' << pivot_table.cache_definition.rId << '"/>'
        end
        str << '</pivotCaches>'
      end
      str << '</workbook>'
    end
  end
end
