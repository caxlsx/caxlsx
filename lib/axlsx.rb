# frozen_string_literal: true

require_relative 'axlsx/version'

# gemspec dependencies
require 'htmlentities'
require 'marcel'
require 'nokogiri'
require 'zip'

# Ruby core dependencies
require 'cgi'
require 'set'
require 'time'
require 'uri'
require 'net/http'

require_relative 'axlsx/util/simple_typed_list'
require_relative 'axlsx/util/constants'
require_relative 'axlsx/util/validators'
require_relative 'axlsx/util/accessors'
require_relative 'axlsx/util/serialized_attributes'
require_relative 'axlsx/util/options_parser'
require_relative 'axlsx/util/uri_utils'
require_relative 'axlsx/util/mime_type_utils'
require_relative 'axlsx/util/buffered_zip_output_stream'
require_relative 'axlsx/util/zip_command'

require_relative 'axlsx/stylesheet/styles'

require_relative 'axlsx/doc_props/app'
require_relative 'axlsx/doc_props/core'
require_relative 'axlsx/content_type/content_type'
require_relative 'axlsx/rels/relationships'

require_relative 'axlsx/drawing/drawing'
require_relative 'axlsx/workbook/workbook'
require_relative 'axlsx/package'

if Gem.loaded_specs.key?("axlsx_styler")
  raise StandardError, "Please remove `axlsx_styler` from your Gemfile, the associated functionality is now built-in to `caxlsx` directly."
end

# xlsx generation with charts, images, automated column width, customizable styles
# and full schema validation. Axlsx excels at helping you generate beautiful
# Office Open XML Spreadsheet documents without having to understand the entire
# ECMA specification. Check out the README for some examples of how easy it is.
# Best of all, you can validate your xlsx file before serialization so you know
# for sure that anything generated is going to load on your client's machine.
module Axlsx
  # I am a very big fan of activesupports instance_values method, but do not want to require nor include the entire
  # library just for this one method.
  #
  # Defining as a class method on Axlsx to refrain from monkeypatching Object for all users of this gem.
  def self.instance_values_for(object)
    object.instance_variables.to_h { |name| [name.to_s[1..], object.instance_variable_get(name)] }
  end

  # determines the cell range for the items provided
  def self.cell_range(cells, absolute = true)
    return "" unless cells.first.is_a? Cell

    first_cell, last_cell = cells.minmax_by(&:pos)
    reference = "#{first_cell.reference(absolute)}:#{last_cell.reference(absolute)}"
    if absolute
      escaped_name = first_cell.row.worksheet.name.gsub '&apos;', "''"
      "'#{escaped_name}'!#{reference}"
    else
      reference
    end
  end

  # sorts the array of cells provided to start from the minimum x,y to
  # the maximum x.y#
  # @param [Array] cells
  # @return [Array]
  def self.sort_cells(cells)
    cells.sort_by(&:pos)
  end

  # global reference html entity encoding
  # @return [HtmlEntities]
  def self.coder
    @@coder ||= ::HTMLEntities.new
  end

  # returns the x, y position of a cell
  def self.name_to_indices(name)
    raise ArgumentError, 'invalid cell name' unless name.size > 1

    letters_str = name[/[A-Z]+/]

    # capitalization?!?
    v = letters_str.reverse.chars.each_with_object({ base: 1, i: 0 }) do |c, val|
      val[:i] += ((c.bytes.first - 64) * val[:base])

      val[:base] *= 26
    end

    col_index = (v[:i] - 1)

    numbers_str = name[/[1-9][0-9]*/]

    row_index = (numbers_str.to_i - 1)

    [col_index, row_index]
  end

  # converts the column index into alphabetical values.
  # @note This follows the standard spreadsheet convention of naming columns A to Z, followed by AA to AZ etc.
  # @return [String]
  def self.col_ref(index)
    # Every row will call this for each column / cell and so we can cache result and avoid lots of small object
    # allocations.
    @col_ref ||= {}
    @col_ref[index] ||= begin
      i = index
      chars = +''
      while i >= 26
        i, char = i.divmod(26)
        chars.prepend((char + 65).chr)
        i -= 1
      end
      chars.prepend((i + 65).chr)
      chars.freeze
    end
  end

  # converts the row index into string values.
  # @note The spreadsheet rows are 1-based and the passed in index is 0-based, so we add 1.
  # @return [String]
  def self.row_ref(index)
    @row_ref ||= {}
    @row_ref[index] ||= (index + 1).to_s.freeze
  end

  # @return [String] The alpha(column)numeric(row) reference for this sell.
  # @example Relative Cell Reference
  #   ws.rows.first.cells.first.r #=> "A1"
  def self.cell_r(c_index, r_index)
    col_ref(c_index) + row_ref(r_index)
  end

  # Creates an array of individual cell references based on an Excel reference range.
  # @param [String] range A cell range, for example A1:D5
  # @return [Array]
  def self.range_to_a(range)
    range =~ /^(\w+?\d+):(\w+?\d+)$/
    start_col, start_row = name_to_indices(::Regexp.last_match(1))
    end_col,   end_row   = name_to_indices(::Regexp.last_match(2))
    (start_row..end_row).to_a.map do |row_num|
      (start_col..end_col).to_a.map do |col_num|
        cell_r(col_num, row_num)
      end
    end
  end

  # performs the incredible feat of changing snake_case to CamelCase
  # @param [String] s The snake case string to camelize
  # @return [String]
  def self.camel(s = "", all_caps = true)
    s = s.to_s
    s = s.capitalize if all_caps
    s.gsub(/_(.)/) { ::Regexp.last_match(1).upcase }
  end

  # returns the provided string with all invalid control characters
  # removed.
  # @param [String] str The string to process
  # @return [String]
  def self.sanitize(str)
    if str.frozen?
      str.delete(CONTROL_CHARS)
    else
      str.delete!(CONTROL_CHARS)
      str
    end
  end

  # If value is boolean return 1 or 0
  # else return the value
  # @param [Object] value The value to process
  # @return [Object]
  def self.booleanize(value)
    if BOOLEAN_VALUES.include?(value)
      value ? '1' : '0'
    else
      value
    end
  end

  # utility method for performing a deep merge on a Hash
  # @param [Hash] first_hash Hash to merge into
  # @param [Hash] second_hash Hash to be added
  def self.hash_deep_merge(first_hash, second_hash)
    first_hash.merge(second_hash) do |_key, this_val, other_val|
      if this_val.is_a?(Hash) && other_val.is_a?(Hash)
        Axlsx.hash_deep_merge(this_val, other_val)
      else
        other_val
      end
    end
  end

  # Instructs the serializer to not try to escape cell value input.
  # This will give you a huge speed bonus, but if you content has <, > or other xml character data
  # the workbook will be invalid and Excel will complain.
  def self.trust_input
    @trust_input ||= false
  end

  # @param[Boolean] trust_me A boolean value indicating if the cell value content is to be trusted
  # @return [Boolean]
  # @see Axlsx::trust_input
  def self.trust_input=(trust_me)
    @trust_input = trust_me
  end

  # Whether to treat values starting with an equals sign as formulas or as literal strings.
  # Allowing user-generated data to be interpreted as formulas is a security risk.
  # See https://www.owasp.org/index.php/CSV_Injection for details.
  # @return [Boolean]
  def self.escape_formulas
    !defined?(@escape_formulas) || @escape_formulas.nil? ? true : @escape_formulas
  end

  # Sets whether to treat values starting with an equals sign as formulas or as literal strings.
  # @param [Boolean] value The value to set.
  def self.escape_formulas=(value)
    Axlsx.validate_boolean(value)
    @escape_formulas = value
  end

  # Returns a URI parser instance, preferring RFC2396_PARSER if available,
  # otherwise falling back to DEFAULT_PARSER. This method ensures consistent
  # URI parsing across different Ruby versions.
  # This method can be removed when dropping compatibility for Ruby < 3.4
  # See https://github.com/ruby/uri/pull/114 for details.
  # @return [Object]
  def self.uri_parser
    @uri_parser ||=
      if defined?(URI::RFC2396_PARSER)
        URI::RFC2396_PARSER
      else
        URI::DEFAULT_PARSER
      end
  end
end
