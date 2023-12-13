# frozen_string_literal: true

module Axlsx
  # The Font class details a font instance for use in styling cells.
  # @note The recommended way to manage fonts, and other styles is Styles#add_style
  # @see Styles#add_style
  class Font
    include Axlsx::Accessors
    include Axlsx::OptionsParser

    # Creates a new Font
    # @option options [String] name
    # @option options [Integer] charset
    # @option options [Integer] family
    # @option options [Integer] family
    # @option options [Boolean] b
    # @option options [Boolean] i
    # @option options [Boolean] u
    # @option options [Boolean] strike
    # @option options [Boolean] outline
    # @option options [Boolean] shadow
    # @option options [Boolean] condense
    # @option options [Boolean] extend
    # @option options [Color] color
    # @option options [Integer] sz
    def initialize(options = {})
      parse_options options
    end

    # The name of the font
    # @return [String]
    # @!attribute
    string_attr_accessor :name

    # The charset of the font
    # @return [Integer]
    # @note
    #  The following values are defined in the OOXML specification and are OS dependant values
    #   0   ANSI_CHARSET
    #   1   DEFAULT_CHARSET
    #   2   SYMBOL_CHARSET
    #   77  MAC_CHARSET
    #   128 SHIFTJIS_CHARSET
    #   129 HANGUL_CHARSET
    #   130 JOHAB_CHARSET
    #   134 GB2312_CHARSET
    #   136 CHINESEBIG5_CHARSET
    #   161 GREEK_CHARSET
    #   162 TURKISH_CHARSET
    #   163 VIETNAMESE_CHARSET
    #   177 HEBREW_CHARSET
    #   178 ARABIC_CHARSET
    #   186 BALTIC_CHARSET
    #   204 RUSSIAN_CHARSET
    #   222 THAI_CHARSET
    #   238 EASTEUROPE_CHARSET
    #   255 OEM_CHARSET
    # @!attribute
    unsigned_int_attr_accessor :charset

    # The font's family
    # @note
    #  The following are defined OOXML specification
    #   0 Not applicable.
    #   1 Roman
    #   2 Swiss
    #   3 Modern
    #   4 Script
    #   5 Decorative
    #   6..14 Reserved for future use
    # @return [Integer]
    # @!attribute
    unsigned_int_attr_accessor :family

    # Indicates if the font should be rendered in *bold*
    # @return [Boolean]
    # @!attribute
    boolean_attr_accessor :b

    # Indicates if the font should be rendered italicized
    # @return [Boolean]
    # @!attribute
    boolean_attr_accessor :i

    # Indicates if the font should be rendered underlined
    # It must be one of :none, :single, :double, :singleAccounting, :doubleAccounting, true, false
    # @return [String]
    # @note
    #  true or false is for backwards compatibility and is reassigned to :single or :none respectively
    attr_reader :u

    # Indicates if the font should be rendered with a strikthrough
    # @return [Boolean]
    # @!attribute
    boolean_attr_accessor :strike

    # Indicates if the font should be rendered with an outline
    # @return [Boolean]
    # @!attribute
    boolean_attr_accessor :outline

    # Indicates if the font should be rendered with a shadow
    # @return [Boolean]
    # @!attribute
    boolean_attr_accessor :shadow

    # Indicates if the font should be condensed
    # @return [Boolean]
    # @!attribute
    boolean_attr_accessor :condense

    # The font's extend property
    # @return [Boolean]
    # @!attribute
    boolean_attr_accessor :extend

    # The color of the font
    # @return [Color]
    attr_reader :color

    # The size of the font.
    # @return [Integer]
    # @!attribute
    unsigned_int_attr_accessor :sz

    # @see u
    def u=(v)
      v = :single if v == true || v == 1 || v == :true || v == 'true'
      v = :none if v == false || v == 0 || v == :false || v == 'false'
      Axlsx.validate_cell_u v
      @u = v
    end

    # @see color
    def color=(v)
      DataTypeValidator.validate "Font.color", Color, v
      @color = v
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = +'')
      str << '<font>'
      Axlsx.instance_values_for(self).each do |k, v|
        v.is_a?(Color) ? v.to_xml_string(str) : (str << '<' << k.to_s << ' val="' << Axlsx.booleanize(v).to_s << '"/>')
      end
      str << '</font>'
    end
  end
end
