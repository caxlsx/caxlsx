# frozen_string_literal: true

module Axlsx
  # CellAlignment stores information about the cell alignment of a style Xf Object.
  # @note Using Styles#add_style is the recommended way to manage cell alignment.
  # @see Styles#add_style
  class CellAlignment
    include Axlsx::Accessors
    include Axlsx::SerializedAttributes
    include Axlsx::OptionsParser

    serializable_attributes :horizontal, :vertical, :text_rotation, :wrap_text, :indent, :relative_indent, :justify_last_line, :shrink_to_fit, :reading_order
    # Create a new cell_alignment object
    # @option options [Symbol] horizontal
    # @option options [Symbol] vertical
    # @option options [Integer] text_rotation
    # @option options [Boolean] wrap_text
    # @option options [Integer] indent
    # @option options [Integer] relative_indent
    # @option options [Boolean] justify_last_line
    # @option options [Boolean] shrink_to_fit
    # @option options [Integer] reading_order
    def initialize(options = {})
      parse_options options
    end

    # The horizontal alignment of the cell.
    # @note
    #  The horizontal cell alignement style must be one of
    #   :general
    #   :left
    #   :center
    #   :right
    #   :fill
    #   :justify
    #   :centerContinuous
    #   :distributed
    # @return [Symbol]
    # @!attribute
    validated_attr_accessor :horizontal, :validate_horizontal_alignment

    # The vertical alignment of the cell.
    # @note
    #  The vertical cell allingment style must be one of the following:
    #   :top
    #   :center
    #   :bottom
    #   :justify
    #   :distributed
    # @return [Symbol]
    # @!attribute
    validated_attr_accessor :vertical, :validate_vertical_alignment

    # The textRotation of the cell.
    # @return [Integer]
    # @!attribute
    unsigned_int_attr_accessor :text_rotation
    alias :textRotation :text_rotation
    alias :textRotation= :text_rotation=

    # Indicate if the text of the cell should wrap
    # @return [Boolean]
    # @!attribute
    boolean_attr_accessor :wrap_text
    alias :wrapText :wrap_text
    alias :wrapText= :wrap_text=

    # The amount of indent
    # @return [Integer]
    # @!attribute
    unsigned_int_attr_accessor :indent

    # The amount of relativeIndent
    # @return [Integer]
    # @!attribute
    int_attr_accessor :relative_indent
    alias :relativeIndent :relative_indent
    alias :relativeIndent= :relative_indent=

    # Indicate if the last line should be justified.
    # @return [Boolean]
    # @!attribute
    boolean_attr_accessor :justify_last_line
    alias :justifyLastLine :justify_last_line
    alias :justifyLastLine= :justify_last_line=

    # Indicate if the text should be shrunk to the fit in the cell.
    # @return [Boolean]
    # @!attribute
    boolean_attr_accessor :shrink_to_fit
    alias :shrinkToFit :shrink_to_fit
    alias :shrinkToFit= :shrink_to_fit=

    # The reading order of the text
    # 0 Context Dependent
    # 1 Left-to-Right
    # 2 Right-to-Left
    # @return [Integer]
    # @!attribute
    unsigned_int_attr_accessor :reading_order
    alias :readingOrder :reading_order
    alias :readingOrder= :reading_order=

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = +'')
      serialized_tag('alignment', str)
    end
  end
end
