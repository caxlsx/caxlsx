# frozen_string_literal: true

module Axlsx
  # A Title stores information about the title of a chart
  class Title
    # The text to be shown. Setting this property directly with a string will remove the cell reference.
    # @return [String]
    attr_reader :text

    # Text size property
    # @return [String]
    attr_reader :text_size

    # The cell that holds the text for the title. Setting this property will automatically update the text attribute.
    # @return [Cell]
    attr_reader :cell

    # Creates a new Title object
    # @param [String, Cell] title The cell or string to be used for the chart's title
    def initialize(title = "", title_size = "")
      self.cell = title if title.is_a?(Cell)
      self.text = title.to_s unless title.is_a?(Cell)
      self.text_size = if title_size.to_s.empty?
                         "1600"
                       else
                         title_size.to_s
                       end
    end

    # @see text
    def text=(v)
      DataTypeValidator.validate 'Title.text', String, v
      @text = v
      @cell = nil
    end

    # @see text_size
    def text_size=(v)
      DataTypeValidator.validate 'Title.text_size', String, v
      @text_size = v
      @cell = nil
    end

    # @see cell
    def cell=(v)
      DataTypeValidator.validate 'Title.text', Cell, v
      @cell = v
      @text = v.value.to_s
    end

    # Check if the title is empty.
    #
    # A title is considered empty if it is an empty string. If the title references a cell it is *not* empty,
    # even if the referenced cell is blank (because the cell’s value could still change later).
    #
    # @return [Boolean]
    def empty?
      @text.empty? && @cell.nil?
    end

    # Not implemented at this time.
    # def layout=(v) DataTypeValidator.validate 'Title.layout', Layout, v; @layout = v; end
    # def overlay=(v) Axlsx::validate_boolean v; @overlay=v; end
    # def spPr=(v) DataTypeValidator.validate 'Title.spPr', SpPr, v; @spPr = v; end

    # Serializes the object
    # @param [#<<] str A String, buffer or IO to append the serialization to.
    # @return [void]
    def to_xml_string(str = +'')
      str << '<c:title>'
      unless empty?
        clean_value = Axlsx.trust_input ? @text.to_s : ::CGI.escapeHTML(Axlsx.sanitize(@text.to_s))
        str << '<c:tx>'
        if @cell.is_a?(Cell)
          str << '<c:strRef>'
          str << '<c:f>' << Axlsx.cell_range([@cell]) << '</c:f>'
          str << '<c:strCache>'
          str << '<c:ptCount val="1"/>'
          str << '<c:pt idx="0">'
          str << '<c:v>' << clean_value << '</c:v>'
          str << '</c:pt>'
          str << '</c:strCache>'
          str << '</c:strRef>'
        else
          str << '<c:rich>'
          str << '<a:bodyPr/>'
          str << '<a:lstStyle/>'
          str << '<a:p>'
          str << '<a:r>'
          str << '<a:rPr sz="' << @text_size.to_s << '"/>'
          str << '<a:t>' << clean_value << '</a:t>'
          str << '</a:r>'
          str << '</a:p>'
          str << '</c:rich>'
        end
        str << '</c:tx>'
      end
      str << '<c:layout/>'
      str << '<c:overlay val="0"/>'
      str << '</c:title>'
    end
  end
end
