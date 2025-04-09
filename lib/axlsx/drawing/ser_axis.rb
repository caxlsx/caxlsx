# frozen_string_literal: true

module Axlsx
  # A SerAxis object defines a series axis
  class SerAxis < Axis
    # The number of tick labels to skip between labels
    # @return [Integer]
    attr_reader :tick_lbl_skip
    alias :tickLblSkip :tick_lbl_skip

    # The number of tickmarks to be skipped before the next one is rendered.
    # @return [Boolean]
    attr_reader :tick_mark_skip
    alias :tickMarkSkip :tick_mark_skip

    # Creates a new SerAxis object
    # @option options [Integer] tick_lbl_skip
    # @option options [Integer] tick_mark_skip
    def initialize(options = {})
      @tick_lbl_skip, @tick_mark_skip = 1, 1
      super
    end

    # @see tickLblSkip
    def tick_lbl_skip=(v)
      Axlsx.validate_unsigned_int(v)
      @tick_lbl_skip = v
    end
    alias :tickLblSkip= :tick_lbl_skip=

    # @see tickMarkSkip
    def tick_mark_skip=(v)
      Axlsx.validate_unsigned_int(v)
      @tick_mark_skip = v
    end
    alias :tickMarkSkip= :tick_mark_skip=

    # Serializes the object
    # @param [#<<] str A String, buffer or IO to append the serialization to.
    # @return [void]
    def to_xml_string(str = +'')
      str << '<c:serAx>'
      super
      str << '<c:tickLblSkip val="' << @tick_lbl_skip.to_s << '"/>' unless @tick_lbl_skip.nil?
      str << '<c:tickMarkSkip val="' << @tick_mark_skip.to_s << '"/>' unless @tick_mark_skip.nil?
      str << '</c:serAx>'
    end
  end
end
