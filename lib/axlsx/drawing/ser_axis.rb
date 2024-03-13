# frozen_string_literal: true

module Axlsx
  # A SerAxis object defines a series axis
  class SerAxis < Axis
    include Axlsx::Accessors

    # The number of tick lables to skip between labels
    # @return [Integer]
    # @!attribute
    unsigned_int_attr_accessor :tick_lbl_skip
    alias :tickLblSkip :tick_lbl_skip
    alias :tickLblSkip= :tick_lbl_skip=

    # The number of tickmarks to be skipped before the next one is rendered.
    # @return [Integer]
    # @!attribute
    unsigned_int_attr_accessor :tick_mark_skip
    alias :tickMarkSkip :tick_mark_skip
    alias :tickMarkSkip= :tick_mark_skip=

    # Creates a new SerAxis object
    # @option options [Integer] tick_lbl_skip
    # @option options [Integer] tick_mark_skip
    def initialize(options = {})
      @tick_lbl_skip, @tick_mark_skip = 1, 1
      super(options)
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = +'')
      str << '<c:serAx>'
      super(str)
      str << '<c:tickLblSkip val="' << @tick_lbl_skip.to_s << '"/>' unless @tick_lbl_skip.nil?
      str << '<c:tickMarkSkip val="' << @tick_mark_skip.to_s << '"/>' unless @tick_mark_skip.nil?
      str << '</c:serAx>'
    end
  end
end
