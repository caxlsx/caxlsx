# frozen_string_literal: true

module Axlsx
  # Conditional Format Rule icon sets
  # Describes an icon set conditional formatting rule.

  # @note The recommended way to manage these rules is via Worksheet#add_conditional_formatting
  # @see Worksheet#add_conditional_formatting
  # @see ConditionalFormattingRule#initialize
  class IconSet
    include Axlsx::OptionsParser
    include Axlsx::SerializedAttributes

    # Creates a new icon set object
    # @option options [String] iconSet
    # @option options [Boolean] reverse
    # @option options [Boolean] percent
    # @option options [Boolean] showValue
    def initialize(options = {})
      @percent = @showValue = true
      @reverse = false
      @iconSet = "3TrafficLights1"
      @interpolationPoints = [0, 33, 67]

      parse_options options

      yield self if block_given?
    end

    serializable_attributes :iconSet, :percent, :reverse, :showValue

    # The icon set to display.
    # Allowed values are: 3Arrows, 3ArrowsGray, 3Flags, 3TrafficLights1, 3TrafficLights2, 3Signs, 3Symbols, 3Symbols2, 4Arrows, 4ArrowsGray, 4RedToBlack, 4Rating, 4TrafficLights, 5Arrows, 5ArrowsGray, 5Rating, 5Quarters
    # The default value is 3TrafficLights1
    # @return [String]
    attr_reader :iconSet

    # Indicates whether the thresholds indicate percentile values, instead of number values.
    # The default value is true
    # @return [Boolean]
    attr_reader :percent

    # If true, reverses the default order of the icons in this icon set.maxLength attribute
    # The default value is false
    # @return [Boolean]
    attr_reader :reverse

    # Indicates whether to show the values of the cells on which this data bar is applied.
    # The default value is true
    # @return [Boolean]
    attr_reader :showValue

    # Sets the values of the interpolation points in the scale.
    # The default value is [0, 33, 67]
    # @return [Integer]
    attr_reader :interpolationPoints

    # @see iconSet
    def iconSet=(v)
      Axlsx.validate_icon_set(v)
      @iconSet = v
    end

    # @see interpolationPoints
    def interpolationPoints=(v)
      v.each { |point| Axlsx.validate_int(point) }
      @value_objects = nil
      @interpolationPoints = v
    end

    # @see showValue
    def showValue=(v)
      Axlsx.validate_boolean(v)
      @showValue = v
    end

    # @see percent
    def percent=(v)
      Axlsx.validate_boolean(v)
      @percent = v
    end

    # @see reverse
    def reverse=(v)
      Axlsx.validate_boolean(v)
      @reverse = v
    end

    # Serialize this object to an xml string
    # @param [String] str
    # @return [String]
    def to_xml_string(str = +'')
      initialize_value_objects if @value_objects.nil?

      serialized_tag('iconSet', str) do
        @value_objects.each { |cfvo| cfvo.to_xml_string(str) }
      end
    end

    private

    # Initialize the simple typed list of value objects
    def initialize_value_objects
      @value_objects = SimpleTypedList.new Cfvo
      @interpolationPoints.each { |point| @value_objects << Cfvo.new(type: :percent, val: point) }
      @value_objects.lock
    end
  end
end
