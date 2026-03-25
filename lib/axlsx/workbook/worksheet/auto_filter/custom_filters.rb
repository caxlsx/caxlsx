module Axlsx
  class CustomFilters
    include Axlsx::OptionsParser
    include Axlsx::SerializedAttributes
    include Axlsx::Accessors

    # Creates a new CustomFilters object
    # @param [Hash] options Options used to set this objects attributes and
    #                       create custom_filter_items.
    # @option [Boolean] and Value used to determine whether one or both conditions need to apply in order to be shown.
    # @option [Array] custom_filter_items An array of values that will be used to create custom_filter objects.
    # @note The recommended way to interact with custom_filter objects is via AutoFilter#add_column
    # @example
    #   ws.auto_filter.add_column(
    #     0,
    #     :custom_filters,
    #     and: true,
    #     custom_filter_items: [
    #       { operator: :greaterThan, val: 5 },
    #       { operator: :lessThanOrEqual, val: 10 }
    #     ]
    #   )
    def initialize(options = {})
      @and = false
      parse_options options
    end

    # A list of serializable attributes.
    serializable_attributes :and

    # Boolean attribute accessors
    boolean_attr_accessor :and

    # Alias needed for reader because `and` is a reserved keyword in Ruby
    alias_method :logical_and, :and

    # Tells us if the row of the cell provided should be hidden as it
    # does not meet any or all (based on the logical_and boolean) of
    # the specified custom_filter_items restrictions.
    # @param [Cell] cell The cell to test against items
    def apply(cell)
      return false unless cell

      if logical_and
        # If not all of the custom filters for this row evaluate as true then hide it.
        !custom_filter_items.all? { |custom_filter| custom_filter.apply(cell) }
      else
        # If none of the custom filter conditions for this row evaluate as true then hide it.
        custom_filter_items.none? { |custom_filter| custom_filter.apply(cell) }
      end
    end

    # Serialize the object to xml
    # @param [String] str The string to concat the serialization information to.
    def to_xml_string(str = '')
      str << "<customFilters #{serialized_attributes}>"
      custom_filter_items.each { |custom_filter| custom_filter.to_xml_string(str) }
      str << "</customFilters>"
    end

    def custom_filter_items
      @custom_filter_items ||= SimpleTypedList.new(CustomFilter)
    end

    def custom_filter_items=(values)
      values.each do |value|
        if value[:operator] == :between
          custom_filter_items << CustomFilter.new(:operator => :greaterThan, :val => value[:val][0])
          custom_filter_items << CustomFilter.new(:operator => :lessThan, :val => value[:val][1])
          @and = true
        else
          custom_filter_items << CustomFilter.new(value)
        end
      end
    end

    class CustomFilter
      include Axlsx::OptionsParser

      COMPARATOR_METHOD_MAP = {
        lessThan:           [:<, ->(v) { v }],
        lessThanOrEqual:    [:<=, ->(v) { v }],
        equal:              [:==, ->(v) { v }],
        notBlank:           [:!=, ->(v) { v }],
        notEqual:           [:!=, ->(v) { v }],
        greaterThanOrEqual: [:>=, ->(v) { v }],
        greaterThan:        [:>, ->(v) { v }],
        contains:           [:match?, ->(v) { /#{Regexp.escape(v)}/i }],
        notContains:        [:match?, ->(v) { /^(?!.*#{Regexp.escape(v)}).*$/i }],
        beginsWith:         [:match?, ->(v) { /\A#{Regexp.escape(v)}.*/i }],
        endsWith:           [:match?, ->(v) { /.*#{Regexp.escape(v)}\z/i }],
      }.freeze
      VALID_OPERATOR_MAP = {
        lessThan:           "lessThan",
        lessThanOrEqual:    "lessThanOrEqual",
        equal:              "equal",
        notBlank:           "notEqual",
        notEqual:           "notEqual",
        greaterThanOrEqual: "greaterThanOrEqual",
        greaterThan:        "greaterThan",
        contains:           "equal",
        notContains:        "notEqual",
        beginsWith:         "equal",
        endsWith:           "equal",
      }.freeze

      def initialize(options = {})
        raise ArgumentError, "You must specify an operator for the custom filter" unless options[:operator]
        parse_options options
      end

      attr_reader :operator
      attr_accessor :val

      def operator=(operation_type)
        RestrictionValidator.validate "CustomFilter.operator", VALID_OPERATOR_MAP.keys, operation_type
        RestrictionValidator.validate "CustomFilter.comparator_method", COMPARATOR_METHOD_MAP.keys, operation_type
        @operator = operation_type
      end

      def apply(cell)
        return true unless cell

        method, formatter = COMPARATOR_METHOD_MAP[operator]
        cell_val = cell.value
        return true if cell_val.nil?
        cell_val = cell_val.to_s if method == :match?
        return true if cell_val.send(method, formatter.call(val))

        false
      end

      # Serializes the custom_filter object
      # @param [String] str The string to concat the serialization information to.
      def to_xml_string(str = '')
        str << "<customFilter operator=\"#{VALID_OPERATOR_MAP[operator]}\" val=\"#{leading_wildcard + safe_val + trailing_wildcard}\" />"
      end

      private

      def leading_wildcard
        [:contains, :notContains, :endsWith].include?(operator) ? "*" : ""
      end

      def safe_val
        if operator == :notBlank && val.nil?
          " "
        else
          val.to_s
        end
      end

      def trailing_wildcard
        [:contains, :notContains, :beginsWith].include?(operator) ? "*" : ""
      end
    end
  end
end
