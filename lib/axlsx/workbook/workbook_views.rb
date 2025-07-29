# frozen_string_literal: true

module Axlsx
  # a simple types list of BookView objects
  class WorkbookViews < SimpleTypedList
    # creates the book views object
    def initialize
      super(WorkbookView)
    end

    # Serialize to xml
    # @param [#<<] str A String, buffer or IO to append the serialization to.
    # @return [void]
    def to_xml_string(str = +'')
      return if empty?

      str << "<bookViews>"
      each { |view| view.to_xml_string(str) }
      str << '</bookViews>'
    end
  end
end
