# encoding: UTF-8
module Axlsx
  # A series title is a Title with a slightly different serialization than chart titles.
  class SeriesTitle < Title

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      clean_value = Axlsx::trust_input ? @text.to_s : ::CGI.escapeHTML(Axlsx::sanitize(@text.to_s))

      str << '<c:tx>'
      str << '<c:strRef>'
      str << ('<c:f>' << Axlsx::cell_range([@cell]) << '</c:f>')
      str << '<c:strCache>'
      str << '<c:ptCount val="1"/>'
      str << '<c:pt idx="0">'
      str << ('<c:v>' << clean_value << '</c:v>')
      str << '</c:pt>'
      str << '</c:strCache>'
      str << '</c:strRef>'
      str << '</c:tx>'
    end
  end
end
