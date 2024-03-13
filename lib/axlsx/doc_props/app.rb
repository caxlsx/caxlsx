# frozen_string_literal: true

module Axlsx
  # App represents the app.xml document. The attributes for this object are primarily managed by the application the end user uses to edit the document. None of the attributes are required to serialize a valid xlsx object.
  # @see shared-documentPropertiesExtended.xsd
  # @note Support is not implemented for the following complex types:
  #
  #    HeadingPairs (VectorVariant),
  #    TitlesOfParts (VectorLpstr),
  #    HLinks (VectorVariant),
  #    DigSig (DigSigBlob)
  class App
    include Axlsx::Accessors
    include Axlsx::OptionsParser

    # Creates an App object
    # @option options [String] template
    # @option options [String] manager
    # @option options [Integer] pages
    # @option options [Integer] words
    # @option options [Integer] characters
    # @option options [String] presentation_format
    # @option options [Integer] lines
    # @option options [Integer] paragraphs
    # @option options [Integer] slides
    # @option options [Integer] notes
    # @option options [Integer] total_time
    # @option options [Integer] hidden_slides
    # @option options [Integer] m_m_clips
    # @option options [Boolean] scale_crop
    # @option options [Boolean] links_up_to_date
    # @option options [Integer] characters_with_spaces
    # @option options [Boolean] share_doc
    # @option options [String] hyperlink_base
    # @option options [String] hyperlinks_changed
    # @option options [String] application
    # @option options [String] app_version
    # @option options [Integer] doc_security
    def initialize(options = {})
      parse_options options
    end

    # @return [String] The name of the document template.
    # @!attribute
    string_attr_accessor :template
    alias :Template :template
    alias :Template= :template=

    # @return [String] The name of the manager for the document.
    # @!attribute
    string_attr_accessor :manager
    alias :Manager :manager
    alias :Manager= :manager=

    # @return [String] The name of the company generating the document.
    # @!attribute
    string_attr_accessor :company
    alias :Company :company
    alias :Company= :company=

    # @return [Integer] The number of pages in the document.
    # @!attribute
    int_attr_accessor :pages
    alias :Pages :pages
    alias :Pages= :pages=

    # @return [Integer] The number of words in the document.
    # @!attribute
    int_attr_accessor :words
    alias :Words :words
    alias :Words= :words=

    # @return [Integer] The number of characters in the document.
    # @!attribute
    int_attr_accessor :characters
    alias :Characters :characters
    alias :Characters= :characters=

    # @return [String] The intended format of the presentation.
    # @!attribute
    string_attr_accessor :presentation_format
    alias :PresentationFormat :presentation_format
    alias :PresentationFormat= :presentation_format=

    # @return [Integer] The number of lines in the document.
    # @!attribute
    int_attr_accessor :lines
    alias :Lines :lines
    alias :Lines= :lines=

    # @return [Integer] The number of paragraphs in the document
    # @!attribute
    int_attr_accessor :paragraphs
    alias :Paragraphs :paragraphs
    alias :Paragraphs= :paragraphs=

    # @return [Intger] The number of slides in the document.
    # @!attribute
    int_attr_accessor :slides
    alias :Slides :slides
    alias :Slides= :slides=

    # @return [Integer] The number of slides that have notes.
    # @!attribute
    int_attr_accessor :notes
    alias :Notes :notes
    alias :Notes= :notes=

    # @return [Integer] The total amount of time spent editing.
    # @!attribute
    int_attr_accessor :total_time
    alias :TotalTime :total_time
    alias :TotalTime= :total_time=

    # @return [Integer] The number of hidden slides.
    # @!attribute
    int_attr_accessor :hidden_slides
    alias :HiddenSlides :hidden_slides
    alias :HiddenSlides= :hidden_slides=

    # @return [Integer] The total number multimedia clips
    # @!attribute
    int_attr_accessor :m_m_clips
    alias :MMClips :m_m_clips
    alias :MMClips= :m_m_clips=

    # @return [Boolean] The display mode for the document thumbnail.
    # @!attribute
    boolean_attr_accessor :scale_crop
    alias :ScaleCrop :scale_crop
    alias :ScaleCrop= :scale_crop=

    # @return [Boolean] The links in the document are up to date.
    # @!attribute
    boolean_attr_accessor :links_up_to_date
    alias :LinksUpToDate :links_up_to_date
    alias :LinksUpToDate= :links_up_to_date=

    # @return [Integer] The number of characters in the document including spaces.
    # @!attribute
    int_attr_accessor :characters_with_spaces
    alias :CharactersWithSpaces :characters_with_spaces
    alias :CharactersWithSpaces= :characters_with_spaces=

    # @return [Boolean] Indicates if the document is shared.
    # @!attribute
    boolean_attr_accessor :shared_doc
    alias :SharedDoc :shared_doc
    alias :SharedDoc= :shared_doc=

    # @return [String] The base for hyper links in the document.
    # @!attribute
    string_attr_accessor :hyperlink_base
    alias :HyperlinkBase :hyperlink_base
    alias :HyperlinkBase= :hyperlink_base=

    # @return [Boolean] Indicates that the hyper links in the document have been changed.
    # @!attribute
    boolean_attr_accessor :hyperlinks_changed
    alias :HyperlinksChanged :hyperlinks_changed
    alias :HyperLinksChanged= :hyperlinks_changed=

    # @return [String] The name of the application
    # @!attribute
    string_attr_accessor :application
    alias :Application :application
    alias :Application= :application=

    # @return [String] The version of the application.
    # @!attribute
    string_attr_accessor :app_version
    alias :AppVersion :app_version
    alias :AppVersion= :app_version=

    # @return [Integer] Document security
    # @!attribute
    int_attr_accessor :doc_security
    alias :DocSecurity :doc_security
    alias :DocSecurity= :doc_security=

    # Serialize the app.xml document
    # @return [String]
    def to_xml_string(str = +'')
      str << '<?xml version="1.0" encoding="UTF-8"?>'
      str << '<Properties xmlns="' << APP_NS << '" xmlns:vt="' << APP_NS_VT << '">'
      Axlsx.instance_values_for(self).each do |key, value|
        node_name = Axlsx.camel(key)
        str << "<#{node_name}>#{value}</#{node_name}>"
      end
      str << '</Properties>'
    end
  end
end
