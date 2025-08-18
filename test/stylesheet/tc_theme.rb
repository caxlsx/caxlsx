# frozen_string_literal: true

require 'tc_helper'

class TestTheme < Minitest::Test
  def setup
    @theme = Axlsx::Theme.new
  end

  def test_pn
    assert_equal(Axlsx::THEME_PN, @theme.pn)
  end

  def test_to_xml_string_returns_valid_xml
    xml = @theme.to_xml_string

    # Basic structure checks
    assert_includes(xml, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    assert_includes(xml, '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">')
    assert_includes(xml, '</a:theme>')

    # Required sections
    assert_includes(xml, '<a:themeElements>')
    assert_includes(xml, '<a:clrScheme name="Office">')
    assert_includes(xml, '<a:fontScheme name="Office">')
    assert_includes(xml, '<a:fmtScheme name="Office">')
    assert_includes(xml, '<a:objectDefaults>')
    assert_includes(xml, '<a:extraClrSchemeLst/>')
  end

  def test_to_xml_string_with_string_parameter
    str = ''
    result = @theme.to_xml_string(str)

    # Should return the same string object that was passed in
    assert_same(str, result)
    refute_empty(str)
    refute_includes(str, '<a:theme')
  end

  def test_color_scheme_elements
    xml = @theme.to_xml_string

    # Check for required color scheme elements
    assert_includes(xml, '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>')
    assert_includes(xml, '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>')
    assert_includes(xml, '<a:dk2><a:srgbClr val="1F497D"/></a:dk2>')
    assert_includes(xml, '<a:lt2><a:srgbClr val="EEECE1"/></a:lt2>')

    # Check accent colors
    (1..6).each do |i|
      assert_includes(xml, "<a:accent#{i}>")
    end

    # Check hyperlink colors
    assert_includes(xml, '<a:hlink><a:srgbClr val="0000FF"/></a:hlink>')
    assert_includes(xml, '<a:folHlink><a:srgbClr val="800080"/></a:folHlink>')
  end

  def test_font_scheme_elements
    xml = @theme.to_xml_string

    # Check for major and minor fonts
    assert_includes(xml, '<a:majorFont>')
    assert_includes(xml, '<a:latin typeface="Cambria"/>')
    assert_includes(xml, '<a:minorFont>')
    assert_includes(xml, '<a:latin typeface="Calibri"/>')
  end

  def test_format_scheme_elements
    xml = @theme.to_xml_string

    # Check for format scheme sections
    assert_includes(xml, '<a:fillStyleLst>')
    assert_includes(xml, '<a:lnStyleLst>')
    assert_includes(xml, '<a:effectStyleLst>')
    assert_includes(xml, '<a:bgFillStyleLst>')
  end

  def test_object_defaults
    xml = @theme.to_xml_string

    # Check for object defaults
    assert_includes(xml, '<a:spDef>')
    assert_includes(xml, '<a:lnDef>')
    assert_includes(xml, '<a:spPr/>')
    assert_includes(xml, '<a:bodyPr/>')
    assert_includes(xml, '<a:lstStyle/>')
  end

  def test_3d_elements_present
    xml = @theme.to_xml_string

    # Check for 3D elements that are crucial for Excel compatibility
    assert_includes(xml, '<a:scene3d>')
    assert_includes(xml, '<a:camera prst="orthographicFront">')
    assert_includes(xml, '<a:lightRig rig="threePt" dir="t">')
    assert_includes(xml, '<a:sp3d>')
    assert_includes(xml, '<a:bevelT w="63500" h="25400"/>')
  end

  def test_xml_is_single_line_with_no_whitespace_padding
    xml = @theme.to_xml_string

    # XML should not contain extra whitespace or newlines
    refute_includes(xml, "\n")
    refute_includes(xml, "  ") # No double spaces
  end
end
