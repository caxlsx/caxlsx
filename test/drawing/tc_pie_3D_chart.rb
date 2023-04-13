require 'tc_helper'

class TestPie3DChart < Test::Unit::TestCase
  def setup
    p = Axlsx::Package.new
    ws = p.workbook.add_worksheet
    @row = ws.add_row ["one", 1, Time.now]
    @chart = ws.add_chart Axlsx::Pie3DChart, :title => "fishery"
  end

  def teardown; end

  def test_initialization
    assert_equal(30, @chart.view_3D.rot_x, "view 3d default rot_x incorrect")
    assert_equal(30, @chart.view_3D.perspective, "view_3d default perspective incorrect")
    assert_equal(@chart.series_type, Axlsx::PieSeries, "series type incorrect")
  end

  def test_to_xml
    schema = Nokogiri::XML::Schema(File.open(Axlsx::DRAWING_XSD))
    doc = Nokogiri::XML(@chart.to_xml_string)
    errors = schema.validate(doc).map { |error| puts error.message; error }

    assert_empty(errors, "error free validation")
  end
end
