# frozen_string_literal: true

require 'tc_helper'

class TestPieChart < Test::Unit::TestCase
  def setup
    p = Axlsx::Package.new
    ws = p.workbook.add_worksheet
    @row = ws.add_row ["one", 1, Time.now]
    @chart = ws.add_chart Axlsx::PieChart, title: "fishery"
  end

  def teardown; end

  def test_initialization
    assert_equal(@chart.series_type, Axlsx::PieSeries, "series type incorrect")
  end

  def test_to_xml
    schema = Nokogiri::XML::Schema(File.open(Axlsx::DRAWING_XSD))
    doc = Nokogiri::XML(@chart.to_xml_string)
    errors = schema.validate(doc)

    assert_empty(errors)
  end
end
