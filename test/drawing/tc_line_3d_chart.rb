# frozen_string_literal: true

require 'tc_helper'

class TestLine3DChart < Test::Unit::TestCase
  def setup
    @p = Axlsx::Package.new
    ws = @p.workbook.add_worksheet
    @row = ws.add_row ["one", 1, Time.now]
    @chart = ws.add_chart Axlsx::Line3DChart, title: "fishery"
  end

  def teardown; end

  def test_initialization
    assert_equal(:standard, @chart.grouping, "grouping defualt incorrect")
    assert_equal(@chart.series_type, Axlsx::LineSeries, "series type incorrect")
    assert_kind_of(Axlsx::CatAxis, @chart.catAxis, "category axis not created")
    assert_kind_of(Axlsx::ValAxis, @chart.valAxis, "value access not created")
    assert_kind_of(Axlsx::SerAxis, @chart.serAxis, "value access not created")
  end

  def test_grouping
    assert_raise(ArgumentError, "require valid grouping") { @chart.grouping = :inverted }
    assert_nothing_raised("allow valid grouping") { @chart.grouping = :stacked }
    assert_equal(:stacked, @chart.grouping)
  end

  def test_gapDepth
    assert_raise(ArgumentError, "require valid gapDepth") { @chart.gapDepth = 200 }
    assert_nothing_raised("allow valid gapDepth") { @chart.gapDepth = "200%" }
    assert_equal("200%", @chart.gapDepth)
  end

  def test_to_xml
    schema = Nokogiri::XML::Schema(File.open(Axlsx::DRAWING_XSD))
    doc = Nokogiri::XML(@chart.to_xml_string)
    errors = schema.validate(doc)

    assert_empty(errors)
  end
end
