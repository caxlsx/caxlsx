# frozen_string_literal: true

require 'tc_helper'

class TestScatterChart < Minitest::Test
  def setup
    @p = Axlsx::Package.new
    @chart = nil
    @p.workbook.add_worksheet do |sheet|
      sheet.add_row ["First",  1,  5, 7, 9]
      sheet.add_row ["",       1, 25, 49, 81]
      sheet.add_row ["Second", 5, 2, 14, 9]
      sheet.add_row ["",       5, 10, 15, 20]
      sheet.add_chart(Axlsx::ScatterChart, title: "example 7: Scatter Chart") do |chart|
        chart.start_at 0, 4
        chart.end_at 10, 19
        chart.add_series xData: sheet["B1:E1"], yData: sheet["B2:E2"], title: sheet["A1"]
        chart.add_series xData: sheet["B3:E3"], yData: sheet["B4:E4"], title: sheet["A3"]
        @chart = chart
      end
    end
  end

  def teardown; end

  def test_scatter_style
    @chart.scatterStyle = :marker

    assert_equal(:marker, @chart.scatterStyle)
    assert_raises(ArgumentError) { @chart.scatterStyle = :buckshot }
  end

  def test_initialization
    assert_equal(:lineMarker, @chart.scatterStyle, "scatterStyle default incorrect")
    assert_equal(@chart.series_type, Axlsx::ScatterSeries, "series type incorrect")
    assert_kind_of(Axlsx::ValAxis, @chart.xValAxis, "independent value axis not created")
    assert_kind_of(Axlsx::ValAxis, @chart.yValAxis, "dependent value axis not created")
  end

  def test_to_xml_string
    schema = Nokogiri::XML::Schema(File.open(Axlsx::DRAWING_XSD))
    doc = Nokogiri::XML(@chart.to_xml_string)
    errors = schema.validate(doc)

    assert_empty(errors)
  end
end
