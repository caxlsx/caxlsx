# frozen_string_literal: true

require 'tc_helper'

class TestBarChart < Minitest::Test
  def setup
    @p = Axlsx::Package.new
    ws = @p.workbook.add_worksheet
    @row = ws.add_row ["one", 1, Time.now]
    @chart = ws.add_chart Axlsx::BarChart, title: "fishery"
  end

  def teardown; end

  def test_initialization
    assert_equal(:clustered, @chart.grouping, "grouping default incorrect")
    assert_equal(@chart.series_type, Axlsx::BarSeries, "series type incorrect")
    assert_equal(:bar, @chart.bar_dir, " bar direction incorrect")
    assert_kind_of(Axlsx::CatAxis, @chart.cat_axis, "category axis not created")
    assert_kind_of(Axlsx::ValAxis, @chart.val_axis, "value access not created")
  end

  def test_bar_direction
    assert_raises(ArgumentError, "require valid bar direction") { @chart.bar_dir = :left }
    refute_raises { @chart.bar_dir = :col }
    assert_equal(:col, @chart.bar_dir)
  end

  def test_grouping
    assert_raises(ArgumentError, "require valid grouping") { @chart.grouping = :inverted }
    refute_raises { @chart.grouping = :standard }
    assert_equal(:standard, @chart.grouping)
  end

  def test_gap_width
    assert_raises(ArgumentError, "require valid gap width") { @chart.gap_width = -1 }
    assert_raises(ArgumentError, "require valid gap width") { @chart.gap_width = 501 }
    refute_raises { @chart.gap_width = 200 }
    assert_equal(200, @chart.gap_width, 'gap width is incorrect')
  end

  def test_overlap
    assert_raises(ArgumentError, "require valid overlap") { @chart.overlap = -101 }
    assert_raises(ArgumentError, "require valid overlap") { @chart.overlap = 101 }
    refute_raises { @chart.overlap = 100 }
    assert_equal(100, @chart.overlap, 'overlap is incorrect')
  end

  def test_shape
    assert_raises(ArgumentError, "require valid shape") { @chart.shape = :star }
    refute_raises { @chart.shape = :cone }
    assert_equal(:cone, @chart.shape)
  end

  def test_to_xml_string
    schema = Nokogiri::XML::Schema(File.open(Axlsx::DRAWING_XSD))
    doc = Nokogiri::XML(@chart.to_xml_string)
    errors = schema.validate(doc)

    assert_empty(errors)
  end

  def test_to_xml_string_has_axes_in_correct_order
    str = @chart.to_xml_string
    cat_axis_position = str.index(@chart.axes[:cat_axis].id.to_s)
    val_axis_position = str.index(@chart.axes[:val_axis].id.to_s)

    assert_operator(cat_axis_position, :<, val_axis_position, "cat_axis must occur earlier than val_axis in the XML")
  end

  def test_to_xml_string_has_gap_width
    gap_width_value = rand(0..500)
    @chart.gap_width = gap_width_value
    doc = Nokogiri::XML(@chart.to_xml_string)

    assert_equal(doc.xpath("//c:barChart/c:gapWidth").first.attribute('val').value, gap_width_value.to_s)
  end

  def test_to_xml_string_has_overlap
    overlap_value = rand(-100..100)
    @chart.overlap = overlap_value
    doc = Nokogiri::XML(@chart.to_xml_string)

    assert_equal(doc.xpath("//c:barChart/c:overlap").first.attribute('val').value, overlap_value.to_s)
  end
end
