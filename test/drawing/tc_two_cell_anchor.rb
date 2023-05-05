# frozen_string_literal: true

require 'tc_helper'

class TestTwoCellAnchor < Test::Unit::TestCase
  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet
    @ws.add_row ["one", 1, Time.now]
    chart = @ws.add_chart Axlsx::Bar3DChart
    @anchor = chart.graphic_frame.anchor
  end

  def test_initialization
    assert_equal(0, @anchor.from.col)
    assert_equal(0, @anchor.from.row)
    assert_equal(5, @anchor.to.col)
    assert_equal(10, @anchor.to.row)
  end

  def test_index
    assert_equal(@anchor.index, @anchor.drawing.anchors.index(@anchor))
  end

  def test_options
    assert_raise(ArgumentError, 'invalid start_at') { @ws.add_chart Axlsx::Chart, :start_at => "1" }
    assert_raise(ArgumentError, 'invalid end_at') { @ws.add_chart Axlsx::Chart, :start_at => [1, 2], :end_at => ["a", 4] }
    # this is actually raised in the graphic frame
    assert_raise(ArgumentError, 'invalid Chart') { @ws.add_chart Axlsx::TwoCellAnchor }
    a = @ws.add_chart Axlsx::Chart, :start_at => [15, 35], :end_at => [90, 45]

    assert_equal(15, a.graphic_frame.anchor.from.col)
    assert_equal(35, a.graphic_frame.anchor.from.row)
    assert_equal(90, a.graphic_frame.anchor.to.col)
    assert_equal(45, a.graphic_frame.anchor.to.row)
  end
end
