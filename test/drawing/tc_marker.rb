require 'tc_helper'

class TestMarker < Test::Unit::TestCase
  def setup
    @marker = Axlsx::Marker.new
  end

  def teardown; end

  def test_initialization
    assert_equal(0, @marker.col)
    assert_equal(0, @marker.colOff)
    assert_equal(0, @marker.row)
    assert_equal(0, @marker.rowOff)
  end

  def test_col
    assert_raise(ArgumentError) { @marker.col = -1 }
    assert_nothing_raised { @marker.col = 10 }
  end

  def test_colOff
    assert_raise(ArgumentError) { @marker.colOff = "1" }
    assert_nothing_raised { @marker.colOff = -10 }
  end

  def test_row
    assert_raise(ArgumentError) { @marker.row = -1 }
    assert_nothing_raised { @marker.row = 10 }
  end

  def test_rowOff
    assert_raise(ArgumentError) { @marker.rowOff = "1" }
    assert_nothing_raised { @marker.rowOff = -10 }
  end

  def test_coord
    @marker.coord 5, 10

    assert_equal(5, @marker.col)
    assert_equal(10, @marker.row)
  end
end
