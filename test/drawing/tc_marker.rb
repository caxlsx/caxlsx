# frozen_string_literal: true

require 'tc_helper'

class TestMarker < Minitest::Test
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
    assert_raises(ArgumentError) { @marker.col = -1 }
    refute_raises { @marker.col = 10 }
  end

  def test_colOff
    assert_raises(ArgumentError) { @marker.colOff = "1" }
    refute_raises { @marker.colOff = -10 }
  end

  def test_row
    assert_raises(ArgumentError) { @marker.row = -1 }
    refute_raises { @marker.row = 10 }
  end

  def test_rowOff
    assert_raises(ArgumentError) { @marker.rowOff = "1" }
    refute_raises { @marker.rowOff = -10 }
  end

  def test_coord
    @marker.coord 5, 10

    assert_equal(5, @marker.col)
    assert_equal(10, @marker.row)
  end
end
