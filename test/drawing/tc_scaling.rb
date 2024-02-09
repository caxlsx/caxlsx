# frozen_string_literal: true

require 'tc_helper'

class TestScaling < Minitest::Test
  def setup
    @scaling = Axlsx::Scaling.new
  end

  def teardown; end

  def test_initialization
    assert_equal(:minMax, @scaling.orientation)
  end

  def test_logBase
    assert_raises(ArgumentError) { @scaling.logBase = 1 }
    refute_raises { @scaling.logBase = 10 }
  end

  def test_orientation
    assert_raises(ArgumentError) { @scaling.orientation = "1" }
    refute_raises { @scaling.orientation = :maxMin }
  end

  def test_max
    assert_raises(ArgumentError) { @scaling.max = 1 }
    refute_raises { @scaling.max = 10.5 }
  end

  def test_min
    assert_raises(ArgumentError) { @scaling.min = 1 }
    refute_raises { @scaling.min = 10.5 }
  end
end
