# frozen_string_literal: true

require 'tc_helper'

class TestValAxis < Minitest::Test
  def setup
    @axis = Axlsx::ValAxis.new
  end

  def teardown; end

  def test_initialization
    assert_equal(:between, @axis.cross_between, "axis crossBetween default incorrect")
  end

  def test_options
    a = Axlsx::ValAxis.new(cross_between: :midCat)

    assert_equal(:midCat, a.cross_between)
  end

  def test_crossBetween
    assert_raises(ArgumentError, "requires valid crossBetween") { @axis.cross_between = :my_eyes }
    refute_raises { @axis.cross_between = :midCat }
  end
end
