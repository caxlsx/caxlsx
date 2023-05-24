# frozen_string_literal: true

require 'tc_helper'

class TestSerAxis < Test::Unit::TestCase
  def setup
    @axis = Axlsx::SerAxis.new
  end

  def teardown; end

  def test_options
    a = Axlsx::SerAxis.new(tick_lbl_skip: 9, tick_mark_skip: 7)

    assert_equal(9, a.tick_lbl_skip)
    assert_equal(7, a.tick_mark_skip)
  end

  def test_tick_lbl_skip
    assert_raise(ArgumentError, "requires valid tick_lbl_skip") { @axis.tick_lbl_skip = -1 }
    assert_nothing_raised("accepts valid tick_lbl_skip") { @axis.tick_lbl_skip = 1 }
    assert_equal(1, @axis.tick_lbl_skip)
  end

  def test_tick_mark_skip
    assert_raise(ArgumentError, "requires valid tick_mark_skip") { @axis.tick_mark_skip = :my_eyes }
    assert_nothing_raised("accepts valid tick_mark_skip") { @axis.tick_mark_skip = 2 }
    assert_equal(2, @axis.tick_mark_skip)
  end
end
