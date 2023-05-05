# frozen_string_literal: true

require 'tc_helper'

class TestBorderPr < Test::Unit::TestCase
  def setup
    @bpr = Axlsx::BorderPr.new
  end

  def teardown; end

  def test_initialiation
    assert_nil(@bpr.color)
    assert_nil(@bpr.style)
    assert_nil(@bpr.name)
  end

  def test_color
    assert_raise(ArgumentError) { @bpr.color = :red }
    assert_nothing_raised { @bpr.color = Axlsx::Color.new :rgb => "FF000000" }
    assert(@bpr.color.is_a?(Axlsx::Color))
  end

  def test_style
    assert_raise(ArgumentError) { @bpr.style = :red }
    assert_nothing_raised { @bpr.style = :thin }
    assert_equal(:thin, @bpr.style)
  end

  def test_name
    assert_raise(ArgumentError) { @bpr.name = :red }
    assert_nothing_raised { @bpr.name = :top }
    assert_equal(:top, @bpr.name)
  end
end
