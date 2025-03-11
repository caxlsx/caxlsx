# frozen_string_literal: true

require 'tc_helper'

class TestBorderPr < Minitest::Test
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
    assert_raises(ArgumentError) { @bpr.color = :red }
    refute_raises { @bpr.color = Axlsx::Color.new rgb: "FF000000" }
    assert_kind_of(Axlsx::Color, @bpr.color)
  end

  def test_style
    assert_raises(ArgumentError) { @bpr.style = :red }
    refute_raises { @bpr.style = :thin }
    assert_equal(:thin, @bpr.style)
  end

  def test_name
    assert_raises(ArgumentError) { @bpr.name = :red }
    refute_raises { @bpr.name = :top }
    assert_equal(:top, @bpr.name)
  end
end
