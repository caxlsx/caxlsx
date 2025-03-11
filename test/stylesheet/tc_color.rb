# frozen_string_literal: true

require 'tc_helper'

class TestColor < Minitest::Test
  def setup
    @item = Axlsx::Color.new
  end

  def teardown; end

  def test_initialiation
    assert_nil(@item.auto)
    assert_equal("FF000000", @item.rgb)
    assert_nil(@item.tint)
  end

  def test_auto
    assert_raises(ArgumentError) { @item.auto = -1 }
    refute_raises { @item.auto = true }
    assert(@item.auto)
  end

  def test_rgb
    assert_raises(ArgumentError) { @item.rgb = -1 }
    refute_raises { @item.rgb = "FF00FF00" }
    assert_equal("FF00FF00", @item.rgb)
  end

  def test_rgb_writer_doesnt_mutate_its_argument
    my_rgb = 'ff00ff00'
    @item.rgb = my_rgb

    assert_equal 'ff00ff00', my_rgb
  end

  def test_tint
    assert_raises(ArgumentError) { @item.tint = -1 }
    refute_raises { @item.tint = -1.0 }
    assert_in_delta(@item.tint, -1.0)
  end
end
