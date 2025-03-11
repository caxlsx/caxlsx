# frozen_string_literal: true

require 'tc_helper'

class TestAxes < Minitest::Test
  def test_constructor_requires_cat_axis_first
    assert_raises(ArgumentError) { Axlsx::Axes.new(val_axis: Axlsx::ValAxis, cat_axis: Axlsx::CatAxis) }
    refute_raises { Axlsx::Axes.new(cat_axis: Axlsx::CatAxis, val_axis: Axlsx::ValAxis) }
  end
end
