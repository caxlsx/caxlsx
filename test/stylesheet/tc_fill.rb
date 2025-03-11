# frozen_string_literal: true

require 'tc_helper'

class TestFill < Minitest::Test
  def setup
    @item = Axlsx::Fill.new Axlsx::PatternFill.new
  end

  def teardown; end

  def test_initialiation
    assert_kind_of(Axlsx::PatternFill, @item.fill_type)
    assert_raises(ArgumentError) { Axlsx::Fill.new }
    refute_raises { Axlsx::Fill.new(Axlsx::GradientFill.new) }
  end
end
