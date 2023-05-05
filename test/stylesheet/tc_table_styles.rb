# frozen_string_literal: true

require 'tc_helper'

class TestTableStyles < Test::Unit::TestCase
  def setup
    @item = Axlsx::TableStyles.new
  end

  def teardown; end

  def test_initialiation
    assert_equal("TableStyleMedium9", @item.defaultTableStyle)
    assert_equal("PivotStyleLight16", @item.defaultPivotStyle)
  end

  def test_defaultTableStyle
    assert_raise(ArgumentError) { @item.defaultTableStyle = -1.1 }
    assert_nothing_raised { @item.defaultTableStyle = "anyones guess" }
    assert_equal("anyones guess", @item.defaultTableStyle)
  end

  def test_defaultPivotStyle
    assert_raise(ArgumentError) { @item.defaultPivotStyle = -1.1 }
    assert_nothing_raised { @item.defaultPivotStyle = "anyones guess" }
    assert_equal("anyones guess", @item.defaultPivotStyle)
  end
end
