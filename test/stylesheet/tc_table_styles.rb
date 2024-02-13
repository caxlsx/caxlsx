# frozen_string_literal: true

require 'tc_helper'

class TestTableStyles < Minitest::Test
  def setup
    @item = Axlsx::TableStyles.new
  end

  def teardown; end

  def test_initialiation
    assert_equal("TableStyleMedium9", @item.defaultTableStyle)
    assert_equal("PivotStyleLight16", @item.defaultPivotStyle)
  end

  def test_defaultTableStyle
    assert_raises(ArgumentError) { @item.defaultTableStyle = -1.1 }
    refute_raises { @item.defaultTableStyle = "anyones guess" }
    assert_equal("anyones guess", @item.defaultTableStyle)
  end

  def test_defaultPivotStyle
    assert_raises(ArgumentError) { @item.defaultPivotStyle = -1.1 }
    refute_raises { @item.defaultPivotStyle = "anyones guess" }
    assert_equal("anyones guess", @item.defaultPivotStyle)
  end
end
