# frozen_string_literal: true

require 'tc_helper'

class TestCellProtection < Minitest::Test
  def setup
    @item = Axlsx::CellProtection.new
  end

  def teardown; end

  def test_initialiation
    assert_nil(@item.hidden)
    assert_nil(@item.locked)
  end

  def test_hidden
    assert_raises(ArgumentError) { @item.hidden = -1 }
    refute_raises { @item.hidden = false }
    assert_false(@item.hidden)
  end

  def test_locked
    assert_raises(ArgumentError) { @item.locked = -1 }
    refute_raises { @item.locked = false }
    assert_false(@item.locked)
  end
end
