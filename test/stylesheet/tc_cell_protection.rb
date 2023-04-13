require 'tc_helper.rb'

class TestCellProtection < Test::Unit::TestCase
  def setup
    @item = Axlsx::CellProtection.new
  end

  def teardown; end

  def test_initialiation
    assert_nil(@item.hidden)
    assert_nil(@item.locked)
  end

  def test_hidden
    assert_raise(ArgumentError) { @item.hidden = -1 }
    assert_nothing_raised { @item.hidden = false }
    refute(@item.hidden)
  end

  def test_locked
    assert_raise(ArgumentError) { @item.locked = -1 }
    assert_nothing_raised { @item.locked = false }
    refute(@item.locked)
  end
end
