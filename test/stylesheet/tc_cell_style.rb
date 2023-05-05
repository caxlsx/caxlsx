# frozen_string_literal: true

require 'tc_helper'

class TestCellStyle < Test::Unit::TestCase
  def setup
    @item = Axlsx::CellStyle.new
  end

  def teardown; end

  def test_initialiation
    assert_nil(@item.name)
    assert_nil(@item.xfId)
    assert_nil(@item.builtinId)
    assert_nil(@item.iLevel)
    assert_nil(@item.hidden)
    assert_nil(@item.customBuiltin)
  end

  def test_name
    assert_raise(ArgumentError) { @item.name = -1 }
    assert_nothing_raised { @item.name = "stylin" }
    assert_equal("stylin", @item.name)
  end

  def test_xfId
    assert_raise(ArgumentError) { @item.xfId = -1 }
    assert_nothing_raised { @item.xfId = 5 }
    assert_equal(5, @item.xfId)
  end

  def test_builtinId
    assert_raise(ArgumentError) { @item.builtinId = -1 }
    assert_nothing_raised { @item.builtinId = 5 }
    assert_equal(5, @item.builtinId)
  end

  def test_iLevel
    assert_raise(ArgumentError) { @item.iLevel = -1 }
    assert_nothing_raised { @item.iLevel = 5 }
    assert_equal(5, @item.iLevel)
  end

  def test_hidden
    assert_raise(ArgumentError) { @item.hidden = -1 }
    assert_nothing_raised { @item.hidden = true }
    assert(@item.hidden)
  end

  def test_customBuiltin
    assert_raise(ArgumentError) { @item.customBuiltin = -1 }
    assert_nothing_raised { @item.customBuiltin = true }
    assert(@item.customBuiltin)
  end
end
