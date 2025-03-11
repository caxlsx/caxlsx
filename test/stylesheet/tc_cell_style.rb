# frozen_string_literal: true

require 'tc_helper'

class TestCellStyle < Minitest::Test
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
    assert_raises(ArgumentError) { @item.name = -1 }
    refute_raises { @item.name = "stylin" }
    assert_equal("stylin", @item.name)
  end

  def test_xfId
    assert_raises(ArgumentError) { @item.xfId = -1 }
    refute_raises { @item.xfId = 5 }
    assert_equal(5, @item.xfId)
  end

  def test_builtinId
    assert_raises(ArgumentError) { @item.builtinId = -1 }
    refute_raises { @item.builtinId = 5 }
    assert_equal(5, @item.builtinId)
  end

  def test_iLevel
    assert_raises(ArgumentError) { @item.iLevel = -1 }
    refute_raises { @item.iLevel = 5 }
    assert_equal(5, @item.iLevel)
  end

  def test_hidden
    assert_raises(ArgumentError) { @item.hidden = -1 }
    refute_raises { @item.hidden = true }
    assert(@item.hidden)
  end

  def test_customBuiltin
    assert_raises(ArgumentError) { @item.customBuiltin = -1 }
    refute_raises { @item.customBuiltin = true }
    assert(@item.customBuiltin)
  end
end
