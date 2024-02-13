# frozen_string_literal: true

require 'tc_helper'

class TestXf < Minitest::Test
  def setup
    @item = Axlsx::Xf.new
  end

  def teardown; end

  def test_initialiation
    assert_nil(@item.alignment)
    assert_nil(@item.protection)
    assert_nil(@item.numFmtId)
    assert_nil(@item.fontId)
    assert_nil(@item.fillId)
    assert_nil(@item.borderId)
    assert_nil(@item.xfId)
    assert_nil(@item.quotePrefix)
    assert_nil(@item.pivotButton)
    assert_nil(@item.applyNumberFormat)
    assert_nil(@item.applyFont)
    assert_nil(@item.applyFill)
    assert_nil(@item.applyBorder)
    assert_nil(@item.applyAlignment)
    assert_nil(@item.applyProtection)
  end

  def test_alignment
    assert_raises(ArgumentError) { @item.alignment = -1.1 }
    refute_raises { @item.alignment = Axlsx::CellAlignment.new }
    assert_kind_of(Axlsx::CellAlignment, @item.alignment)
  end

  def test_protection
    assert_raises(ArgumentError) { @item.protection = -1.1 }
    refute_raises { @item.protection = Axlsx::CellProtection.new }
    assert_kind_of(Axlsx::CellProtection, @item.protection)
  end

  def test_numFmtId
    assert_raises(ArgumentError) { @item.numFmtId = -1.1 }
    refute_raises { @item.numFmtId = 0 }
    assert_equal(0, @item.numFmtId)
  end

  def test_fillId
    assert_raises(ArgumentError) { @item.fillId = -1.1 }
    refute_raises { @item.fillId = 0 }
    assert_equal(0, @item.fillId)
  end

  def test_fontId
    assert_raises(ArgumentError) { @item.fontId = -1.1 }
    refute_raises { @item.fontId = 0 }
    assert_equal(0, @item.fontId)
  end

  def test_borderId
    assert_raises(ArgumentError) { @item.borderId = -1.1 }
    refute_raises { @item.borderId = 0 }
    assert_equal(0, @item.borderId)
  end

  def test_xfId
    assert_raises(ArgumentError) { @item.xfId = -1.1 }
    refute_raises { @item.xfId = 0 }
    assert_equal(0, @item.xfId)
  end

  def test_quotePrefix
    assert_raises(ArgumentError) { @item.quotePrefix = -1.1 }
    refute_raises { @item.quotePrefix = false }
    assert_false(@item.quotePrefix)
  end

  def test_pivotButton
    assert_raises(ArgumentError) { @item.pivotButton = -1.1 }
    refute_raises { @item.pivotButton = false }
    assert_false(@item.pivotButton)
  end

  def test_applyNumberFormat
    assert_raises(ArgumentError) { @item.applyNumberFormat = -1.1 }
    refute_raises { @item.applyNumberFormat = false }
    assert_false(@item.applyNumberFormat)
  end

  def test_applyFont
    assert_raises(ArgumentError) { @item.applyFont = -1.1 }
    refute_raises { @item.applyFont = false }
    assert_false(@item.applyFont)
  end

  def test_applyFill
    assert_raises(ArgumentError) { @item.applyFill = -1.1 }
    refute_raises { @item.applyFill = false }
    assert_false(@item.applyFill)
  end

  def test_applyBorder
    assert_raises(ArgumentError) { @item.applyBorder = -1.1 }
    refute_raises { @item.applyBorder = false }
    assert_false(@item.applyBorder)
  end

  def test_applyAlignment
    assert_raises(ArgumentError) { @item.applyAlignment = -1.1 }
    refute_raises { @item.applyAlignment = false }
    assert_false(@item.applyAlignment)
  end

  def test_applyProtection
    assert_raises(ArgumentError) { @item.applyProtection = -1.1 }
    refute_raises { @item.applyProtection = false }
    assert_false(@item.applyProtection)
  end
end
