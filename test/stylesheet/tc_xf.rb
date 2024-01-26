# frozen_string_literal: true

require 'tc_helper'

class TestXf < Test::Unit::TestCase
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
    assert_raise(ArgumentError) { @item.alignment = -1.1 }
    assert_nothing_raised { @item.alignment = Axlsx::CellAlignment.new }
    assert_kind_of(Axlsx::CellAlignment, @item.alignment)
  end

  def test_protection
    assert_raise(ArgumentError) { @item.protection = -1.1 }
    assert_nothing_raised { @item.protection = Axlsx::CellProtection.new }
    assert_kind_of(Axlsx::CellProtection, @item.protection)
  end

  def test_numFmtId
    assert_raise(ArgumentError) { @item.numFmtId = -1.1 }
    assert_nothing_raised { @item.numFmtId = 0 }
    assert_equal(0, @item.numFmtId)
  end

  def test_fillId
    assert_raise(ArgumentError) { @item.fillId = -1.1 }
    assert_nothing_raised { @item.fillId = 0 }
    assert_equal(0, @item.fillId)
  end

  def test_fontId
    assert_raise(ArgumentError) { @item.fontId = -1.1 }
    assert_nothing_raised { @item.fontId = 0 }
    assert_equal(0, @item.fontId)
  end

  def test_borderId
    assert_raise(ArgumentError) { @item.borderId = -1.1 }
    assert_nothing_raised { @item.borderId = 0 }
    assert_equal(0, @item.borderId)
  end

  def test_xfId
    assert_raise(ArgumentError) { @item.xfId = -1.1 }
    assert_nothing_raised { @item.xfId = 0 }
    assert_equal(0, @item.xfId)
  end

  def test_quotePrefix
    assert_raise(ArgumentError) { @item.quotePrefix = -1.1 }
    assert_nothing_raised { @item.quotePrefix = false }
    refute(@item.quotePrefix)
  end

  def test_pivotButton
    assert_raise(ArgumentError) { @item.pivotButton = -1.1 }
    assert_nothing_raised { @item.pivotButton = false }
    refute(@item.pivotButton)
  end

  def test_applyNumberFormat
    assert_raise(ArgumentError) { @item.applyNumberFormat = -1.1 }
    assert_nothing_raised { @item.applyNumberFormat = false }
    refute(@item.applyNumberFormat)
  end

  def test_applyFont
    assert_raise(ArgumentError) { @item.applyFont = -1.1 }
    assert_nothing_raised { @item.applyFont = false }
    refute(@item.applyFont)
  end

  def test_applyFill
    assert_raise(ArgumentError) { @item.applyFill = -1.1 }
    assert_nothing_raised { @item.applyFill = false }
    refute(@item.applyFill)
  end

  def test_applyBorder
    assert_raise(ArgumentError) { @item.applyBorder = -1.1 }
    assert_nothing_raised { @item.applyBorder = false }
    refute(@item.applyBorder)
  end

  def test_applyAlignment
    assert_raise(ArgumentError) { @item.applyAlignment = -1.1 }
    assert_nothing_raised { @item.applyAlignment = false }
    refute(@item.applyAlignment)
  end

  def test_applyProtection
    assert_raise(ArgumentError) { @item.applyProtection = -1.1 }
    assert_nothing_raised { @item.applyProtection = false }
    refute(@item.applyProtection)
  end
end
