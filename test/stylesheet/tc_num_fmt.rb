# frozen_string_literal: true

require 'tc_helper'

class TestNumFmt < Minitest::Test
  def setup
    @item = Axlsx::NumFmt.new
  end

  def teardown; end

  def test_initialiation
    assert_equal(0, @item.numFmtId)
    assert_equal("", @item.formatCode)
  end

  def test_numFmtId
    assert_raises(ArgumentError) { @item.numFmtId = -1.1 }
    refute_raises { @item.numFmtId = 2 }
    assert_equal(2, @item.numFmtId)
  end

  def test_fomatCode
    assert_raises(ArgumentError) { @item.formatCode = -1.1 }
    refute_raises { @item.formatCode = "0" }
    assert_equal("0", @item.formatCode)
  end
end
