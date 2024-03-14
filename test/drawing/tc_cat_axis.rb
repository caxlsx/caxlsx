# frozen_string_literal: true

require 'tc_helper'

class TestCatAxis < Minitest::Test
  def setup
    @axis = Axlsx::CatAxis.new
  end

  def teardown; end

  def test_initialization
    assert_equal(1, @axis.auto, "axis auto default incorrect")
    assert_equal(:ctr, @axis.lbl_algn, "label align default incorrect")
    assert_equal("100", @axis.lbl_offset, "label offset default incorrect")
  end

  def test_auto
    assert_raises(ArgumentError, "requires valid auto") { @axis.auto = :nowhere }
    refute_raises { @axis.auto = false }
  end

  def test_lbl_algn
    assert_raises(ArgumentError, "requires valid label alignment") { @axis.lbl_algn = :nowhere }
    refute_raises { @axis.lbl_algn = :r }
  end

  def test_lbl_offset
    assert_raises(ArgumentError, "requires valid label offset") { @axis.lbl_offset = 'foo' }
    refute_raises { @axis.lbl_offset = "20" }
  end
end
