# frozen_string_literal: true

require 'tc_helper'

class TestPictureLocking < Minitest::Test
  def setup
    @item = Axlsx::PictureLocking.new
  end

  def teardown; end

  def test_initialiation
    assert_equal(1, Axlsx.instance_values_for(@item).size)
    assert(@item.noChangeAspect)
  end

  def test_noGrp
    assert_raises(ArgumentError) { @item.noGrp = -1 }
    refute_raises { @item.noGrp = false }
    assert_false(@item.noGrp)
  end

  def test_noRot
    assert_raises(ArgumentError) { @item.noRot = -1 }
    refute_raises { @item.noRot = false }
    assert_false(@item.noRot)
  end

  def test_noChangeAspect
    assert_raises(ArgumentError) { @item.noChangeAspect = -1 }
    refute_raises { @item.noChangeAspect = false }
    assert_false(@item.noChangeAspect)
  end

  def test_noMove
    assert_raises(ArgumentError) { @item.noMove = -1 }
    refute_raises { @item.noMove = false }
    assert_false(@item.noMove)
  end

  def test_noResize
    assert_raises(ArgumentError) { @item.noResize = -1 }
    refute_raises { @item.noResize = false }
    assert_false(@item.noResize)
  end

  def test_noEditPoints
    assert_raises(ArgumentError) { @item.noEditPoints = -1 }
    refute_raises { @item.noEditPoints = false }
    assert_false(@item.noEditPoints)
  end

  def test_noAdjustHandles
    assert_raises(ArgumentError) { @item.noAdjustHandles = -1 }
    refute_raises { @item.noAdjustHandles = false }
    assert_false(@item.noAdjustHandles)
  end

  def test_noChangeArrowheads
    assert_raises(ArgumentError) { @item.noChangeArrowheads = -1 }
    refute_raises { @item.noChangeArrowheads = false }
    assert_false(@item.noChangeArrowheads)
  end

  def test_noChangeShapeType
    assert_raises(ArgumentError) { @item.noChangeShapeType = -1 }
    refute_raises { @item.noChangeShapeType = false }
    assert_false(@item.noChangeShapeType)
  end
end
