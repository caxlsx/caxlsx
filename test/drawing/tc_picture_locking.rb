require 'tc_helper.rb'

class TestPictureLocking < Test::Unit::TestCase
  def setup
    @item = Axlsx::PictureLocking.new
  end

  def teardown; end

  def test_initialiation
    assert_equal(1, Axlsx.instance_values_for(@item).size)
    assert(@item.noChangeAspect)
  end

  def test_noGrp
    assert_raise(ArgumentError) { @item.noGrp = -1 }
    assert_nothing_raised { @item.noGrp = false }
    refute(@item.noGrp)
  end

  def test_noRot
    assert_raise(ArgumentError) { @item.noRot = -1 }
    assert_nothing_raised { @item.noRot = false }
    refute(@item.noRot)
  end

  def test_noChangeAspect
    assert_raise(ArgumentError) { @item.noChangeAspect = -1 }
    assert_nothing_raised { @item.noChangeAspect = false }
    refute(@item.noChangeAspect)
  end

  def test_noMove
    assert_raise(ArgumentError) { @item.noMove = -1 }
    assert_nothing_raised { @item.noMove = false }
    refute(@item.noMove)
  end

  def test_noResize
    assert_raise(ArgumentError) { @item.noResize = -1 }
    assert_nothing_raised { @item.noResize = false }
    refute(@item.noResize)
  end

  def test_noEditPoints
    assert_raise(ArgumentError) { @item.noEditPoints = -1 }
    assert_nothing_raised { @item.noEditPoints = false }
    refute(@item.noEditPoints)
  end

  def test_noAdjustHandles
    assert_raise(ArgumentError) { @item.noAdjustHandles = -1 }
    assert_nothing_raised { @item.noAdjustHandles = false }
    refute(@item.noAdjustHandles)
  end

  def test_noChangeArrowheads
    assert_raise(ArgumentError) { @item.noChangeArrowheads = -1 }
    assert_nothing_raised { @item.noChangeArrowheads = false }
    refute(@item.noChangeArrowheads)
  end

  def test_noChangeShapeType
    assert_raise(ArgumentError) { @item.noChangeShapeType = -1 }
    assert_nothing_raised { @item.noChangeShapeType = false }
    refute(@item.noChangeShapeType)
  end
end
