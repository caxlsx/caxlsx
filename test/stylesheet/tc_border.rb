# frozen_string_literal: true

require 'tc_helper'

class TestBorder < Minitest::Test
  def setup
    @b = Axlsx::Border.new
  end

  def teardown; end

  def test_initialiation
    assert_nil(@b.diagonalUp)
    assert_nil(@b.diagonalDown)
    assert_nil(@b.outline)
    assert_kind_of(Axlsx::SimpleTypedList, @b.prs)
  end

  def test_diagonalUp
    assert_raises(ArgumentError) { @b.diagonalUp = :red }
    refute_raises { @b.diagonalUp = true }
    assert(@b.diagonalUp)
  end

  def test_diagonalDown
    assert_raises(ArgumentError) { @b.diagonalDown = :red }
    refute_raises { @b.diagonalDown = true }
    assert(@b.diagonalDown)
  end

  def test_outline
    assert_raises(ArgumentError) { @b.outline = :red }
    refute_raises { @b.outline = true }
    assert(@b.outline)
  end

  def test_prs
    refute_raises { @b.prs << Axlsx::BorderPr.new(name: :top, style: :thin, color: Axlsx::Color.new(rgb: "FF0000FF")) }
  end
end
