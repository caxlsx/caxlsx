# frozen_string_literal: true

require 'tc_helper'

class TestCellAlignment < Test::Unit::TestCase
  def setup
    @item = Axlsx::CellAlignment.new
  end

  def test_initialiation
    assert_nil(@item.horizontal)
    assert_nil(@item.vertical)
    assert_nil(@item.textRotation)
    assert_nil(@item.wrapText)
    assert_nil(@item.indent)
    assert_nil(@item.relativeIndent)
    assert_nil(@item.justifyLastLine)
    assert_nil(@item.shrinkToFit)
    assert_nil(@item.readingOrder)
    options = { :horizontal => :left, :vertical => :top, :textRotation => 3,
                :wrapText => true, :indent => 2, :relativeIndent => 5,
                :justifyLastLine => true, :shrinkToFit => true, :readingOrder => 2 }
    ca = Axlsx::CellAlignment.new options

    options.each do |key, value|
      assert_equal(ca.send(key.to_sym), value)
    end
  end

  def test_horizontal
    assert_raise(ArgumentError) { @item.horizontal = :red }
    assert_nothing_raised { @item.horizontal = :left }
    assert_equal(:left, @item.horizontal)
  end

  def test_vertical
    assert_raise(ArgumentError) { @item.vertical = :red }
    assert_nothing_raised { @item.vertical = :top }
    assert_equal(:top, @item.vertical)
  end

  def test_textRotation
    assert_raise(ArgumentError) { @item.textRotation = -1 }
    assert_nothing_raised { @item.textRotation = 5 }
    assert_equal(5, @item.textRotation)
  end

  def test_wrapText
    assert_raise(ArgumentError) { @item.wrapText = -1 }
    assert_nothing_raised { @item.wrapText = false }
    refute(@item.wrapText)
  end

  def test_indent
    assert_raise(ArgumentError) { @item.indent = -1 }
    assert_nothing_raised { @item.indent = 5 }
    assert_equal(5, @item.indent)
  end

  def test_relativeIndent
    assert_raise(ArgumentError) { @item.relativeIndent = :symbol }
    assert_nothing_raised { @item.relativeIndent = 5 }
    assert_equal(5, @item.relativeIndent)
  end

  def test_justifyLastLine
    assert_raise(ArgumentError) { @item.justifyLastLine = -1 }
    assert_nothing_raised { @item.justifyLastLine = true }
    assert(@item.justifyLastLine)
  end

  def test_shrinkToFit
    assert_raise(ArgumentError) { @item.shrinkToFit = -1 }
    assert_nothing_raised { @item.shrinkToFit = true }
    assert(@item.shrinkToFit)
  end

  def test_readingOrder
    assert_raise(ArgumentError) { @item.readingOrder = -1 }
    assert_nothing_raised { @item.readingOrder = 2 }
    assert_equal(2, @item.readingOrder)
  end
end
