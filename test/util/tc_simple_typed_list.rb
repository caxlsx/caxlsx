# frozen_string_literal: true

require 'tc_helper'

class TestSimpleTypedList < Test::Unit::TestCase
  def setup
    @list = Axlsx::SimpleTypedList.new Integer
  end

  def teardown; end

  def test_type_is_a_class_or_array_of_class
    assert_nothing_raised { Axlsx::SimpleTypedList.new Integer }
    assert_nothing_raised { Axlsx::SimpleTypedList.new [Integer, String] }
    assert_raise(ArgumentError) { Axlsx::SimpleTypedList.new }
    assert_raise(ArgumentError) { Axlsx::SimpleTypedList.new "1" }
    assert_raise(ArgumentError) { Axlsx::SimpleTypedList.new [Integer, "Class"] }
  end

  def test_indexed_based_assignment
    # should not allow nil assignment
    assert_raise(ArgumentError) { @list[0] = nil }
    assert_raise(ArgumentError) { @list[0] = "1" }
    assert_nothing_raised { @list[0] = 1 }
  end

  def test_concat_assignment
    assert_raise(ArgumentError) { @list << nil }
    assert_raise(ArgumentError) { @list << "1" }
    assert_nothing_raised { @list << 1 }
  end

  def test_concat_should_return_index
    assert_empty(@list)
    assert_equal(0, @list << 1)
    assert_equal(1, @list << 2)
    @list.delete_at 0

    assert_equal(1, @list << 3)
    assert_equal(0, @list.index(2))
  end

  def test_push_should_return_index
    assert_equal(0, @list.push(1))
    assert_equal(1, @list.push(2))
    @list.delete_at 0

    assert_equal(1, @list.push(3))
    assert_equal(0, @list.index(2))
  end

  def test_locking
    @list.push 1
    @list.push 2
    @list.push 3
    @list.lock

    assert_raise(ArgumentError) { @list.delete 1  }
    assert_raise(ArgumentError) { @list.delete_at 1 }
    assert_raise(ArgumentError) { @list.delete_at 2 }
    assert_raise(ArgumentError) { @list.insert(1, 3) }
    assert_raise(ArgumentError) { @list[1] = 3 }

    @list.push 4
    assert_nothing_raised { @list.delete_at 3 }
    @list.unlock
    # ignore garbage
    assert_nothing_raised { @list.delete 0 }
    assert_nothing_raised { @list.delete 9 }
  end

  def test_delete
    @list.push 1

    assert_equal(1, @list.size)
    @list.delete 1

    assert_empty(@list)
  end

  def test_equality
    @list.push 1
    @list.push 2

    assert_equal([1, 2], @list)
  end

  def test_to_a
    refute_same(@list, @list.to_a)
    assert_instance_of(Array, @list.to_a)
  end

  def test_to_ary
    assert_same(@list, @list.to_ary)
  end

  def test_insert
    assert_raise(ArgumentError) { @list << nil }

    assert_equal(1, @list.insert(0, 1))
    assert_equal(2, @list.insert(1, 2))
    assert_equal(3, @list.insert(0, 3))

    assert_equal([3, 1, 2], @list)
  end

  def test_setter
    assert_raise(ArgumentError) { @list[0] = nil }

    assert_equal(1, @list[0] = 1)
    assert_equal(2, @list[1] = 2)
    assert_equal(3, @list[0] = 3)

    assert_equal([3, 2], @list)
  end
end
