# frozen_string_literal: true

require 'tc_helper'

class TestSimpleTypedList < Minitest::Test
  def setup
    @list = Axlsx::SimpleTypedList.new Integer
  end

  def teardown; end

  def test_type_is_a_class_or_array_of_class
    refute_raises { Axlsx::SimpleTypedList.new Integer }
    refute_raises { Axlsx::SimpleTypedList.new [Integer, String] }
    assert_raises(ArgumentError) { Axlsx::SimpleTypedList.new }
    assert_raises(ArgumentError) { Axlsx::SimpleTypedList.new "1" }
    assert_raises(ArgumentError) { Axlsx::SimpleTypedList.new [Integer, "Class"] }
  end

  def test_indexed_based_assignment
    # should not allow nil assignment
    assert_raises(ArgumentError) { @list[0] = nil }
    assert_raises(ArgumentError) { @list[0] = "1" }
    refute_raises { @list[0] = 1 }
  end

  def test_push_operator_assignment
    assert_raises(ArgumentError) { @list << nil }
    assert_raises(ArgumentError) { @list << "1" }
    refute_raises { @list << 1 }
  end

  def test_push_operator_returns_index
    assert_empty(@list)
    assert_equal(0, @list << 1)
    assert_equal(1, @list << 2)
    @list.delete_at 0

    assert_equal(1, @list << 3)
    assert_equal(0, @list.index(2))
  end

  # rubocop:disable Style/ConcatArrayLiterals
  def test_concat_assignment
    assert_raises(ArgumentError) { @list.concat([1, nil]) }
    refute_raises { @list.concat [1] }
  end

  def test_concat_multiple_arrays
    @list.concat([1], [2])

    assert_equal([1, 2], @list)
  end

  def test_concat_returns_self
    assert_equal(@list, @list.concat([1]))
  end

  def test_concat_mutates_object
    @list.concat([1])

    assert_equal([1], @list)
  end
  # rubocop:enable Style/ConcatArrayLiterals

  def test_push_assignment
    assert_raises(ArgumentError) { @list.push("1") }
    assert_raises(ArgumentError) { @list.push(1, nil) }
    refute_raises { @list.push(1, 2) }
  end

  def test_push_returns_self
    assert_equal(@list, @list.push(1))
  end

  def test_push_mutates_object
    @list.push(1)

    assert_equal([1], @list)
  end

  def test_locking
    @list.push 1, 2, 3
    @list.lock

    assert_raises(ArgumentError) { @list.delete 1  }
    assert_raises(ArgumentError) { @list.delete_at 1 }
    assert_raises(ArgumentError) { @list.delete_at 2 }
    assert_raises(ArgumentError) { @list.insert(1, 3) }
    assert_raises(ArgumentError) { @list[1] = 3 }

    @list.push 4
    refute_raises { @list.delete_at 3 }
    @list.unlock
    # ignore garbage
    refute_raises { @list.delete 0 }
    refute_raises { @list.delete 9 }
  end

  def test_delete
    @list.push 1

    assert_equal(1, @list.size)
    @list.delete 1

    assert_empty(@list)
  end

  def test_equality
    @list.push 1, 2

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
    assert_raises(ArgumentError) { @list << nil }

    assert_equal(1, @list.insert(0, 1))
    assert_equal(2, @list.insert(1, 2))
    assert_equal(3, @list.insert(0, 3))

    assert_equal([3, 1, 2], @list)
  end

  def test_setter
    assert_raises(ArgumentError) { @list[0] = nil }

    assert_equal(1, @list[0] = 1)
    assert_equal(2, @list[1] = 2)
    assert_equal(3, @list[0] = 3)

    assert_equal([3, 2], @list)
  end
end
