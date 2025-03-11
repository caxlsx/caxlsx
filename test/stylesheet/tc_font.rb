# frozen_string_literal: true

require 'tc_helper'

class TestFont < Minitest::Test
  def setup
    @item = Axlsx::Font.new
  end

  def teardown; end

  def test_initialiation
    assert_nil(@item.name)
    assert_nil(@item.charset)
    assert_nil(@item.family)
    assert_nil(@item.b)
    assert_nil(@item.i)
    assert_nil(@item.u)
    assert_nil(@item.strike)
    assert_nil(@item.outline)
    assert_nil(@item.shadow)
    assert_nil(@item.condense)
    assert_nil(@item.extend)
    assert_nil(@item.color)
    assert_nil(@item.sz)
  end

  # def name=(v) Axlsx::validate_string v; @name = v end
  def test_name
    assert_raises(ArgumentError) { @item.name = 7 }
    refute_raises { @item.name = "bob" }
    assert_equal("bob", @item.name)
  end

  # def charset=(v) Axlsx::validate_unsigned_int v; @charset = v end
  def test_charset
    assert_raises(ArgumentError) { @item.charset = -7 }
    refute_raises { @item.charset = 5 }
    assert_equal(5, @item.charset)
  end

  # def family=(v) Axlsx::validate_unsigned_int v; @family = v end
  def test_family
    assert_raises(ArgumentError) { @item.family = -7 }
    refute_raises { @item.family = 5 }
    assert_equal(5, @item.family)
  end

  # def b=(v) Axlsx::validate_boolean v; @b = v end
  def test_b
    assert_raises(ArgumentError) { @item.b = -7 }
    refute_raises { @item.b = true }
    assert(@item.b)
  end

  # def i=(v) Axlsx::validate_boolean v; @i = v end
  def test_i
    assert_raises(ArgumentError) { @item.i = -7 }
    refute_raises { @item.i = true }
    assert(@item.i)
  end

  # def u=(v) Axlsx::validate_cell_u v; @u = v end
  def test_u
    assert_raises(ArgumentError) { @item.u = -7 }
    refute_raises { @item.u = :single }
    assert_equal(:single, @item.u)
    doc = Nokogiri::XML(@item.to_xml_string)

    assert(doc.xpath('//u[@val="single"]'))
  end

  def test_u_backward_compatibility
    # backward compatibility for true
    refute_raises { @item.u = true }
    assert_equal(:single, @item.u)

    # backward compatibility for false
    refute_raises { @item.u = false }
    assert_equal(:none, @item.u)
  end

  # def strike=(v) Axlsx::validate_boolean v; @strike = v end
  def test_strike
    assert_raises(ArgumentError) { @item.strike = -7 }
    refute_raises { @item.strike = true }
    assert(@item.strike)
  end

  # def outline=(v) Axlsx::validate_boolean v; @outline = v end
  def test_outline
    assert_raises(ArgumentError) { @item.outline = -7 }
    refute_raises { @item.outline = true }
    assert(@item.outline)
  end

  # def shadow=(v) Axlsx::validate_boolean v; @shadow = v end
  def test_shadow
    assert_raises(ArgumentError) { @item.shadow = -7 }
    refute_raises { @item.shadow = true }
    assert(@item.shadow)
  end

  # def condense=(v) Axlsx::validate_boolean v; @condense = v end
  def test_condense
    assert_raises(ArgumentError) { @item.condense = -7 }
    refute_raises { @item.condense = true }
    assert(@item.condense)
  end

  # def extend=(v) Axlsx::validate_boolean v; @extend = v end
  def test_extend
    assert_raises(ArgumentError) { @item.extend = -7 }
    refute_raises { @item.extend = true }
    assert(@item.extend)
  end

  # def color=(v) DataTypeValidator.validate "Font.color", Color, v; @color=v end
  def test_color
    assert_raises(ArgumentError) { @item.color = -7 }
    refute_raises { @item.color = Axlsx::Color.new(rgb: "00000000") }
    assert_kind_of(Axlsx::Color, @item.color)
  end

  # def sz=(v) Axlsx::validate_unsigned_int v; @sz=v end
  def test_sz
    assert_raises(ArgumentError) { @item.sz = -7 }
    refute_raises { @item.sz = 5 }
    assert_equal(5, @item.sz)
  end
end
