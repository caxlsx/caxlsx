# frozen_string_literal: true

require 'tc_helper'

class RichTextRun < Minitest::Test
  def setup
    @p = Axlsx::Package.new
    @ws = @p.workbook.add_worksheet name: "hmmmz"
    @p.workbook.styles.add_style sz: 20
    @rtr = Axlsx::RichTextRun.new('hihihi', b: true, i: false)
    @rtr2 = Axlsx::RichTextRun.new('hihi2hi2', b: false, i: true)
    @rt = Axlsx::RichText.new
    @rt.runs << @rtr
    @rt.runs << @rtr2
    @row = @ws.add_row [@rt]
    @c = @row.first
  end

  def test_initialize
    assert_equal('hihihi', @rtr.value)
    assert(@rtr.b)
    assert_false(@rtr.i)
  end

  def test_font_size_with_custom_style_and_no_sz
    @c.style = @c.row.worksheet.workbook.styles.add_style bg_color: 'FF00FF'
    sz = @rtr.send(:font_size)

    assert_equal(sz, @c.row.worksheet.workbook.styles.fonts.first.sz * 1.5)
    sz = @rtr2.send(:font_size)

    assert_equal(sz, @c.row.worksheet.workbook.styles.fonts.first.sz)
  end

  def test_font_size_with_bolding
    @c.style = @c.row.worksheet.workbook.styles.add_style b: true

    assert_equal(@c.row.worksheet.workbook.styles.fonts.first.sz * 1.5, @rtr.send(:font_size))
    assert_equal(@c.row.worksheet.workbook.styles.fonts.first.sz * 1.5, @rtr2.send(:font_size)) # is this the correct behaviour?
  end

  def test_font_size_with_custom_sz
    @c.style = @c.row.worksheet.workbook.styles.add_style sz: 52
    sz = @rtr.send(:font_size)

    assert_equal(52 * 1.5, sz)
    sz2 = @rtr2.send(:font_size)

    assert_equal(52, sz2)
  end

  def test_rtr_with_sz
    @rtr.sz = 25

    assert_equal(25, @rtr.send(:font_size))
  end

  def test_color
    assert_raises(ArgumentError) { @rtr.color = -1.1 }
    refute_raises { @rtr.color = "FF00FF00" }
    assert_equal("FF00FF00", @rtr.color.rgb)
  end

  def test_scheme
    assert_raises(ArgumentError) { @rtr.scheme = -1.1 }
    refute_raises { @rtr.scheme = :major }
    assert_equal(:major, @rtr.scheme)
  end

  def test_vertAlign
    assert_raises(ArgumentError) { @rtr.vertAlign = -1.1 }
    refute_raises { @rtr.vertAlign = :baseline }
    assert_equal(:baseline, @rtr.vertAlign)
  end

  def test_sz
    assert_raises(ArgumentError) { @rtr.sz = -1.1 }
    refute_raises { @rtr.sz = 12 }
    assert_equal(12, @rtr.sz)
  end

  def test_extend
    assert_raises(ArgumentError) { @rtr.extend = -1.1 }
    refute_raises { @rtr.extend = false }
    assert_false(@rtr.extend)
  end

  def test_condense
    assert_raises(ArgumentError) { @rtr.condense = -1.1 }
    refute_raises { @rtr.condense = false }
    assert_false(@rtr.condense)
  end

  def test_shadow
    assert_raises(ArgumentError) { @rtr.shadow = -1.1 }
    refute_raises { @rtr.shadow = false }
    assert_false(@rtr.shadow)
  end

  def test_outline
    assert_raises(ArgumentError) { @rtr.outline = -1.1 }
    refute_raises { @rtr.outline = false }
    assert_false(@rtr.outline)
  end

  def test_strike
    assert_raises(ArgumentError) { @rtr.strike = -1.1 }
    refute_raises { @rtr.strike = false }
    assert_false(@rtr.strike)
  end

  def test_u
    @c.type = :string
    assert_raises(ArgumentError) { @c.u = -1.1 }
    refute_raises { @c.u = :single }
    assert_equal(:single, @c.u)
    doc = Nokogiri::XML(@c.to_xml_string(1, 1))

    assert(doc.xpath('//u[@val="single"]'))
  end

  def test_i
    assert_raises(ArgumentError) { @c.i = -1.1 }
    refute_raises { @c.i = false }
    assert_false(@c.i)
  end

  def test_rFont
    assert_raises(ArgumentError) { @c.font_name = -1.1 }
    refute_raises { @c.font_name = "Arial" }
    assert_equal("Arial", @c.font_name)
  end

  def test_charset
    assert_raises(ArgumentError) { @c.charset = -1.1 }
    refute_raises { @c.charset = 1 }
    assert_equal(1, @c.charset)
  end

  def test_family
    assert_raises(ArgumentError) { @rtr.family = 0 }
    refute_raises { @rtr.family = 1 }
    assert_equal(1, @rtr.family)
  end

  def test_b
    assert_raises(ArgumentError) { @c.b = -1.1 }
    refute_raises { @c.b = false }
    assert_false(@c.b)
  end

  def test_multiline_autowidth
    wrap = @p.workbook.styles.add_style({ alignment: { wrap_text: true } })
    awtr = Axlsx::RichTextRun.new("I'm bold\n", b: true)
    rt = Axlsx::RichText.new
    rt.runs << awtr
    @ws.add_row [rt], style: wrap
    ar = [0]
    awtr.autowidth(ar)

    assert_equal(2, ar.length)
    assert_in_delta(13.2, ar[0])
    assert_equal(0, ar[1])
  end

  def test_to_xml
    schema = Nokogiri::XML::Schema(File.open(Axlsx::SML_XSD))
    doc = Nokogiri::XML(@ws.to_xml_string)
    errors = schema.validate(doc)

    assert_empty(errors)

    assert(doc.xpath('//rPr/b[@val=1]'))
    assert(doc.xpath('//rPr/i[@val=0]'))
    assert(doc.xpath('//rPr/b[@val=0]'))
    assert(doc.xpath('//rPr/i[@val=1]'))
    assert(doc.xpath('//is//t[contains(text(), "hihihi")]'))
    assert(doc.xpath('//is//t[contains(text(), "hihi2hi2")]'))
  end
end
