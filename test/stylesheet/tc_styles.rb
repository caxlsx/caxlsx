# frozen_string_literal: true

require 'tc_helper'
require 'support/capture_warnings'

class TestStyles < Minitest::Test
  include CaptureWarnings

  def setup
    @styles = Axlsx::Styles.new
  end

  def teardown; end

  def test_valid_document
    schema = Nokogiri::XML::Schema(File.open(Axlsx::SML_XSD))
    doc = Nokogiri::XML(@styles.to_xml_string)
    errors = schema.validate(doc)

    assert_empty(errors)
  end

  def test_add_style_border_hash
    border_count = @styles.borders.size
    @styles.add_style border: { style: :thin, color: "FFFF0000" }

    assert_equal(@styles.borders.size, border_count + 1)
    assert_equal("FFFF0000", @styles.borders.last.prs.last.color.rgb)
    assert_raises(ArgumentError) { @styles.add_style border: { color: "FFFF0000" } }
    assert_equal(4, @styles.borders.last.prs.size)
  end

  def test_add_style_border_array
    prev_border_count = @styles.borders.size

    borders_array = [
      { style: :thin, color: "DDDDDD" },
      { edges: [:top], style: :thin, color: "000000" },
      { edges: [:bottom], style: :thick, color: "FF0000" },
      { edges: [:left], style: :dotted, color: "FFFF00" },
      { edges: [:right], style: :dashed, color: "FFFFFF" },
      { style: :thick, color: "CCCCCC" }
    ]

    @styles.add_style(border: borders_array)

    assert_equal(@styles.borders.size, (prev_border_count + 1))

    current_border = @styles.borders.last

    borders_array.each do |b_opts|
      next unless b_opts[:edges]

      border_pr = current_border.prs.detect { |x| x.name == b_opts[:edges].first }

      assert_equal(border_pr.color.rgb, "FF#{b_opts[:color]}")
    end
  end

  def test_add_style_border_edges
    @styles.add_style border: { style: :thin, color: "0000FFFF", edges: [:top, :bottom] }
    parts = @styles.borders.last.prs

    parts.each { |pr| assert_equal("0000FFFF", pr.color.rgb, "Style is applied to #{pr.name} properly") }
    assert_equal(2, (parts.map { |pr| pr.name.to_s }.sort && ['bottom', 'top']).size, "specify two edges, and you get two border prs")
  end

  def test_do_not_alter_options_in_add_style
    # This should test all options, but for now - just the bits that we know caused some pain
    options = { border: { style: :thin, color: "FF000000" } }
    @styles.add_style options

    assert_equal(:thin, options[:border][:style], 'thin style is stil in option')
    assert_equal("FF000000", options[:border][:color], 'color is stil in option')
  end

  def test_parse_num_fmt
    f_code = { format_code: "YYYY/MM" }
    num_fmt = { num_fmt: 5 }

    assert_nil(@styles.parse_num_fmt_options, 'noop if neither :format_code or :num_fmt exist')
    max = @styles.numFmts.map(&:numFmtId).max
    @styles.parse_num_fmt_options(f_code)

    assert_equal(@styles.numFmts.last.numFmtId, max + 1, "new numfmts gets next available id")
    assert_kind_of(Integer, @styles.parse_num_fmt_options(num_fmt), "Should return the provided num_fmt if not dxf")
    assert_kind_of(Axlsx::NumFmt, @styles.parse_num_fmt_options(num_fmt.merge({ type: :dxf })), "Makes a new NumFmt if dxf")
  end

  def test_parse_border_options_hash_required_keys
    assert_raises(ArgumentError, "Require color key") { @styles.parse_border_options(border: { style: :thin }) }
    assert_raises(ArgumentError, "Require style key") { @styles.parse_border_options(border: { color: "FF0d0d0d" }) }
    refute_raises { @styles.parse_border_options(border: { style: :thin, color: "FF000000" }) }
  end

  def test_parse_border_basic_options
    b_opts = { border: { diagonalUp: 1, edges: [:left, :right], color: "FFDADADA", style: :thick } }
    b = @styles.parse_border_options b_opts

    assert_kind_of(Integer, b)
    assert_equal(@styles.parse_border_options(b_opts.merge({ type: :dxf })).class, Axlsx::Border)
    assert_equal(1, @styles.borders.last.diagonalUp, "border options are passed in to the initializer")
  end

  def test_parse_border_options_edges
    b_opts = { border: { diagonalUp: 1, edges: [:left, :right], color: "FFDADADA", style: :thick } }
    @styles.parse_border_options b_opts
    b = @styles.borders.last
    left = b.prs.find { |bpr| bpr.name == :left }
    right = b.prs.find { |bpr| bpr.name == :right }
    top = b.prs.find { |bpr| bpr.name == :top }
    bottom = b.prs.find { |bpr| bpr.name == :bottom }

    assert_nil(top, "unspecified top edge should not be created")
    assert_nil(bottom, "unspecified bottom edge should not be created")
    assert_kind_of(Axlsx::BorderPr, left, "specified left edge is set")
    assert_kind_of(Axlsx::BorderPr, right, "specified right edge is set")
    assert_equal(left.style, right.style, "edge parts have the same style")
    assert_equal(:thick, left.style, "the style is THICK")
    assert_equal(right.color.rgb, left.color.rgb, "edge parts are colors are the same")
    assert_equal("FFDADADA", right.color.rgb, "edge color rgb is correct")
  end

  def test_parse_border_options_noop
    assert_nil(@styles.parse_border_options({}), "noop if the border key is not in options")
  end

  def test_parse_border_options_integer_xf
    assert_equal(1, @styles.parse_border_options(border: 1))
    assert_raises(ArgumentError, "unknown border index") { @styles.parse_border_options(border: 100) }
  end

  def test_parse_border_options_integer_dxf
    b_opts = { border: { edges: [:left, :right], color: "FFFFFFFF", style: :thick } }
    b = @styles.parse_border_options(b_opts)
    b2 = @styles.parse_border_options(border: b, type: :dxf)

    assert_kind_of(Axlsx::Border, b2, "Cloned existing border object")
  end

  def test_parse_alignment_options
    assert_nil(@styles.parse_alignment_options, "noop if :alignment is not set")
    assert_kind_of(Axlsx::CellAlignment, @styles.parse_alignment_options(alignment: {}))
  end

  def test_parse_font_using_defaults
    original = @styles.fonts.first
    @styles.add_style b: 1, sz: 99
    created = @styles.fonts.last
    original_attributes = Axlsx.instance_values_for(original)

    assert_equal(1, created.b)
    assert_equal(99, created.sz)
    attributes_to_reject = %w(b sz)
    copied = original_attributes.reject { |key, _value| attributes_to_reject.include? key }
    instance_vals = Axlsx.instance_values_for(created)

    copied.each do |key, value|
      assert_equal(instance_vals[key], value)
    end
  end

  def test_parse_font_options
    options = {
      fg_color: "FF050505",
      sz: 20,
      b: 1,
      i: 1,
      u: :single,
      strike: 1,
      outline: 1,
      shadow: 1,
      charset: 9,
      family: 1,
      font_name: "woot font"
    }

    assert_nil(@styles.parse_font_options, "noop if no font keys are set")
    assert_kind_of(Integer, @styles.parse_font_options(b: 1), "return index of font if not :dxf type")
    assert_equal(@styles.parse_font_options(b: 1, type: :dxf).class, Axlsx::Font, "return font object if :dxf type")

    f = @styles.parse_font_options(options.merge(type: :dxf))
    color = options.delete(:fg_color)
    options[:name] = options.delete(:font_name)

    options.each do |key, value|
      assert_equal(f.send(key), value, "assert that #{key} was parsed")
    end
    assert_equal(f.color.rgb, color)
  end

  def test_parse_fill_options
    assert_nil(@styles.parse_fill_options, "noop if no fill keys are set")

    assert_instance_of(Axlsx::Fill, @styles.parse_fill_options(bg_color: "AB", type: :dxf), "return fill object if :dxf type")
    @styles.parse_fill_options(bg_color: "CD", type: :dxf).tap do |fill|
      assert_equal("FFCDCDCD", fill.fill_type.bgColor.rgb)
      assert_equal(:solid, fill.fill_type.patternType)
    end

    assert_instance_of(Integer, @styles.parse_fill_options(bg_color: "AB"), "return index of fill if not :dxf type")
    @styles.parse_fill_options(bg_color: "DE").tap do |fill_id|
      fill = @styles.fills[fill_id]

      assert_equal("FFDEDEDE", fill.fill_type.fgColor.rgb)
      assert_equal(:solid, fill.fill_type.patternType)
    end
  end

  def test_parse_fill_options_fill_type
    @styles.parse_fill_options(pattern_type: :darkHorizontal, pattern_bg_color: 'AB', pattern_fg_color: 'BC', type: :dxf).tap do |fill|
      assert_equal("FFABABAB", fill.fill_type.bgColor.rgb)
      assert_equal("FFBCBCBC", fill.fill_type.fgColor.rgb)
      assert_equal(:darkHorizontal, fill.fill_type.patternType)
    end

    @styles.parse_fill_options(pattern_type: :darkHorizontal, bg_color: 'AB', type: :dxf).tap do |fill|
      assert_equal("FFABABAB", fill.fill_type.bgColor.rgb, "use bg_color if pattern_bg_color is not defined")
    end

    warnings = capture_warnings do
      @styles.parse_fill_options(pattern_type: :darkHorizontal, bg_color: 'AB', pattern_bg_color: 'BC', type: :dxf).tap do |fill|
        assert_equal("FFBCBCBC", fill.fill_type.bgColor.rgb, "use pattern_bg_color if both bg_color and pattern_bg_color is defined")
      end
    end

    assert_equal 1, warnings.size
    assert_includes warnings.first, 'Both `bg_color` and `pattern_bg_color` got defined. To get a solid background without defining it in `patter_type`, use only `bg_color`, otherwise use only `pattern_bg_color` to avoid confusion.'
  end

  def test_parse_protection_options
    assert_nil(@styles.parse_protection_options, "noop if no protection keys are set")
    assert_equal(@styles.parse_protection_options(hidden: 1).class, Axlsx::CellProtection, "creates a new cell protection object")
  end

  def test_add_style
    fill_count = @styles.fills.size
    font_count = @styles.fonts.size
    xf_count = @styles.cellXfs.size

    @styles.add_style bg_color: "FF000000", fg_color: "FFFFFFFF", sz: 13, num_fmt: Axlsx::NUM_FMT_PERCENT, alignment: { horizontal: :left }, border: Axlsx::STYLE_THIN_BORDER, hidden: true, locked: true

    assert_equal(@styles.fills.size, fill_count + 1)
    assert_equal(@styles.fonts.size, font_count + 1)
    assert_equal(@styles.cellXfs.size, xf_count + 1)
    xf = @styles.cellXfs.last

    assert_equal(xf.fillId, (@styles.fills.size - 1), "points to the last created fill")
    assert_equal("FF000000", @styles.fills.last.fill_type.fgColor.rgb, "fill created with color")

    assert_equal(xf.fontId, (@styles.fonts.size - 1), "points to the last created font")
    assert_equal(13, @styles.fonts.last.sz, "font sz applied")
    assert_equal("FFFFFFFF", @styles.fonts.last.color.rgb, "font color applied")

    assert_equal(xf.borderId, Axlsx::STYLE_THIN_BORDER, "border id is set")
    assert_equal(xf.numFmtId, Axlsx::NUM_FMT_PERCENT, "number format id is set")

    assert_kind_of(Axlsx::CellAlignment, xf.alignment, "alignment was created")
    assert_equal(:left, xf.alignment.horizontal, "horizontal alignment applied")
    assert(xf.protection.hidden, "hidden protection set")
    assert(xf.protection.locked, "cell locking set")
    assert_raises(ArgumentError, "should reject invalid borderId") { @styles.add_style border: 2 }

    assert(xf.applyProtection, "protection applied")
    assert(xf.applyBorder, "border applied")
    assert(xf.applyNumberFormat, "number format applied")
    assert(xf.applyAlignment, "alignment applied")
  end

  def test_basic_add_style_dxf
    border_count = @styles.borders.size
    @styles.add_style border: { style: :thin, color: "FFFF0000" }, type: :dxf

    assert_equal(@styles.borders.size, border_count, "styles borders not affected")
    assert_equal("FFFF0000", @styles.dxfs.last.border.prs.last.color.rgb)
    assert_raises(ArgumentError) { @styles.add_style border: { color: "FFFF0000" }, type: :dxf }
    assert_equal(4, @styles.borders.last.prs.size)
  end

  def test_add_style_dxf
    fill_count = @styles.fills.size
    font_count = @styles.fonts.size
    dxf_count = @styles.dxfs.size

    style = @styles.add_style bg_color: "FF000000", fg_color: "FFFFFFFF", sz: 13, alignment: { horizontal: :left }, border: { style: :thin, color: "FFFF0000" }, hidden: true, locked: true, type: :dxf

    assert_equal(@styles.dxfs.size, dxf_count + 1)
    assert_equal(0, style, "returns the zero-based dxfId")

    dxf = @styles.dxfs.last

    assert_equal("FF000000", @styles.dxfs.last.fill.fill_type.bgColor.rgb, "fill created with color")

    assert_equal(font_count, @styles.fonts.size, "font not created under styles")
    assert_equal(fill_count, @styles.fills.size, "fill not created under styles")

    assert_kind_of(Axlsx::Border, dxf.border, "border is set")
    assert_nil(dxf.numFmt, "number format is not set")

    assert_kind_of(Axlsx::CellAlignment, dxf.alignment, "alignment was created")
    assert_equal(:left, dxf.alignment.horizontal, "horizontal alignment applied")
    assert(dxf.protection.hidden, "hidden protection set")
    assert(dxf.protection.locked, "cell locking set")
    assert_raises(ArgumentError, "should reject invalid borderId") { @styles.add_style border: 3 }
  end

  def test_multiple_dxf
    # add a second style
    style = @styles.add_style bg_color: "00000000", fg_color: "FFFFFFFF", sz: 13, alignment: { horizontal: :left }, border: { style: :thin, color: "FFFF0000" }, hidden: true, locked: true, type: :dxf

    assert_equal(0, style, "returns the first dxfId")
    style = @styles.add_style bg_color: "FF000000", fg_color: "FFFFFFFF", sz: 13, alignment: { horizontal: :left }, border: { style: :thin, color: "FFFF0000" }, hidden: true, locked: true, type: :dxf

    assert_equal(1, style, "returns the second dxfId")
  end

  def test_valid_document_with_font_options
    font_options = {
      fg_color: "FF050505",
      sz: 20,
      b: 1,
      i: 1,
      u: :single,
      strike: 1,
      outline: 1,
      shadow: 1,
      charset: 9,
      family: 1,
      font_name: "woot font"
    }
    @styles.add_style font_options

    schema = Nokogiri::XML::Schema(File.open(Axlsx::SML_XSD))
    doc = Nokogiri::XML(@styles.to_xml_string)
    errors = schema.validate(doc)

    assert_empty(errors)
  end

  def test_border_top_without_border_regression
    # https://github.com/axlsx-styler-gem/axlsx_styler/issues/31

    borders = {
      top: { style: :double, color: '0000FF' },
      right: { style: :thick, color: 'FF0000' },
      bottom: { style: :double, color: '0000FF' },
      left: { style: :thick, color: 'FF0000' }
    }

    borders.each do |edge, b_opts|
      @styles.add_style("border_#{edge}": b_opts)

      current_border = @styles.borders.last

      border_pr = current_border.prs.detect { |x| x.name == edge }

      assert_equal(border_pr.color.rgb, "FF#{b_opts[:color]}")
    end
  end
end
