# frozen_string_literal: true

require 'tc_helper'

class TestCell < Minitest::Test
  def setup
    p = Axlsx::Package.new
    p.use_shared_strings = true
    @ws = p.workbook.add_worksheet name: "hmmm"
    p.workbook.styles.add_style sz: 20
    @row = @ws.add_row
    @c = @row.add_cell 1, type: :float, style: 1, escape_formulas: true
    data = (0..26).map { |index| index }
    @ws.add_row data
    @cAA = @ws["AA2"]
  end

  def test_initialize
    assert_equal(@row.cells.last, @c, "the cell was added to the row")
    assert_equal(:float, @c.type, "type option is applied")
    assert_equal(1, @c.style, "style option is applied")
    assert_in_delta(@c.value, 1.0, 0.001, "type option is applied and value is casted")
    assert(@c.escape_formulas, "escape formulas option is applied")
  end

  def test_style_date_data
    c = Axlsx::Cell.new(@c.row, Time.now)

    assert_equal(Axlsx::STYLE_DATE, c.style)
  end

  def test_row
    assert_equal(@c.row, @row)
  end

  def test_index
    assert_equal(@c.index, @row.cells.index(@c))
  end

  def test_pos
    assert_equal(@c.pos, [@c.index, @c.row.index(@c)])
  end

  def test_r
    assert_equal("A1", @c.r, "calculate cell reference")
  end

  def test_wide_r
    assert_equal("AA2", @cAA.r, "calculate cell reference")
  end

  def test_r_abs
    assert_equal("$A$1", @c.r_abs, "calculate absolute cell reference")
    assert_equal("$AA$2", @cAA.r_abs, "needs to accept multi-digit columns")
  end

  def test_name
    @c.name = 'foo'

    assert_equal(1, @ws.workbook.defined_names.size)
    assert_equal('foo', @ws.workbook.defined_names.last.name)
  end

  def test_autowidth
    style = @c.row.worksheet.workbook.styles.add_style({ alignment: { horizontal: :center, vertical: :center, wrap_text: true } })
    @c.style = style

    assert_in_delta(6.6, @c.autowidth, 0.01)
  end

  def test_autowidth_with_bold_font_multiplier
    style = @c.row.worksheet.workbook.styles.add_style(b: true)
    @c.row.worksheet.workbook.bold_font_multiplier = 1.05
    @c.style = style

    assert_in_delta(6.93, @c.autowidth, 0.01)
  end

  def test_autowidth_with_font_scale_divisor
    @c.row.worksheet.workbook.font_scale_divisor = 11.0

    assert_in_delta(6.0, @c.autowidth, 0.01)
  end

  def test_time
    @c.type = :time
    now = DateTime.now
    @c.value = now

    assert_equal(@c.value, now.to_time)
  end

  def test_date
    @c.type = :date
    now = Time.now
    @c.value = now

    assert_equal(@c.value, now.to_date)
  end

  def test_style
    assert_raises(ArgumentError, "must reject invalid style indexes") { @c.style = @c.row.worksheet.workbook.styles.cellXfs.size }
    refute_raises { @c.style = 1 }
    assert_equal(1, @c.style)
  end

  def test_type
    assert_raises(ArgumentError, "type must be :string, :integer, :float, :date, :time, :boolean") { @c.type = :array }
    refute_raises { @c.type = :string }
    assert_equal("1.0", @c.value, "changing type casts the value")
    assert_equal(:float, @row.add_cell(1.0 / (10**7)).type, 'properly identify exponential floats as float type')
    assert_equal(:time, @row.add_cell(Time.now).type, 'time should be time')
    assert_equal(:date, @row.add_cell(Date.today).type, 'date should be date')
    assert_equal(:boolean, @row.add_cell(true).type, 'boolean should be boolean')
  end

  def test_value
    assert_raises(ArgumentError, "type must be :string, :integer, :float, :date, :time, :boolean") { @c.type = :array }
    refute_raises { @c.type = :string }
    assert_equal("1.0", @c.value, "changing type casts the value")
  end

  def test_col_ref
    # TODO: move to axlsx spec
    assert_equal("A", Axlsx.col_ref(0))
  end

  def test_cell_type_from_value
    assert_equal(:float, @c.send(:cell_type_from_value, 1.0))
    assert_equal(:float, @c.send(:cell_type_from_value, "1e1"))
    assert_equal(:float, @c.send(:cell_type_from_value, "1e#{Float::MAX_10_EXP}"))
    assert_equal(:string, @c.send(:cell_type_from_value, "1e#{Float::MAX_10_EXP + 1}"))
    assert_equal(:float, @c.send(:cell_type_from_value, "1e-1"))
    assert_equal(:float, @c.send(:cell_type_from_value, "1e#{Float::MIN_10_EXP}"))
    assert_equal(:string, @c.send(:cell_type_from_value, "1e#{Float::MIN_10_EXP - 1}"))
    assert_equal(:integer, @c.send(:cell_type_from_value, 1))
    assert_equal(:date, @c.send(:cell_type_from_value, Date.today))
    assert_equal(:time, @c.send(:cell_type_from_value, Time.now))
    assert_equal(:string, @c.send(:cell_type_from_value, []))
    assert_equal(:string, @c.send(:cell_type_from_value, "d"))
    assert_equal(:string, @c.send(:cell_type_from_value, nil))
    assert_equal(:integer, @c.send(:cell_type_from_value, -1))
    assert_equal(:boolean, @c.send(:cell_type_from_value, true))
    assert_equal(:boolean, @c.send(:cell_type_from_value, false))
    assert_equal(:float, @c.send(:cell_type_from_value, 1.0 / (10**6)))
    assert_equal(:richtext, @c.send(:cell_type_from_value, Axlsx::RichText.new))
    assert_equal(:iso_8601, @c.send(:cell_type_from_value, '2008-08-30T01:45:36.123+09:00'))
  end

  def test_cell_type_from_value_looks_like_number_but_is_not
    mimic_number = Class.new do
      def initialize(to_s_value)
        @to_s_value = to_s_value
      end

      def to_s
        @to_s_value
      end
    end

    number_strings = [
      '1',
      '1234567890',
      '1.0',
      '1e1',
      '0',
      "1e#{Float::MIN_10_EXP}"
    ]

    number_strings.each do |number_string|
      assert_equal(:string, @c.send(:cell_type_from_value, mimic_number.new(number_string)))
    end
  end

  def test_cast_value
    @c.type = :string

    assert_equal("1.0", @c.send(:cast_value, 1.0))
    @c.type = :integer

    assert_equal(1, @c.send(:cast_value, 1.0))
    @c.type = :float

    assert_in_delta(@c.send(:cast_value, "1.0"), 1.0)
    @c.type = :string

    assert_nil(@c.send(:cast_value, nil))
    @c.type = :richtext

    assert_nil(@c.send(:cast_value, nil))
    @c.type = :float

    assert_nil(@c.send(:cast_value, nil))
    @c.type = :boolean

    assert_equal(1, @c.send(:cast_value, true))
    assert_equal(0, @c.send(:cast_value, false))
    @c.type = :iso_8601

    assert_equal("2012-10-10T12:24", @c.send(:cast_value, "2012-10-10T12:24"))
  end

  def test_cast_time_subclass
    subtime = Class.new(Time) do
      def to_time
        raise "#to_time of Time subclass should not be called"
      end
    end

    time = subtime.now

    @c.type = :time

    assert_equal(time, @c.send(:cast_value, time))
  end

  def test_color
    assert_raises(ArgumentError) { @c.color = -1.1 }
    refute_raises { @c.color = "FF00FF00" }
    assert_equal("FF00FF00", @c.color.rgb)
  end

  def test_scheme
    assert_raises(ArgumentError) { @c.scheme = -1.1 }
    refute_raises { @c.scheme = :major }
    assert_equal(:major, @c.scheme)
  end

  def test_vertAlign
    assert_raises(ArgumentError) { @c.vertAlign = -1.1 }
    refute_raises { @c.vertAlign = :baseline }
    assert_equal(:baseline, @c.vertAlign)
  end

  def test_sz
    assert_raises(ArgumentError) { @c.sz = -1.1 }
    refute_raises { @c.sz = 12 }
    assert_equal(12, @c.sz)
  end

  def test_extend
    assert_raises(ArgumentError) { @c.extend = -1.1 }
    refute_raises { @c.extend = false }
    assert_false(@c.extend)
  end

  def test_condense
    assert_raises(ArgumentError) { @c.condense = -1.1 }
    refute_raises { @c.condense = false }
    assert_false(@c.condense)
  end

  def test_shadow
    assert_raises(ArgumentError) { @c.shadow = -1.1 }
    refute_raises { @c.shadow = false }
    assert_false(@c.shadow)
  end

  def test_outline
    assert_raises(ArgumentError) { @c.outline = -1.1 }
    refute_raises { @c.outline = false }
    assert_false(@c.outline)
  end

  def test_strike
    assert_raises(ArgumentError) { @c.strike = -1.1 }
    refute_raises { @c.strike = false }
    assert_false(@c.strike)
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
    assert_raises(ArgumentError) { @c.family = -1.1 }
    refute_raises { @c.family = 5 }
    assert_equal(5, @c.family)
  end

  def test_b
    assert_raises(ArgumentError) { @c.b = -1.1 }
    refute_raises { @c.b = false }
    assert_false(@c.b)
  end

  def test_merge_with_string
    @c.row.add_cell 2
    @c.row.add_cell 3
    @c.merge "A2"

    assert_equal("A1:A2", @c.row.worksheet.send(:merged_cells).last)
  end

  def test_merge_with_cell
    @c.row.add_cell 2
    @c.row.add_cell 3
    @c.merge @row.cells.last

    assert_equal("A1:C1", @c.row.worksheet.send(:merged_cells).last)
  end

  def test_reverse_merge_with_cell
    @c.row.add_cell 2
    @c.row.add_cell 3
    @row.cells.last.merge @c

    assert_equal("A1:C1", @c.row.worksheet.send(:merged_cells).last)
  end

  def test_ssti
    assert_raises(ArgumentError, "ssti must be an unsigned integer!") { @c.send(:ssti=, -1) }
    @c.send :ssti=, 1

    assert_equal(1, @c.ssti)
  end

  def test_plain_string
    @c.escape_formulas = false

    @c.type = :integer

    refute_predicate(@c, :plain_string?)

    @c.type = :string
    @c.value = 'plain string'

    assert_predicate(@c, :plain_string?)

    @c.value = nil

    refute_predicate(@c, :plain_string?)

    @c.value = ''

    refute_predicate(@c, :plain_string?)

    @c.value = '=sum'

    refute_predicate(@c, :plain_string?)

    @c.value = '{=sum}'

    refute_predicate(@c, :plain_string?)

    @c.escape_formulas = true

    @c.value = '=sum'

    assert_predicate(@c, :plain_string?)

    @c.value = '{=sum}'

    assert_predicate(@c, :plain_string?)

    @c.value = 'plain string'
    @c.font_name = 'Arial'

    refute_predicate(@c, :plain_string?)
  end

  def test_to_xml_string
    c_xml = Nokogiri::XML(@c.to_xml_string(1, 1))

    assert_equal(1, c_xml.xpath("/c[@s=1]").size)
  end

  def test_to_xml_string_nil
    @c.value = nil
    c_xml = Nokogiri::XML(@c.to_xml_string(1, 1))

    assert_equal(1, c_xml.xpath("/c[@s=1]").size)
  end

  def test_to_xml_string_with_run
    # Actually quite a number of similar run styles
    # but the processing should be the same
    @c.b = true
    @c.type = :string
    @c.value = "a"
    @c.font_name = 'arial'
    @c.color = 'FF0000'
    c_xml = Nokogiri::XML(@c.to_xml_string(1, 1))

    assert_predicate(c_xml.xpath("//b"), :any?)
  end

  def test_to_xml_string_formula
    p = Axlsx::Package.new
    ws = p.workbook.add_worksheet(escape_formulas: false) do |sheet|
      sheet.add_row ["=IF(2+2=4,4,5)"]
    end
    doc = Nokogiri::XML(ws.to_xml_string)
    doc.remove_namespaces!

    assert_predicate(doc.xpath("//f[text()='IF(2+2=4,4,5)']"), :any?)
  end

  def test_to_xml_string_formula_escaped
    p = Axlsx::Package.new
    ws = p.workbook.add_worksheet do |sheet|
      sheet.add_row ["=IF(2+2=4,4,5)"], escape_formulas: true
    end
    doc = Nokogiri::XML(ws.to_xml_string)
    doc.remove_namespaces!

    assert_predicate(doc.xpath("//t[text()='=IF(2+2=4,4,5)']"), :any?)
  end

  def test_to_xml_string_numeric_escaped
    p = Axlsx::Package.new
    ws = p.workbook.add_worksheet do |sheet|
      sheet.add_row ["-1", "+2"], escape_formulas: true, types: :text
    end
    doc = Nokogiri::XML(ws.to_xml_string)
    doc.remove_namespaces!

    assert_predicate(doc.xpath("//t[text()='-1']"), :any?)
    assert_predicate(doc.xpath("//t[text()='+2']"), :any?)
  end

  def test_to_xml_string_owasp_prefixes_that_are_no_excel_formulas
    # OWASP mentions various prefixes that might designate formulas when data is read as CSV:
    # https://owasp.org/www-community/attacks/CSV_Injection
    # Except for `=` none of these prefixes are valid prefixes for formulas in Excel however,
    # so they should never be interpreted / serialized as formulas by Caxlsx.
    p = Axlsx::Package.new
    ws = p.workbook.add_worksheet do |sheet|
      sheet.add_row [
        "@1",
        "%2",
        "|3",
        "\rfoo",
        "\tbar"
      ], escape_formulas: false
    end
    doc = Nokogiri::XML(ws.to_xml_string)
    doc.remove_namespaces!

    assert_predicate(doc.xpath("//t[text()='@1']"), :any?)
    assert_predicate(doc.xpath("//t[text()='%2']"), :any?)
    assert_predicate(doc.xpath("//t[text()='|3']"), :any?)
    assert_predicate(doc.xpath("//t[text()='\nfoo']"), :any?)
    assert_predicate(doc.xpath("//t[text()='\tbar']"), :any?)
  end

  def test_to_xml_string_owasp_prefixes_that_are_no_excel_formulas_with_escape_formulas
    # OWASP mentions various prefixes that might designate formulas when data is read as CSV:
    # https://owasp.org/www-community/attacks/CSV_Injection
    # Except for `=` none of these prefixes are valid prefixes for formulas in Excel however,
    # so they should never be interpreted / serialized as formulas by Caxlsx.
    p = Axlsx::Package.new
    ws = p.workbook.add_worksheet do |sheet|
      sheet.add_row [
        "@1",
        "%2",
        "|3",
        "\rfoo",
        "\tbar"
      ], escape_formulas: true
    end
    doc = Nokogiri::XML(ws.to_xml_string)
    doc.remove_namespaces!

    assert_predicate(doc.xpath("//t[text()='@1']"), :any?)
    assert_predicate(doc.xpath("//t[text()='%2']"), :any?)
    assert_predicate(doc.xpath("//t[text()='|3']"), :any?)
    assert_predicate(doc.xpath("//t[text()='\nfoo']"), :any?)
    assert_predicate(doc.xpath("//t[text()='\tbar']"), :any?)
  end

  def test_to_xml_string_formula_escape_array_parameter
    p = Axlsx::Package.new
    ws = p.workbook.add_worksheet do |sheet|
      sheet.add_row [
        "=IF(2+2=4,4,5)",
        "=IF(13+13=4,4,5)",
        "=IF(99+99=4,4,5)"
      ], escape_formulas: [true, false, true]
    end
    doc = Nokogiri::XML(ws.to_xml_string)
    doc.remove_namespaces!

    assert_predicate(doc.xpath("//t[text()='=IF(2+2=4,4,5)']"), :any?)
    assert_predicate(doc.xpath("//f[text()='IF(13+13=4,4,5)']"), :any?)
    assert_predicate(doc.xpath("//t[text()='=IF(99+99=4,4,5)']"), :any?)
  end

  def test_to_xml_string_array_formula
    p = Axlsx::Package.new
    ws = p.workbook.add_worksheet(escape_formulas: false) do |sheet|
      sheet.add_row ["{=SUM(C2:C11*D2:D11)}"]
    end
    doc = Nokogiri::XML(ws.to_xml_string)
    doc.remove_namespaces!

    assert_predicate(doc.xpath("//f[text()='SUM(C2:C11*D2:D11)']"), :any?)
    assert_predicate(doc.xpath("//f[@t='array']"), :any?)
    assert_predicate(doc.xpath("//f[@ref='A1']"), :any?)
  end

  def test_to_xml_string_text_formula
    p = Axlsx::Package.new
    ws = p.workbook.add_worksheet do |sheet|
      sheet.add_row ["=1+1", "-1+1"], types: :text
    end
    doc = Nokogiri::XML(ws.to_xml_string)
    doc.remove_namespaces!

    assert_empty(doc.xpath("//f[text()='1+1']"))
    assert_predicate(doc.xpath("//t[text()='=1+1']"), :any?)

    assert_empty(doc.xpath("//f[text()='1+1']"))
    assert_predicate(doc.xpath("//t[text()='-1+1']"), :any?)
  end

  def test_font_size_with_custom_style_and_no_sz
    @c.style = @c.row.worksheet.workbook.styles.add_style bg_color: 'FF00FF'
    sz = @c.send(:font_size)

    assert_equal(sz, @c.row.worksheet.workbook.styles.fonts.first.sz)
  end

  def test_font_size_with_bolding
    @c.style = @c.row.worksheet.workbook.styles.add_style b: true

    assert_equal(@c.row.worksheet.workbook.styles.fonts.first.sz * 1.5, @c.send(:font_size))
  end

  def test_font_size_with_custom_sz
    @c.style = @c.row.worksheet.workbook.styles.add_style sz: 52
    sz = @c.send(:font_size)

    assert_equal(52, sz)
  end

  def test_cell_with_sz
    @c.sz = 25

    assert_equal(25, @c.send(:font_size))
  end

  def test_to_xml
    # TODO: This could use some much more stringent testing related to the xml content generated!
    @ws.add_row [Time.now, Date.today, true, 1, 1.0, "text", "=sum(A1:A2)", "2013-01-13T13:31:25.123"]
    @ws.rows.last.cells[5].u = :single

    schema = Nokogiri::XML::Schema(File.open(Axlsx::SML_XSD))
    doc = Nokogiri::XML(@ws.to_xml_string)
    errors = schema.validate(doc)

    assert_empty(errors)
  end
end
