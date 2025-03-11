# frozen_string_literal: true

require 'tc_helper'

class TestWorkbook < Minitest::Test
  def setup
    p = Axlsx::Package.new
    @wb = p.workbook
  end

  def teardown; end

  def test_worksheet_users_xml_space
    sheet = @wb.add_worksheet(name: 'foo')
    ws_xml = Nokogiri::XML(sheet.to_xml_string)

    assert(ws_xml.xpath("//xmlns:worksheet/@xml:space='preserve'"))

    @wb.xml_space = :default
    ws_xml = Nokogiri::XML(sheet.to_xml_string)

    assert(ws_xml.xpath("//xmlns:worksheet/@xml:space='default'"))
  end

  def test_xml_space
    assert_equal(:preserve, @wb.xml_space)
    @wb.xml_space = :default

    assert_equal(:default, @wb.xml_space)
    assert_raises(ArgumentError) { @wb.xml_space = :none }
  end

  def test_no_autowidth
    assert(@wb.use_autowidth)
    assert_raises(ArgumentError) { @wb.use_autowidth = 0.1 }
    refute_raises { @wb.use_autowidth = false }
    assert_false(@wb.use_autowidth)
  end

  def test_is_reversed
    assert_nil(@wb.is_reversed)
    assert_raises(ArgumentError) { @wb.is_reversed = 0.1 }
    refute_raises { @wb.is_reversed = true }
    assert(@wb.use_autowidth)
  end

  def test_sheet_by_name_retrieval
    @wb.add_worksheet(name: 'foo')
    @wb.add_worksheet(name: 'bar')
    @wb.add_worksheet(name: "testin'")

    assert_equal('foo', @wb.sheet_by_name('foo').name)
    assert_equal("testin&apos;", @wb.sheet_by_name("testin'").name)
  end

  def test_worksheet_empty_name
    assert_raises(ArgumentError) { @wb.add_worksheet(name: '') }
  end

  def test_date1904
    assert_equal(Axlsx::Workbook.date1904, @wb.date1904)
    @wb.date1904 = :false

    assert_equal(Axlsx::Workbook.date1904, @wb.date1904)
    Axlsx::Workbook.date1904 = :true

    assert_equal(Axlsx::Workbook.date1904, @wb.date1904)
  end

  def test_add_defined_name
    @wb.add_defined_name 'Sheet1!1:1', name: '_xlnm.Print_Titles', hidden: true

    assert_equal(1, @wb.defined_names.size)
  end

  def test_add_view
    @wb.add_view visibility: :hidden, window_width: 800

    assert_equal(1, @wb.views.size)
  end

  def test_shared_strings
    assert_nil(@wb.use_shared_strings)
    assert_raises(ArgumentError) { @wb.use_shared_strings = 'bpb' }
    refute_raises { @wb.use_shared_strings = :true }
  end

  def test_add_worksheet
    assert_empty(@wb.worksheets, "workbook has no worksheets by default")
    ws = @wb.add_worksheet(name: "bob")

    assert_equal(1, @wb.worksheets.size, "add_worksheet adds a worksheet!")
    assert_equal(@wb.worksheets.first, ws, "the worksheet returned is the worksheet added")
    assert_equal("bob", ws.name, "name option gets passed to worksheet")
  end

  def test_insert_worksheet
    @wb.add_worksheet(name: 'A')
    @wb.add_worksheet(name: 'B')
    ws3 = @wb.insert_worksheet(0, name: 'C')

    assert_equal(ws3.name, @wb.worksheets.first.name)
  end

  def test_relationships
    # current relationship size is 1 due to style relation
    assert_equal(1, @wb.relationships.size)
    @wb.add_worksheet

    assert_equal(2, @wb.relationships.size)
    @wb.use_shared_strings = true

    assert_equal(3, @wb.relationships.size)
  end

  def test_to_xml
    schema = Nokogiri::XML::Schema(File.open(Axlsx::SML_XSD))
    doc = Nokogiri::XML(@wb.to_xml_string)
    errors = schema.validate(doc)

    assert_empty(errors)
  end

  def test_to_xml_reversed
    @wb.is_reversed = true
    @wb.add_worksheet(name: 'first')
    second = @wb.add_worksheet(name: 'second')
    doc = Nokogiri::XML(@wb.to_xml_string)

    assert_equal second.name, doc.xpath('//xmlns:workbook/xmlns:sheets/*[1]/@name').to_s
  end

  def test_range_requires_valid_sheet
    ws = @wb.add_worksheet name: 'fish'
    ws.add_row [1, 2, 3]
    ws.add_row [4, 5, 6]
    assert_raises(ArgumentError, "no sheet name part") { @wb["A1:C2"] }
    assert_equal(6, @wb['fish!A1:C2'].size)
  end

  def test_to_xml_adds_worksheet_when_worksheets_is_empty
    assert_empty(@wb.worksheets)
    @wb.to_xml_string

    assert_equal(1, @wb.worksheets.size)
  end

  def test_to_xml_string_defined_names
    @wb.add_worksheet do |sheet|
      sheet.add_row [1, "two"]
      sheet.auto_filter = "A1:B1"
    end
    doc = Nokogiri::XML(@wb.to_xml_string)

    assert_equal(doc.xpath('//xmlns:workbook/xmlns:definedNames/xmlns:definedName').inner_text, @wb.worksheets[0].auto_filter.defined_name)
  end

  def test_to_xml_string_book_views
    @wb.add_worksheet do |sheet|
      sheet.add_row [1, "two"]
    end
    @wb.add_view active_tab: 0, first_sheet: 0
    doc = Nokogiri::XML(@wb.to_xml_string)

    assert_equal(1, doc.xpath('//xmlns:workbook/xmlns:bookViews/xmlns:workbookView[@activeTab=0]').size)
  end

  def test_to_xml_uses_correct_rIds_for_pivotCache
    ws = @wb.add_worksheet
    pivot_table = ws.add_pivot_table('G5:G6', 'A1:D5')
    doc = Nokogiri::XML(@wb.to_xml_string)

    assert_equal pivot_table.cache_definition.rId, doc.xpath("//xmlns:pivotCache").first["r:id"]
  end

  def test_worksheet_name_is_intact_after_serialized_into_xml
    sheet = @wb.add_worksheet(name: '_Example')
    wb_xml = Nokogiri::XML(@wb.to_xml_string)

    assert_equal sheet.name, wb_xml.xpath('//xmlns:workbook/xmlns:sheets/*[1]/@name').to_s
  end

  def test_escape_formulas
    Axlsx.escape_formulas = false
    p = Axlsx::Package.new
    @wb = p.workbook

    assert_false @wb.escape_formulas
    assert_false @wb.add_worksheet.escape_formulas
    assert_false @wb.add_worksheet(escape_formulas: false).escape_formulas
    assert @wb.add_worksheet(escape_formulas: true).escape_formulas

    Axlsx.escape_formulas = true
    p = Axlsx::Package.new
    @wb = p.workbook

    assert @wb.escape_formulas
    assert @wb.add_worksheet.escape_formulas
    assert_false @wb.add_worksheet(escape_formulas: false).escape_formulas
    assert @wb.add_worksheet(escape_formulas: true).escape_formulas

    @wb.escape_formulas = false

    assert_false @wb.escape_formulas
    assert_false @wb.add_worksheet.escape_formulas
    assert_false @wb.add_worksheet(escape_formulas: false).escape_formulas
    assert @wb.add_worksheet(escape_formulas: true).escape_formulas

    @wb.escape_formulas = true
    p = Axlsx::Package.new
    @wb = p.workbook

    assert @wb.escape_formulas
    assert @wb.add_worksheet.escape_formulas
    assert_false @wb.add_worksheet(escape_formulas: false).escape_formulas
    assert @wb.add_worksheet(escape_formulas: true).escape_formulas
  ensure
    Axlsx.instance_variable_set(:@escape_formulas, nil)
  end
end
