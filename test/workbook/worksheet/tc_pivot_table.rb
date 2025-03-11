# frozen_string_literal: true

require 'tc_helper'

def shared_test_pivot_table_xml_validity(pivot_table)
  schema = Nokogiri::XML::Schema(File.open(Axlsx::SML_XSD))
  doc = Nokogiri::XML(pivot_table.to_xml_string)
  errors = schema.validate(doc)

  assert_empty(errors)
end

class TestPivotTable < Minitest::Test
  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet

    @ws << ["Year", "Month", "Region", "Type", "Sales"]
    @ws << [2012, "Nov", "East", "Soda", "12345"]
  end

  def test_initialization
    assert_empty(@ws.workbook.pivot_tables)
    assert_empty(@ws.pivot_tables)
  end

  def test_add_pivot_table
    pivot_table = @ws.add_pivot_table('G5:G6', 'A1:D5')

    assert_equal('G5:G6', pivot_table.ref, 'ref assigned from first parameter')
    assert_equal('A1:D5', pivot_table.range, 'range assigned from second parameter')
    assert_equal('PivotTable1', pivot_table.name, 'name automatically generated')
    assert_kind_of(Axlsx::PivotTable, pivot_table, "must create a pivot table")
    assert_equal(@ws.workbook.pivot_tables.last, pivot_table, "must be added to workbook pivot tables collection")
    assert_equal(@ws.pivot_tables.last, pivot_table, "must be added to worksheet pivot tables collection")
  end

  def test_set_pivot_table_data_sheet
    pivot_table = @ws.add_pivot_table('G5:G6', 'A1:D5')
    data_sheet = @ws.clone
    data_sheet.name = "Pivot Table Data Source"

    assert_equal(pivot_table.data_sheet.name, @ws.name, "must default to the same sheet the pivot table is added to")
    pivot_table.data_sheet = data_sheet

    assert_equal(pivot_table.data_sheet.name, data_sheet.name, "data sheet assigned to pivot table")
  end

  def test_add_pivot_table_with_config
    pivot_table = @ws.add_pivot_table('G5:G6', 'A1:E5') do |pt|
      pt.rows = ['Year', 'Month']
      pt.columns = ['Type']
      pt.data = ['Sales']
      pt.pages = ['Region']
    end

    assert_equal(['Year', 'Month'], pivot_table.rows)
    assert_equal(['Type'], pivot_table.columns)
    assert_equal([{ ref: "Sales" }], pivot_table.data)
    assert_equal(['Region'], pivot_table.pages)
    shared_test_pivot_table_xml_validity(pivot_table)
  end

  def test_add_pivot_table_with_options_on_data_field
    pivot_table = @ws.add_pivot_table('G5:G6', 'A1:D5') do |pt|
      pt.data = [{ ref: "Sales", subtotal: 'average' }]
    end

    assert_equal([{ ref: "Sales", subtotal: 'average' }], pivot_table.data)
  end

  def test_add_pivot_table_with_style_info
    style_info_data = { name: "PivotStyleMedium9", showRowHeaders: "1", showLastColumn: "0" }
    pivot_table = @ws.add_pivot_table('G5:G6', 'A1:E5', { style_info: style_info_data }) do |pt|
      pt.rows = ['Year', 'Month']
      pt.columns = ['Type']
      pt.data = ['Sales']
      pt.pages = ['Region']
    end

    assert_equal(style_info_data, pivot_table.style_info)
    shared_test_pivot_table_xml_validity(pivot_table)
  end

  def test_add_pivot_table_with_row_without_subtotals
    pivot_table = @ws.add_pivot_table('G5:G6', 'A1:D5', { no_subtotals_on_headers: ['Year'] }) do |pt|
      pt.data = ['Sales']
      pt.rows = ['Year', 'Month']
    end

    assert_equal(['Year'], pivot_table.no_subtotals_on_headers)
  end

  def test_add_pivot_table_with_months_sorted
    pivot_table = @ws.add_pivot_table('G5:G6', 'A1:E5', { sort_on_headers: ['Month'] }) do |pt|
      pt.data = ['Sales']
      pt.rows = ['Year', 'Month']
    end

    assert_equal({ 'Month' => :ascending }, pivot_table.sort_on_headers)

    pivot_table.sort_on_headers = { 'Month' => :descending }

    assert_equal({ 'Month' => :descending }, pivot_table.sort_on_headers)

    shared_test_pivot_table_xml_validity(pivot_table)
  end

  def test_grand_totals
    pivot_table = @ws.add_pivot_table('G5:G6', 'A1:E5', { grand_totals: :row_only }) do |pt|
      pt.data = ['Sales']
      pt.rows = ['Year', 'Month']
    end

    assert_equal(:row_only, pivot_table.grand_totals)

    shared_test_pivot_table_xml_validity(pivot_table)
  end

  def test_header_indices
    pivot_table = @ws.add_pivot_table('G5:G6', 'A1:E5')

    assert_equal(0,   pivot_table.header_index_of('Year'))
    assert_equal(1,   pivot_table.header_index_of('Month'))
    assert_equal(2,   pivot_table.header_index_of('Region'))
    assert_equal(3,   pivot_table.header_index_of('Type'))
    assert_equal(4,   pivot_table.header_index_of('Sales'))
    assert_nil(pivot_table.header_index_of('Missing'))
    assert_equal(%w(A1 B1 C1 D1 E1), pivot_table.header_cell_refs)
  end

  def test_pn
    @ws.add_pivot_table('G5:G6', 'A1:D5')

    assert_equal("pivotTables/pivotTable1.xml", @ws.pivot_tables.first.pn)
  end

  def test_index
    @ws.add_pivot_table('G5:G6', 'A1:D5')

    assert_equal(@ws.pivot_tables.first.index, @ws.workbook.pivot_tables.index(@ws.pivot_tables.first))
  end

  def test_relationships
    assert_empty(@ws.relationships)
    @ws.add_pivot_table('G5:G6', 'A1:D5')

    assert_equal(1, @ws.relationships.size, "adding a pivot table adds a relationship")
    @ws.add_pivot_table('G10:G11', 'A1:D5')

    assert_equal(2, @ws.relationships.size, "adding a pivot table adds a relationship")
  end

  def test_rels_pn
    @ws.add_pivot_table('G5:G6', 'A1:D5')

    assert_equal("pivotTables/_rels/pivotTable1.xml.rels", @ws.pivot_tables.first.rels_pn)
  end

  def test_to_xml_string
    pivot_table = @ws.add_pivot_table('G5:G6', 'A1:E5', { no_subtotals_on_headers: ['Year'] }) do |pt|
      pt.rows = ['Year', 'Month']
      pt.columns = ['Type']
      pt.data = ['Sales']
      pt.pages = ['Region']
    end
    shared_test_pivot_table_xml_validity(pivot_table)
  end

  def test_to_xml_string_with_options_on_data_field
    pivot_table = @ws.add_pivot_table('G5:G6', 'A1:E5') do |pt|
      pt.data = [{ ref: "Sales", subtotal: 'average' }]
    end
    shared_test_pivot_table_xml_validity(pivot_table)
  end

  def test_add_pivot_table_with_format_options_on_data_field
    pivot_table = @ws.add_pivot_table('G5:G6', 'A1:E5') do |pt|
      pt.data = [{ ref: "Sales", subtotal: 'sum', num_fmt: 4 }]
    end
    doc = Nokogiri::XML(pivot_table.to_xml_string)

    assert_equal('4', doc.at_css('dataFields dataField')['numFmtId'], 'adding format options to pivot_table')
  end

  def test_pivot_table_with_more_than_one_data_row
    # https://github.com/caxlsx/caxlsx/issues/110

    pivot_table = @ws.add_pivot_table('G5:G6', 'A1:E5') do |pt|
      pt.rows = ["Date", "Name"]
      pt.data = [
        { ref: "Gross amount", num_fmt: 2 },
        { ref: "Net amount", num_fmt: 2 },
        { ref: "Margin", num_fmt: 2 }
      ]
    end

    xml = pivot_table.to_xml_string

    refute_includes(xml, 'dataOnRows')
    assert_includes(xml, 'colFields')
    assert_includes(xml, 'colItems')

    doc = Nokogiri::XML(pivot_table.to_xml_string)

    assert_equal('1', doc.at_css('colFields')['count'])
    assert_equal('-2', doc.at_css('colFields field')['x'])

    assert_equal('3', doc.at_css('colItems')['count'])
    assert_equal(3, doc.at_css('colItems').children.size)
    assert_nil(doc.at_css('colItems i')['x'])
    assert_equal('1', doc.at_css('colItems i[i=1] x')['v'])
    assert_equal('2', doc.at_css('colItems i[i=2] x')['v'])
  end

  def test_pivot_table_with_only_one_data_row
    # https://github.com/caxlsx/caxlsx/issues/110

    pivot_table = @ws.add_pivot_table('G5:G6', 'A1:E5') do |pt|
      pt.rows = ["Date", "Name"]
      pt.data = [
        { ref: "Gross amount", num_fmt: 2 }
      ]
    end

    xml = pivot_table.to_xml_string

    assert_includes(xml, 'dataOnRows')
    assert_includes(xml, 'colItems')

    refute_includes(xml, 'colFields')
  end
end
