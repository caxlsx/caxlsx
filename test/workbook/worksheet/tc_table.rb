# frozen_string_literal: true

require 'tc_helper'

class TestTable < Test::Unit::TestCase
  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet
    40.times do
      @ws << ["aa", "aa", "aa", "aa", "aa", "aa"]
    end
  end

  def test_initialization
    assert_empty(@ws.workbook.tables)
    assert_empty(@ws.tables)
  end

  def test_table_style_info
    table = @ws.add_table('A1:D5', name: 'foo', style_info: { show_row_stripes: true, name: "TableStyleMedium25" })

    assert_equal('TableStyleMedium25', table.table_style_info.name)
    assert(table.table_style_info.show_row_stripes)
  end

  def test_add_table
    name = "test"
    table = @ws.add_table("A1:D5", name: name)

    assert_kind_of(Axlsx::Table, table, "must create a table")
    assert_equal(@ws.workbook.tables.last, table, "must be added to workbook table collection")
    assert_equal(@ws.tables.last, table, "must be added to worksheet table collection")
    assert_equal(table.name, name, "options for name are applied")
  end

  def test_pn
    @ws.add_table("A1:D5")

    assert_equal("tables/table1.xml", @ws.tables.first.pn)
  end

  def test_rId
    table = @ws.add_table("A1:D5")

    assert_equal @ws.relationships.for(table).Id, table.rId
  end

  def test_index
    @ws.add_table("A1:D5")

    assert_equal(@ws.tables.first.index, @ws.workbook.tables.index(@ws.tables.first))
  end

  def test_relationships
    assert_empty(@ws.relationships)
    @ws.add_table("A1:D5")

    assert_equal(1, @ws.relationships.size, "adding a table adds a relationship")
    @ws.add_table("F1:J5")

    assert_equal(2, @ws.relationships.size, "adding a table adds a relationship")
  end

  def test_to_xml_string
    table = @ws.add_table("A1:D5")
    schema = Nokogiri::XML::Schema(File.open(Axlsx::SML_XSD))
    doc = Nokogiri::XML(table.to_xml_string)
    errors = schema.validate(doc)

    assert_empty(errors)
  end

  def test_to_xml_string_for_special_characters
    cell = @ws.rows.first.cells.first
    cell.value = "&><'\""

    table = @ws.add_table("A1:D5")
    doc = Nokogiri::XML(table.to_xml_string)

    assert_empty(doc.errors)
  end
end
