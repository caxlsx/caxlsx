# frozen_string_literal: true

require 'tc_helper'

class TestAutoFilter < Test::Unit::TestCase
  def setup
    ws = Axlsx::Package.new.workbook.add_worksheet
    ws.add_row ['first', 'second', 'third']
    3.times { |index| ws.add_row [1 * index, 2 * index, 3 * index] }
    ws.auto_filter = 'A1:C4'
    @auto_filter = ws.auto_filter
    @auto_filter.add_column 0, :filters, filter_items: [1]
    @auto_filter.sort_state.add_sort_condition 0, order: :desc
  end

  def test_defined_name
    assert_equal("'Sheet1'!$A$1:$C$4", @auto_filter.defined_name)
  end

  def test_to_xml_string
    doc = Nokogiri::XML(@auto_filter.worksheet.to_xml_string)

    assert_equal(doc.at_xpath('//xmlns:autoFilter')['ref'], 'A1:C4')
    assert_equal(doc.at_xpath('//xmlns:autoFilter').children.size, 2)
  end

  def test_columns
    assert @auto_filter.columns.is_a?(Axlsx::SimpleTypedList)
    assert_equal @auto_filter.columns.allowed_types, [Axlsx::FilterColumn]
  end

  def test_add_column
    @auto_filter.add_column(0, :filters) do |column|
      assert column.is_a? FilterColumn
    end
  end

  def test_apply
    assert_nil @auto_filter.worksheet.rows.last.hidden
    # assert_equal @auto_filter.worksheet.rows.last.cells, ''
    assert @auto_filter.worksheet.rows.last.cells.none? { |cell| cell.value == 0 }

    @auto_filter.apply

    assert @auto_filter.worksheet.rows.last.hidden
    assert_equal @auto_filter.worksheet.rows.last.cells, ''
    # assert @auto_filter.worksheet.rows.last.cells.all? { |cell| cell.value == 0 }
  end
end
