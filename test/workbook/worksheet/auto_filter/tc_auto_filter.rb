# frozen_string_literal: true

require 'tc_helper'

class TestAutoFilter < Minitest::Test
  def setup
    ws = Axlsx::Package.new.workbook.add_worksheet
    3.times { |index| ws.add_row [1 * index, 2 * index, 3 * index] }
    @auto_filter = ws.auto_filter
    @auto_filter.range = 'A1:C3'
    @auto_filter.add_column 0, :filters, filter_items: [1]
  end

  def test_defined_name
    assert_equal("'Sheet1'!$A$1:$C$3", @auto_filter.defined_name)
  end

  def test_defined_name_for_missing_cell
    @auto_filter.range = 'A1:D4'

    assert_equal("'Sheet1'!$A$1:$D$4", @auto_filter.defined_name)
  end

  def test_to_xml_string
    doc = Nokogiri::XML(@auto_filter.to_xml_string)

    assert(doc.xpath("autoFilter[@ref='#{@auto_filter.range}']"))
  end

  def test_columns
    assert_kind_of Axlsx::SimpleTypedList, @auto_filter.columns
    assert_equal @auto_filter.columns.allowed_types, [Axlsx::FilterColumn]
  end

  def test_add_column
    assert_kind_of Axlsx::FilterColumn, @auto_filter.add_column(0, :filters)
  end

  def test_apply
    assert_nil @auto_filter.worksheet.rows.last.hidden
    @auto_filter.apply

    assert @auto_filter.worksheet.rows.last.hidden
  end

  def test_add_column_with_date_group_items_xml
    ws = Axlsx::Package.new.workbook.add_worksheet
    ws.add_row ['Date']
    ws.add_row [Date.new(2026, 5, 3)]
    ws.add_row [Date.new(2026, 6, 3)]
    ws.auto_filter.range = 'A1:A3'
    ws.auto_filter.add_column(0, :filters, date_group_items: [
      { date_time_grouping: :month, year: 2026, month: 5 },
      { date_time_grouping: :day,   year: 2026, month: 6, day: 3 }
    ])
    doc = Nokogiri::XML(ws.auto_filter.to_xml_string)

    assert_equal(1, doc.xpath("//filterColumn[@colId='0']").size)
    assert_equal(2, doc.xpath('//dateGroupItem').size)
    assert_equal(1, doc.xpath("//dateGroupItem[@dateTimeGrouping='month']").size)
    assert_equal(1, doc.xpath("//dateGroupItem[@dateTimeGrouping='day']").size)
  end
end
