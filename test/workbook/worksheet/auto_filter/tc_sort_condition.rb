# frozen_string_literal: true

require 'tc_helper'

class TestSortCondition < Minitest::Test
  def setup
    ws = Axlsx::Package.new.workbook.add_worksheet
    ws.add_row ['first', 'second', 'third']
    3.times { |index| ws.add_row [1 * index, 2 * index, 3 * index] }
    ws.auto_filter = 'A1:C4'
    @auto_filter = ws.auto_filter
    @auto_filter.sort_state.add_sort_condition(column_index: 0)
    @auto_filter.sort_state.add_sort_condition(column_index: 1, order: :desc)
    @auto_filter.sort_state.add_sort_condition(column_index: 2, custom_list: ['low', 'middle', 'high'])
    @sort_state = @auto_filter.sort_state
    @sort_conditions = @sort_state.sort_conditions
  end

  def test_ref_to_single_column
    assert_equal('A2:A4', @sort_conditions[0].ref_to_single_column('A2:C4', 0))
  end

  def test_to_xml_string
    doc = Nokogiri::XML(@sort_state.to_xml_string)

    assert_equal(3, doc.xpath("sortState//sortCondition").size)
    assert_equal('A2:A4', doc.xpath("sortState//sortCondition")[0].attribute('ref').value)
    assert_nil doc.xpath("sortState//sortCondition")[0].attribute('descending')
    assert_nil doc.xpath("sortState//sortCondition")[0].attribute('customList')
    assert_equal('1', doc.xpath("sortState//sortCondition")[1].attribute('descending').value)
    assert_equal('B2:B4', doc.xpath("sortState//sortCondition")[1].attribute('ref').value)
    assert_nil doc.xpath("sortState//sortCondition")[1].attribute('customList')
    assert_equal('C2:C4', doc.xpath("sortState//sortCondition")[2].attribute('ref').value)
    assert_equal('low,middle,high', doc.xpath("sortState//sortCondition")[2].attribute('customList').value)
    assert_nil doc.xpath("sortState//sortCondition")[2].attribute('descending')
  end
end
