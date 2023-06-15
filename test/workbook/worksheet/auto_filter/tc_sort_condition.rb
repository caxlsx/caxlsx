require 'tc_helper'

class TestSortCondition < Test::Unit::TestCase
  def setup
    ws = Axlsx::Package.new.workbook.add_worksheet
    ws.add_row ['first', 'second', 'third']
    3.times { |index| ws.add_row [1 * index, 2 * index, 3 * index] }
    ws.auto_filter = 'A1:C4'
    @auto_filter = ws.auto_filter
    @auto_filter.sort_state.add_sort_condition(0)
    @auto_filter.sort_state.add_sort_condition(1, order: :desc)
    @auto_filter.sort_state.add_sort_condition(2, custom_list: ['low', 'middle', 'high'])
    @sort_state = @auto_filter.sort_state
    @sort_conditions = @sort_state.sort_conditions
  end

  def test_ref_to_single_column
    assert_equal @sort_conditions[0].ref_to_single_column('A2:C4',0), 'A2:A4'
  end

  def test_get_column_letter
    assert_equal @sort_conditions[0].get_column_letter(100), 'CW'
  end

  def test_to_xml_string
    doc = Nokogiri::XML(@sort_state.to_xml_string)

    assert_equal doc.xpath("sortState//sortCondition").size, 3
    assert_equal doc.xpath("sortState//sortCondition")[0].attribute('ref').value, 'A2:A4'
    assert_nil doc.xpath("sortState//sortCondition")[0].attribute('descending')
    assert_nil doc.xpath("sortState//sortCondition")[0].attribute('customList')
    assert_equal doc.xpath("sortState//sortCondition")[1].attribute('descending').value, '1'
    assert_equal doc.xpath("sortState//sortCondition")[1].attribute('ref').value, 'B2:B4'
    assert_nil doc.xpath("sortState//sortCondition")[1].attribute('customList')
    assert_equal doc.xpath("sortState//sortCondition")[2].attribute('ref').value, 'C2:C4'
    assert_equal doc.xpath("sortState//sortCondition")[2].attribute('customList').value, 'low,middle,high'
    assert_nil doc.xpath("sortState//sortCondition")[2].attribute('descending')
  end
end
