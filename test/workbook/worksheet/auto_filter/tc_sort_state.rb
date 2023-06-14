require 'tc_helper'

class TestSortState < Test::Unit::TestCase
  def setup
    ws = Axlsx::Package.new.workbook.add_worksheet
    ws.add_row ['number', 'letter', 'custom order']
    ws.add_row [1, 'B', 'high']
    ws.add_row [2, 'C', 'low']
    ws.add_row [3, 'A', 'middle']
    @auto_filter = ws.auto_filter
    @auto_filter.range = 'A1:C4'
    @auto_filter.sort_state.add_sort_condition(0)
    @auto_filter.sort_state.add_sort_condition(1, true)
    @auto_filter.sort_state.add_sort_condition(2, false, ['low', 'middle', 'high'])
    @sort_state = @auto_filter.sort_state
  end

  def test_sort_conditions
    assert @sort_state.sort_conditions.is_a?(Axlsx::SimpleTypedList)
    assert_equal @sort_state.sort_conditions.allowed_types, [Axlsx::SortCondition]
  end

  def test_add_sort_conditions
    @sort_state.add_sort_condition(0) do |condition|
      assert condition.is_a? SortCondition
    end
  end

  def test_increment_cell_value
    assert_equal @sort_state.increment_cell_value('A1'), 'A2'
  end

  def test_to_xml_string
    doc = Nokogiri::XML(@sort_state.to_xml_string)

    assert_equal(doc.xpath("sortState")[0].attribute('ref').value, 'A2:C4')
  end
end
