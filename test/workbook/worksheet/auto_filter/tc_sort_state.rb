# frozen_string_literal: true

require 'tc_helper'

class TestSortState < Minitest::Test
  def setup
    ws = Axlsx::Package.new.workbook.add_worksheet
    ws.add_row ['first', 'second', 'third']
    3.times { |index| ws.add_row [1 * index, 2 * index, 3 * index] }
    ws.auto_filter = 'A1:C4'
    @auto_filter = ws.auto_filter
    @auto_filter.sort_state.add_sort_condition(column_index: 0)
    @sort_state = @auto_filter.sort_state
  end

  def test_sort_conditions
    assert_kind_of Axlsx::SimpleTypedList, @sort_state.sort_conditions
    assert_equal @sort_state.sort_conditions.allowed_types, [Axlsx::SortCondition]
  end

  def test_add_sort_conditions
    @sort_state.add_sort_condition(column_index: 0) do |condition|
      assert_kind_of SortCondition, condition
    end
  end

  def test_increment_cell_value
    assert_equal('A2', @sort_state.increment_cell_value('A1'))
  end

  def test_to_xml_string
    doc = Nokogiri::XML(@sort_state.to_xml_string)

    assert_equal('A2:C4', doc.xpath("sortState")[0].attribute('ref').value)
  end
end
