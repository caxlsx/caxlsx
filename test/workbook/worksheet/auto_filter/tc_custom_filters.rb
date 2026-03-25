require 'tc_helper.rb'

class TestCustomFilters < Test::Unit::TestCase

  def setup
    @p = Axlsx::Package.new
    @ws = @p.workbook.add_worksheet
    @ws.add_row %w[x amount]
    @ws.add_row ['a', 1]
    @ws.add_row ['b', 5]
    @ws.add_row ['c', 15]
    @auto_filter = @ws.auto_filter
    @auto_filter.range = 'A1:B4'
  end

  def test_add_column_custom_filters
    col = @auto_filter.add_column(
      1,
      :custom_filters,
      :custom_filter_items => [{ :operator => :greaterThan, :val => 2 }]
    )
    assert col.is_a?(Axlsx::FilterColumn)
    assert col.filter.is_a?(Axlsx::CustomFilters)
    assert_equal 1, col.filter.custom_filter_items.size
  end

  def test_to_xml_string
    cf = Axlsx::CustomFilters.new(
      :custom_filter_items => [{ :operator => :greaterThan, :val => 1 }]
    )
    xml = cf.to_xml_string
    assert_match(/<customFilters/, xml)
    assert_match(/operator="greaterThan"/, xml)
    assert_match(/val="1"/, xml)
  end

  def test_apply_or_mode_hides_non_matching_rows
    @auto_filter.add_column(
      1,
      :custom_filters,
      :custom_filter_items => [
        { :operator => :equal, :val => 1 },
        { :operator => :equal, :val => 15 }
      ]
    )
    @auto_filter.apply

    rows = @ws.rows
    assert_equal false, rows[1].hidden
    assert rows[2].hidden
    assert_equal false, rows[3].hidden
  end

  def test_apply_and_mode
    @auto_filter.add_column(
      1,
      :custom_filters,
      :and => true,
      :custom_filter_items => [
        { :operator => :greaterThan, :val => 2 },
        { :operator => :lessThan, :val => 10 }
      ]
    )
    @auto_filter.apply

    rows = @ws.rows
    assert rows[1].hidden
    assert_equal false, rows[2].hidden
    assert rows[3].hidden
  end

  def test_between_operator_logic
    cf = Axlsx::CustomFilters.new
    cf.custom_filter_items = [{ :operator => :between, :val => [3, 12] }]

    assert cf.logical_and, 'Between should force logical AND'
    assert_equal 2, cf.custom_filter_items.size

    operators = cf.custom_filter_items.map(&:operator)
    assert_includes operators, :greaterThan
    assert_includes operators, :lessThan
  end

  def test_regex_operators_xml
    cf = Axlsx::CustomFilters.new(
      :custom_filter_items => [{ :operator => :contains, :val => 'blue' }]
    )
    xml = cf.to_xml_string
    assert_match(/val="\*blue\*" /, xml)
  end

  def test_custom_filter_apply_for_remaining_operators
    cases = [
      [:lessThanOrEqual, 10, 10, true],
      [:equal, 7, 7, true],
      [:notEqual, 7, 7, false],
      [:greaterThanOrEqual, 3, 4, true],
      [:contains, 'blu', 'Blue bird', true],
      [:notContains, 'red', 'Blue bird', true],
      [:beginsWith, 'bl', 'Blue bird', true],
      [:endsWith, 'bird', 'Blue bird', true]
    ]

    cases.each do |operator, filter_value, cell_value, expected|
      filter = Axlsx::CustomFilters::CustomFilter.new(:operator => operator, :val => filter_value)
      cell = Object.new
      cell.define_singleton_method(:value) { cell_value }
      assert_equal expected, filter.apply(cell), "Expected #{operator} to return #{expected}"
    end
  end

  def test_not_blank_filter_handles_nil_filter_value
    filter = Axlsx::CustomFilters::CustomFilter.new(:operator => :notBlank, :val => nil)
    xml = filter.to_xml_string
    assert_match(/operator="notEqual"/, xml)
    assert_match(/val=" "/, xml)
  end
end
