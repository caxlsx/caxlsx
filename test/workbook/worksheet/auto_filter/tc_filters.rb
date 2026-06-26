# frozen_string_literal: true

require 'tc_helper'

class TestFilters < Minitest::Test
  def setup
    @filters = Axlsx::Filters.new(filter_items: [1, 'a'],
                                  date_group_items: [{ date_time_grouping: :year, year: 2011, month: 11, day: 11, hour: 0, minute: 0, second: 0 }],
                                  blank: true)
  end

  def test_blank
    assert @filters.blank
    assert_raises(ArgumentError) { @filters.blank = :only_if_you_want_it }
    @filters.blank = true

    assert @filters.blank
  end

  def test_calendar_type
    assert_raises(ArgumentError) { @filters.calendar_type = 'monkey calendar' }
    @filters.calendar_type = 'japan'

    assert_equal('japan', @filters.calendar_type)
  end

  def test_filters_items
    assert_kind_of Array, @filters.filter_items
    assert_equal 2, @filters.filter_items.size
  end

  def test_date_group_items
    assert_kind_of Array, @filters.date_group_items
    assert_equal 1, @filters.date_group_items.size
  end

  def test_apply_is_false_for_matching_values
    keeper = Object.new
    def keeper.value
      'a'
    end

    assert_false @filters.apply(keeper)
  end

  def test_apply_is_true_for_non_matching_values
    hidden = Object.new
    def hidden.value
      'b'
    end

    assert @filters.apply(hidden)
  end

  def test_to_xml_string
    doc = Nokogiri::XML(@filters.to_xml_string)

    assert_equal(1, doc.xpath('//filters[@blank=1]').size)
  end

  def test_date_group_items_to_xml_string_month_granularity
    filters = Axlsx::Filters.new(date_group_items: [{ date_time_grouping: :month, year: 2026, month: 5 }])
    doc = Nokogiri::XML(filters.to_xml_string)

    assert_equal(1, doc.xpath("//dateGroupItem[@dateTimeGrouping='month'][@year='2026'][@month='5']").size)
    assert_equal(0, doc.xpath('//dateGroupItem[@day]').size)
  end

  def test_date_group_items_to_xml_string_day_granularity
    filters = Axlsx::Filters.new(date_group_items: [{ date_time_grouping: :day, year: 2026, month: 6, day: 3 }])
    doc = Nokogiri::XML(filters.to_xml_string)

    assert_equal(1, doc.xpath("//dateGroupItem[@dateTimeGrouping='day'][@year='2026'][@month='6'][@day='3']").size)
  end

  def test_normalize_cell_datetime_with_date
    cell = Object.new
    cell.define_singleton_method(:value) { Date.new(2026, 5, 3) }
    result = @filters.send(:normalize_cell_datetime, cell)

    assert_equal({ year: 2026, month: 5, day: 3, hour: 0, minute: 0, second: 0 }, result)
  end

  def test_normalize_cell_datetime_with_time
    t = Time.new(2026, 6, 3, 14, 30, 0)
    cell = Object.new
    cell.define_singleton_method(:value) { t }
    result = @filters.send(:normalize_cell_datetime, cell)

    assert_equal({ year: 2026, month: 6, day: 3, hour: 14, minute: 30, second: 0 }, result)
  end

  def test_normalize_cell_datetime_with_numeric_serial
    Axlsx::Workbook.date1904 = false
    cell = Object.new
    cell.define_singleton_method(:value) { 46_145 }
    result = @filters.send(:normalize_cell_datetime, cell)

    assert_equal 2026, result[:year]
    assert_equal 5,    result[:month]
    assert_equal 3,    result[:day]
  ensure
    Axlsx::Workbook.date1904 = false
  end

  def test_normalize_cell_datetime_with_nil_cell
    assert_nil @filters.send(:normalize_cell_datetime, nil)
  end

  def test_normalize_cell_datetime_with_non_date_value
    cell = Object.new
    cell.define_singleton_method(:value) { 'hello' }
    assert_nil @filters.send(:normalize_cell_datetime, cell)
  end

  def test_date_group_items_to_xml_string_multiple
    filters = Axlsx::Filters.new(date_group_items: [
      { date_time_grouping: :month, year: 2026, month: 5 },
      { date_time_grouping: :day, year: 2026, month: 6, day: 3 }
    ])
    doc = Nokogiri::XML(filters.to_xml_string)

    assert_equal(2, doc.xpath('//dateGroupItem').size)
  end
end
