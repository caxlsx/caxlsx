# frozen_string_literal: true

require 'tc_helper'

class TestDateTimeConverter < Minitest::Test
  def setup
    @margin_of_error = 0.000_001
  end

  def test_date_to_serial_1900
    Axlsx::Workbook.date1904 = false
    { # examples taken straight from the spec
      "1893-08-05" => -2338.0,
      "1900-01-01" => 2.0,
      "1910-02-03" => 3687.0,
      "2006-02-01" => 38_749.0,
      "9999-12-31" => 2_958_465.0
    }.each do |date_string, expected|
      serial = Axlsx::DateTimeConverter.date_to_serial Date.parse(date_string)

      assert_equal expected, serial
    end
  end

  def test_date_to_serial_1904
    Axlsx::Workbook.date1904 = true
    { # examples taken straight from the spec
      "1893-08-05" => -3800.0,
      "1904-01-01" => 0.0,
      "1910-02-03" => 2225.0,
      "2006-02-01" => 37_287.0,
      "9999-12-31" => 2_957_003.0
    }.each do |date_string, expected|
      serial = Axlsx::DateTimeConverter.date_to_serial Date.parse(date_string)

      assert_equal expected, serial
    end
  end

  def test_time_to_serial_1900
    Axlsx::Workbook.date1904 = false
    { # examples taken straight from the spec
      "1893-08-05T00:00:01Z" => -2337.999989,
      "1899-12-28T18:00:00Z" => -1.25,
      "1910-02-03T10:05:54Z" => 3687.4207639,
      "1900-01-01T12:00:00Z" => 2.5, # wrongly indicated as 1.5 in the spec!
      "9999-12-31T23:59:59Z" => 2_958_465.9999884
    }.each do |time_string, expected|
      serial = Axlsx::DateTimeConverter.time_to_serial Time.parse(time_string)

      assert_in_delta expected, serial, @margin_of_error
    end
  end

  def test_time_to_serial_1904
    Axlsx::Workbook.date1904 = true

    { # examples taken straight from the spec
      "1893-08-05T00:00:01Z" => -3799.999989,
      "1910-02-03T10:05:54Z" => 2225.4207639,
      "1904-01-01T12:00:00Z" => 0.5000000,
      "9999-12-31T23:59:59Z" => 2_957_003.9999884
    }.each do |time_string, expected|
      serial = Axlsx::DateTimeConverter.time_to_serial Time.parse(time_string)

      assert_in_delta expected, serial, @margin_of_error
    end
  end

  def test_date_from_serial_1900
    Axlsx::Workbook.date1904 = false

    {
      2.0 => Date.new(1900, 1, 1),
      38_749.0 => Date.new(2006, 2, 1),
      46_145.0 => Date.new(2026, 5, 3)
    }.each do |serial, expected|
      assert_equal expected, Axlsx::DateTimeConverter.date_from_serial(serial)
    end
  end

  def test_date_from_serial_1904
    Axlsx::Workbook.date1904 = true

    assert_equal Date.new(1904, 1, 1), Axlsx::DateTimeConverter.date_from_serial(0)
  ensure
    Axlsx::Workbook.date1904 = false
  end

  def test_datetime_components_from_serial_date_only
    Axlsx::Workbook.date1904 = false
    result = Axlsx::DateTimeConverter.datetime_components_from_serial(46_145)

    assert_equal({ year: 2026, month: 5, day: 3, hour: 0, minute: 0, second: 0 }, result)
  ensure
    Axlsx::Workbook.date1904 = false
  end

  def test_datetime_components_from_serial_with_time
    Axlsx::Workbook.date1904 = false
    serial = 46_176 + (((12 * 3600) + (30 * 60)).to_f / 86_400)  # 2026-06-03 12:30:00
    result = Axlsx::DateTimeConverter.datetime_components_from_serial(serial)

    assert_equal 2026, result[:year]
    assert_equal 6,    result[:month]
    assert_equal 3,    result[:day]
    assert_equal 12,   result[:hour]
    assert_equal 30,   result[:minute]
    assert_equal 0,    result[:second]
  ensure
    Axlsx::Workbook.date1904 = false
  end

  def test_timezone
    utc = Time.utc 2012 # January 1st, 2012 at 0:00 UTC
    local = Time.parse "2012-01-01 09:00:00 +0900"

    assert_equal local, utc
    assert_equal Axlsx::DateTimeConverter.time_to_serial(local) - (local.utc_offset.to_f / 86_400), Axlsx::DateTimeConverter.time_to_serial(utc)
    Axlsx::Workbook.date1904 = true

    assert_equal Axlsx::DateTimeConverter.time_to_serial(local) - (local.utc_offset.to_f / 86_400), Axlsx::DateTimeConverter.time_to_serial(utc)
  end
end
