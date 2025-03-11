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

  def test_timezone
    utc = Time.utc 2012 # January 1st, 2012 at 0:00 UTC
    local = Time.parse "2012-01-01 09:00:00 +0900"

    assert_equal local, utc
    assert_equal Axlsx::DateTimeConverter.time_to_serial(local) - (local.utc_offset.to_f / 86_400), Axlsx::DateTimeConverter.time_to_serial(utc)
    Axlsx::Workbook.date1904 = true

    assert_equal Axlsx::DateTimeConverter.time_to_serial(local) - (local.utc_offset.to_f / 86_400), Axlsx::DateTimeConverter.time_to_serial(utc)
  end
end
