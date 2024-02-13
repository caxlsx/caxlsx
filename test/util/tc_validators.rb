# frozen_string_literal: true

require 'tc_helper'

class TestValidators < Minitest::Test
  def setup; end

  def teardown; end

  def test_validators
    # unsigned_int
    refute_raises { Axlsx.validate_unsigned_int 1 }
    refute_raises { Axlsx.validate_unsigned_int(+1) }
    assert_raises(ArgumentError) { Axlsx.validate_unsigned_int(-1) }
    assert_raises(ArgumentError) { Axlsx.validate_unsigned_int('1') }

    # int
    refute_raises { Axlsx.validate_int(1) }
    refute_raises { Axlsx.validate_int(-1) }
    assert_raises(ArgumentError) { Axlsx.validate_int('a') }
    assert_raises(ArgumentError) { Axlsx.validate_int(Array) }

    # boolean (as 0 or 1, :true, :false, true, false, or "true," "false")
    [0, 1, :true, :false, true, false, "true", "false"].each do |v|
      refute_raises { Axlsx.validate_boolean v }
    end
    assert_raises(ArgumentError) { Axlsx.validate_boolean 2 }

    # string
    refute_raises { Axlsx.validate_string "1" }
    assert_raises(ArgumentError) { Axlsx.validate_string 2 }
    assert_raises(ArgumentError) { Axlsx.validate_string false }

    # float
    refute_raises { Axlsx.validate_float 1.0 }
    assert_raises(ArgumentError) { Axlsx.validate_float 2 }
    assert_raises(ArgumentError) { Axlsx.validate_float false }

    # pattern_type
    refute_raises { Axlsx.validate_pattern_type :none }
    assert_raises(ArgumentError) { Axlsx.validate_pattern_type "none" }
    assert_raises(ArgumentError) { Axlsx.validate_pattern_type "crazy_pattern" }
    assert_raises(ArgumentError) { Axlsx.validate_pattern_type false }

    # gradient_type
    refute_raises { Axlsx.validate_gradient_type :path }
    assert_raises(ArgumentError) { Axlsx.validate_gradient_type nil }
    assert_raises(ArgumentError) { Axlsx.validate_gradient_type "fractal" }
    assert_raises(ArgumentError) { Axlsx.validate_gradient_type false }

    # horizontal alignment
    refute_raises { Axlsx.validate_horizontal_alignment :general }
    assert_raises(ArgumentError) { Axlsx.validate_horizontal_alignment nil }
    assert_raises(ArgumentError) { Axlsx.validate_horizontal_alignment "wavy" }
    assert_raises(ArgumentError) { Axlsx.validate_horizontal_alignment false }

    # vertical alignment
    refute_raises { Axlsx.validate_vertical_alignment :top }
    assert_raises(ArgumentError) { Axlsx.validate_vertical_alignment nil }
    assert_raises(ArgumentError) { Axlsx.validate_vertical_alignment "dynamic" }
    assert_raises(ArgumentError) { Axlsx.validate_vertical_alignment false }

    # contentType
    refute_raises { Axlsx.validate_content_type Axlsx::WORKBOOK_CT }
    assert_raises(ArgumentError) { Axlsx.validate_content_type nil }
    assert_raises(ArgumentError) { Axlsx.validate_content_type "http://some.url" }
    assert_raises(ArgumentError) { Axlsx.validate_content_type false }

    # relationshipType
    refute_raises { Axlsx.validate_relationship_type Axlsx::WORKBOOK_R }
    assert_raises(ArgumentError) { Axlsx.validate_relationship_type nil }
    assert_raises(ArgumentError) { Axlsx.validate_relationship_type "http://some.url" }
    assert_raises(ArgumentError) { Axlsx.validate_relationship_type false }

    # number_with_unit
    refute_raises { Axlsx.validate_number_with_unit "210mm" }
    refute_raises { Axlsx.validate_number_with_unit "8.5in" }
    refute_raises { Axlsx.validate_number_with_unit "29.7cm" }
    refute_raises { Axlsx.validate_number_with_unit "120pt" }
    refute_raises { Axlsx.validate_number_with_unit "0pc" }
    refute_raises { Axlsx.validate_number_with_unit "12.34pi" }
    assert_raises(ArgumentError) { Axlsx.validate_number_with_unit nil }
    assert_raises(ArgumentError) { Axlsx.validate_number_with_unit "210" }
    assert_raises(ArgumentError) { Axlsx.validate_number_with_unit 210 }
    assert_raises(ArgumentError) { Axlsx.validate_number_with_unit "mm" }
    assert_raises(ArgumentError) { Axlsx.validate_number_with_unit "-29cm" }

    # scale_10_400
    refute_raises { Axlsx.validate_scale_10_400 10 }
    refute_raises { Axlsx.validate_scale_10_400 100 }
    refute_raises { Axlsx.validate_scale_10_400 400 }
    assert_raises(ArgumentError) { Axlsx.validate_scale_10_400 9 }
    assert_raises(ArgumentError) { Axlsx.validate_scale_10_400 10.0 }
    assert_raises(ArgumentError) { Axlsx.validate_scale_10_400 400.1 }
    assert_raises(ArgumentError) { Axlsx.validate_scale_10_400 "99" }

    # scale_0_10_400
    refute_raises { Axlsx.validate_scale_0_10_400 0 }
    refute_raises { Axlsx.validate_scale_0_10_400 10 }
    refute_raises { Axlsx.validate_scale_0_10_400 100 }
    refute_raises { Axlsx.validate_scale_0_10_400 400 }
    assert_raises(ArgumentError) { Axlsx.validate_scale_0_10_400 9 }
    assert_raises(ArgumentError) { Axlsx.validate_scale_0_10_400 10.0 }
    assert_raises(ArgumentError) { Axlsx.validate_scale_0_10_400 400.1 }
    assert_raises(ArgumentError) { Axlsx.validate_scale_0_10_400 "99" }

    # page_orientation
    refute_raises { Axlsx.validate_page_orientation :default }
    refute_raises { Axlsx.validate_page_orientation :landscape }
    refute_raises { Axlsx.validate_page_orientation :portrait }
    assert_raises(ArgumentError) { Axlsx.validate_page_orientation nil }
    assert_raises(ArgumentError) { Axlsx.validate_page_orientation 1 }
    assert_raises(ArgumentError) { Axlsx.validate_page_orientation "landscape" }

    # data_validation_error_style
    [:information, :stop, :warning].each do |sym|
      refute_raises { Axlsx.validate_data_validation_error_style sym }
    end
    assert_raises(ArgumentError) { Axlsx.validate_data_validation_error_style :other_symbol }
    assert_raises(ArgumentError) { Axlsx.validate_data_validation_error_style 'warning' }
    assert_raises(ArgumentError) { Axlsx.validate_data_validation_error_style 0 }

    # data_validation_operator
    [:lessThan, :lessThanOrEqual, :equal, :notEqual, :greaterThanOrEqual, :greaterThan, :between, :notBetween].each do |sym|
      refute_raises { Axlsx.validate_data_validation_operator sym }
    end
    assert_raises(ArgumentError) { Axlsx.validate_data_validation_error_style :other_symbol }
    assert_raises(ArgumentError) { Axlsx.validate_data_validation_error_style 'lessThan' }
    assert_raises(ArgumentError) { Axlsx.validate_data_validation_error_style 0 }

    # data_validation_type
    [:custom, :data, :decimal, :list, :none, :textLength, :date, :time, :whole].each do |sym|
      refute_raises { Axlsx.validate_data_validation_type sym }
    end
    assert_raises(ArgumentError) { Axlsx.validate_data_validation_error_style :other_symbol }
    assert_raises(ArgumentError) { Axlsx.validate_data_validation_error_style 'decimal' }
    assert_raises(ArgumentError) { Axlsx.validate_data_validation_error_style 0 }

    # sheet_view_type
    [:normal, :page_break_preview, :page_layout].each do |sym|
      refute_raises { Axlsx.validate_sheet_view_type sym }
    end
    assert_raises(ArgumentError) { Axlsx.validate_data_validation_error_style :other_symbol }
    assert_raises(ArgumentError) { Axlsx.validate_data_validation_error_style 'page_layout' }
    assert_raises(ArgumentError) { Axlsx.validate_data_validation_error_style 0 }

    # active_pane_type
    [:bottom_left, :bottom_right, :top_left, :top_right].each do |sym|
      refute_raises { Axlsx.validate_pane_type sym }
    end
    assert_raises(ArgumentError) { Axlsx.validate_pane_type :other_symbol }
    assert_raises(ArgumentError) { Axlsx.validate_pane_type 'bottom_left' }
    assert_raises(ArgumentError) { Axlsx.validate_pane_type 0 }

    # split_state_type
    [:frozen, :frozen_split, :split].each do |sym|
      refute_raises { Axlsx.validate_split_state_type sym }
    end
    assert_raises(ArgumentError) { Axlsx.validate_split_state_type :other_symbol }
    assert_raises(ArgumentError) { Axlsx.validate_split_state_type 'frozen_split' }
    assert_raises(ArgumentError) { Axlsx.validate_split_state_type 0 }

    # display_blanks_as
    [:gap, :span, :zero].each do |sym|
      refute_raises { Axlsx.validate_display_blanks_as sym }
    end
    assert_raises(ArgumentError) { Axlsx.validate_display_blanks_as :other_symbol }
    assert_raises(ArgumentError) { Axlsx.validate_display_blanks_as 'other_blank' }
    assert_raises(ArgumentError) { Axlsx.validate_display_blanks_as 0 }

    # view_visibility
    [:visible, :hidden, :very_hidden].each do |sym|
      refute_raises { Axlsx.validate_view_visibility sym }
    end
    assert_raises(ArgumentError) { Axlsx.validate_view_visibility :other_symbol }
    assert_raises(ArgumentError) { Axlsx.validate_view_visibility 'other_visibility' }
    assert_raises(ArgumentError) { Axlsx.validate_view_visibility 0 }

    # marker_symbol
    [:default, :circle, :dash, :diamond, :dot, :picture, :plus, :square, :star, :triangle, :x].each do |sym|
      refute_raises { Axlsx.validate_marker_symbol sym }
    end
    assert_raises(ArgumentError) { Axlsx.validate_marker_symbol :other_symbol }
    assert_raises(ArgumentError) { Axlsx.validate_marker_symbol 'other_marker' }
    assert_raises(ArgumentError) { Axlsx.validate_marker_symbol 0 }
  end

  def test_validate_integerish
    assert_raises(ArgumentError) { Axlsx.validate_integerish Axlsx }
    [1, 1.4, "a"].each { |test_value| refute_raises { Axlsx.validate_integerish test_value } }
  end

  def test_validate_family
    assert_raises(ArgumentError) { Axlsx.validate_family 0 }
    (1..5).each do |item|
      refute_raises { Axlsx.validate_family item }
    end
  end

  def test_validate_u
    assert_raises(ArgumentError) { Axlsx.validate_cell_u :hoge }
    [:none, :single, :double, :singleAccounting, :doubleAccounting].each do |sym|
      refute_raises { Axlsx.validate_cell_u sym }
    end
  end

  def test_range_validation
    # exclusive
    assert_raises(ArgumentError) { Axlsx::RangeValidator.validate('foo', 1, 10, 10, false) }
    # inclusive by default
    refute_raises { Axlsx::RangeValidator.validate('foo', 1, 10, 10) }
  end
end
