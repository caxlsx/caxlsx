# frozen_string_literal: true

require 'tc_helper'

class TestOutlinePr < Minitest::Test
  def setup
    @outline_pr = Axlsx::OutlinePr.new(summary_below: false, summary_right: true, apply_styles: false)
  end

  def test_summary_below
    assert_false @outline_pr.summary_below
  end

  def test_summary_right
    assert @outline_pr.summary_right
  end

  def test_apply_styles
    assert_false @outline_pr.apply_styles
  end
end
