# frozen_string_literal: true

require 'tc_helper'

class TestPageSetUpPr < Minitest::Test
  def setup
    @page_setup_pr = Axlsx::PageSetUpPr.new(fit_to_page: true, auto_page_breaks: true)
  end

  def test_fit_to_page
    assert @page_setup_pr.fit_to_page
  end

  def test_auto_page_breaks
    assert @page_setup_pr.auto_page_breaks
  end
end
