# frozen_string_literal: true

require 'tc_helper'

class TestProtectedRange < Minitest::Test
  def setup
    @p = Axlsx::Package.new
    @ws = @p.workbook.add_worksheet { |sheet| sheet.add_row [1, 2, 3, 4, 5, 6, 7, 8, 9] }
  end

  def test_initialize_options
    refute_raises { Axlsx::ProtectedRange.new(sqref: 'A1:B1', name: "only bob") }
  end

  def test_range
    r = @ws.protect_range('A1:B1')

    assert_equal('A1:B1', r.sqref)
  end
end
