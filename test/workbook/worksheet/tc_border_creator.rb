require 'tc_helper.rb'

class TestBorderCreator < Test::Unit::TestCase
  def setup
    @p = Axlsx::Package.new
    @wb = @p.workbook
    @ws = @wb.add_worksheet
  end

  def test_defaults
    3.times do
      @ws.add_row [1,2,3]
    end

    bc = Axlsx::BorderCreator.new(@ws, @ws["A1:B2"], {})
    assert_equal bc.instance_variable_get(:@edges), :all
    assert_equal bc.instance_variable_get(:@width), :thin
    assert_equal bc.instance_variable_get(:@color), "000000"

    bc = Axlsx::BorderCreator.new(@ws, @ws["A1:B2"], [:top])
    assert_equal bc.instance_variable_get(:@edges), [:top]
    assert_equal bc.instance_variable_get(:@width), :thin
    assert_equal bc.instance_variable_get(:@color), "000000"

    bc = Axlsx::BorderCreator.new(@ws, @ws["A1:B2"], {edges: [:top], style: :thick, color: "ffffff"})
    assert_equal bc.instance_variable_get(:@edges), [:top]
    assert_equal bc.instance_variable_get(:@width), :thick
    assert_equal bc.instance_variable_get(:@color), "ffffff"
  end

  def test_draw
    5.times do
      @ws.add_row [1,2,3,4,5]
    end

    bc = Axlsx::BorderCreator.new(@ws, @ws["A1:C3"], {})
    bc.draw
    # TODO add more expectations
  end
end
