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
    assert_equal bc.instance_variable_get(:@edges), Axlsx::Border::EDGES
    assert_equal bc.instance_variable_get(:@width), :thin
    assert_equal bc.instance_variable_get(:@color), "000000"

    bc = Axlsx::BorderCreator.new(@ws, @ws["A1:B2"], ["top"])
    assert_equal bc.instance_variable_get(:@edges), [:top]
    assert_equal bc.instance_variable_get(:@width), :thin
    assert_equal bc.instance_variable_get(:@color), "000000"

    bc = Axlsx::BorderCreator.new(@ws, @ws["A1:B2"], :all)
    assert_equal bc.instance_variable_get(:@edges), Axlsx::Border::EDGES
    assert_equal bc.instance_variable_get(:@width), :thin
    assert_equal bc.instance_variable_get(:@color), "000000"

    bc = Axlsx::BorderCreator.new(@ws, @ws["A1:B2"], [:foo])
    assert_equal bc.instance_variable_get(:@edges), []
    assert_equal bc.instance_variable_get(:@width), :thin
    assert_equal bc.instance_variable_get(:@color), "000000"

    bc = Axlsx::BorderCreator.new(@ws, @ws["A1:B2"], {edges: ["top"], style: :thick, color: "ffffff"})
    assert_equal bc.instance_variable_get(:@edges), [:top]
    assert_equal bc.instance_variable_get(:@width), :thick
    assert_equal bc.instance_variable_get(:@color), "ffffff"
  end

  def test_draw
    5.times do
      @ws.add_row [1,2,3,4,5]
    end

    bc = Axlsx::BorderCreator.new(@ws, @ws["A1:C3"], {edges: ["top", :left], style: :thick, color: "ffffff"})

    bc.draw

    assert_equal 2, @ws.styles.borders.size

    @wb.apply_styles

    assert_equal 5, @ws.styles.borders.size

    assert_equal 2, @ws.styles.borders[2].prs.size
    assert_equal ["FFFFFFFF"], @ws.styles.borders[2].prs.map(&:color).map(&:rgb).uniq
    assert_equal [:thick], @ws.styles.borders[2].prs.map(&:style).uniq
    assert_equal [:left, :top], @ws.styles.borders[2].prs.map(&:name)


    assert_equal 1, @ws.styles.borders[3].prs.size
    assert_equal ["FFFFFFFF"], @ws.styles.borders[3].prs.map(&:color).map(&:rgb).uniq
    assert_equal [:thick], @ws.styles.borders[3].prs.map(&:style).uniq
    assert_equal [:top], @ws.styles.borders[3].prs.map(&:name)

    assert_equal 1, @ws.styles.borders[4].prs.size
    assert_equal ["FFFFFFFF"], @ws.styles.borders[4].prs.map(&:color).map(&:rgb).uniq
    assert_equal [:thick], @ws.styles.borders[4].prs.map(&:style).uniq
    assert_equal [:left], @ws.styles.borders[4].prs.map(&:name)
  end

end
