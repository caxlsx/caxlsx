require 'tc_helper.rb'

class TestSeriesTitle < Test::Unit::TestCase
  def setup
    @p = Axlsx::Package.new
    ws = @p.workbook.add_worksheet
    @row = ws.add_row ["one", 1, Time.now]
    @title = Axlsx::SeriesTitle.new
    @chart = ws.add_chart Axlsx::Bar3DChart
  end

  def teardown; end

  def test_initialization
    assert_equal("", @title.text)
    assert_nil(@title.cell)
  end

  def test_text
    assert_raise(ArgumentError, "text must be a string") { @title.text = 123 }
    @title.cell = @row.cells.first
    @title.text = "bob"

    assert_nil(@title.cell, "setting title with text clears the cell")
  end

  def test_cell
    assert_raise(ArgumentError, "cell must be a Cell") { @title.cell = "123" }
    @title.cell = @row.cells.first

    assert_equal("one", @title.text)
  end

  def test_to_xml_string_for_special_characters
    @chart.add_series(title: @title, data: [3, 7], labels: ['A', 'B'])

    @title.text = "&><'\""

    doc = Nokogiri::XML(@chart.to_xml_string)
    errors = doc.errors

    assert_empty(errors, "invalid xml: #{errors.map(&:to_s).join(', ')}")
  end

  def test_to_xml_string_for_special_characters_in_cell
    @chart.add_series(title: @title, data: [3, 7], labels: ['A', 'B'])

    cell = @row.cells.first
    cell.value = "&><'\""
    @title.cell = cell

    doc = Nokogiri::XML(@chart.to_xml_string)
    errors = doc.errors

    assert_empty(errors, "invalid xml: #{errors.map(&:to_s).join(', ')}")
  end
end
