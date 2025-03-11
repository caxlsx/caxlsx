# frozen_string_literal: true

require 'tc_helper'

class TestPrintOptions < Minitest::Test
  def setup
    p = Axlsx::Package.new
    ws = p.workbook.add_worksheet name: "hmmm"
    @po = ws.print_options
  end

  def test_initialize
    assert_false(@po.grid_lines)
    assert_false(@po.headings)
    assert_false(@po.horizontal_centered)
    assert_false(@po.vertical_centered)
  end

  def test_initialize_with_options
    optioned = Axlsx::PrintOptions.new(grid_lines: true, headings: true, horizontal_centered: true, vertical_centered: true)

    assert(optioned.grid_lines)
    assert(optioned.headings)
    assert(optioned.horizontal_centered)
    assert(optioned.vertical_centered)
  end

  def test_set_all_values
    @po.set(grid_lines: true, headings: true, horizontal_centered: true, vertical_centered: true)

    assert(@po.grid_lines)
    assert(@po.headings)
    assert(@po.horizontal_centered)
    assert(@po.vertical_centered)
  end

  def test_set_some_values
    @po.set(grid_lines: true, headings: true)

    assert(@po.grid_lines)
    assert(@po.headings)
    assert_false(@po.horizontal_centered)
    assert_false(@po.vertical_centered)
  end

  def test_to_xml
    @po.set(grid_lines: true, headings: true, horizontal_centered: true, vertical_centered: true)
    doc = Nokogiri::XML.parse(@po.to_xml_string)

    assert_equal(1, doc.xpath(".//printOptions[@gridLines=1][@headings=1][@horizontalCentered=1][@verticalCentered=1]").size)
  end

  def test_grid_lines
    assert_raises(ArgumentError) { @po.grid_lines = 99 }
    refute_raises { @po.grid_lines = true }
    assert(@po.grid_lines)
  end

  def test_headings
    assert_raises(ArgumentError) { @po.headings = 99 }
    refute_raises { @po.headings = true }
    assert(@po.headings)
  end

  def test_horizontal_centered
    assert_raises(ArgumentError) { @po.horizontal_centered = 99 }
    refute_raises { @po.horizontal_centered = true }
    assert(@po.horizontal_centered)
  end

  def test_vertical_centered
    assert_raises(ArgumentError) { @po.vertical_centered = 99 }
    refute_raises { @po.vertical_centered = true }
    assert(@po.vertical_centered)
  end
end
