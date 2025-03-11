# frozen_string_literal: true

require 'tc_helper'

class TestSelection < Minitest::Test
  def setup
    @options = { active_cell: 'A2', active_cell_id: 1, pane: :top_left, sqref: 'A2' }
    @selection = Axlsx::Selection.new(@options)
  end

  def test_active_cell
    assert_raises(ArgumentError) { @selection.active_cell = :active_cell }
    refute_raises { @selection.active_cell = "F5" }
    assert_equal("F5", @selection.active_cell)
  end

  def test_active_cell_id
    assert_raises(ArgumentError) { @selection.active_cell_id = "foo" }
    refute_raises { @selection.active_cell_id = 11 }
    assert_equal(11, @selection.active_cell_id)
  end

  def test_pane
    assert_raises(ArgumentError) { @selection.pane = "foo´" }
    refute_raises { @selection.pane = :bottom_right }
    assert_equal("bottomRight", @selection.pane)
  end

  def test_sqref
    assert_raises(ArgumentError) { @selection.sqref = :sqref }
    refute_raises { @selection.sqref = "G32" }
    assert_equal("G32", @selection.sqref)
  end

  def test_to_xml
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet name: "sheetview"
    @ws.sheet_view do |vs|
      vs.add_selection(:top_left, { active_cell: 'B2', sqref: 'B2' })
      vs.add_selection(:top_right, { active_cell: 'I10', sqref: 'I10' })
      vs.add_selection(:bottom_left, { active_cell: 'E55', sqref: 'E55' })
      vs.add_selection(:bottom_right, { active_cell: 'I57', sqref: 'I57' })
    end

    doc = Nokogiri::XML.parse(@ws.to_xml_string)

    assert_equal(1, doc.xpath("//xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView/xmlns:selection[@sqref='B2'][@pane='topLeft'][@activeCell='B2']").size)
    assert doc.xpath("//xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView/xmlns:selection[@sqref='B2'][@pane='topLeft'][@activeCell='B2']")
    assert_equal(1, doc.xpath("//xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView/xmlns:selection[@sqref='I10'][@pane='topRight'][@activeCell='I10']").size)
    assert doc.xpath("//xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView/xmlns:selection[@sqref='I10'][@pane='topRight'][@activeCell='I10']")
    assert_equal(1, doc.xpath("//xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView/xmlns:selection[@sqref='E55'][@pane='bottomLeft'][@activeCell='E55']").size)
    assert doc.xpath("//xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView/xmlns:selection[@sqref='E55'][@pane='bottomLeft'][@activeCell='E55']")
    assert_equal(1, doc.xpath("//xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView/xmlns:selection[@sqref='I57'][@pane='bottomRight'][@activeCell='I57']").size)
    assert doc.xpath("//xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView/xmlns:selection[@sqref='I57'][@pane='bottomRight'][@activeCell='I57']")
  end
end
