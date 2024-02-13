# frozen_string_literal: true

require 'tc_helper'

class TestSheetView < Minitest::Test
  def setup
    # inverse defaults for booleans
    @boolean_options = { right_to_left: true, show_formulas: true, show_outline_symbols: false,
                         show_white_space: true, tab_selected: true, default_grid_color: false, show_grid_lines: false,
                         show_row_col_headers: false, show_ruler: false, show_zeros: false, window_protection: true }
    @symbol_options = { view: :page_break_preview }
    @nil_options = { color_id: 2, top_left_cell: 'A2' }
    @int_0 = { zoom_scale_normal: 100, zoom_scale_page_layout_view: 100, zoom_scale_sheet_layout_view: 100, workbook_view_id: 2 }
    @int_100 = { zoom_scale: 10 }

    @integer_options = { color_id: 2, workbook_view_id: 2 }.merge(@int_0).merge(@int_100)
    @string_options = { top_left_cell: 'A2' }

    @options = @boolean_options.merge(@boolean_options).merge(@symbol_options).merge(@nil_options).merge(@int_0).merge(@int_100)

    @sv = Axlsx::SheetView.new(@options)
  end

  def test_initialize
    sv = Axlsx::SheetView.new

    @boolean_options.each do |key, value|
      assert_equal(!value, sv.send(key.to_sym), "initialized default #{key} should be #{!value}")
      assert_equal(value, @sv.send(key.to_sym), "initialized options #{key} should be #{value}")
    end

    @nil_options.each do |key, value|
      assert_nil(sv.send(key.to_sym), "initialized default #{key} should be nil")
      assert_equal(value, @sv.send(key.to_sym), "initialized options #{key} should be #{value}")
    end

    @int_0.each do |key, value|
      assert_equal(0, sv.send(key.to_sym), "initialized default #{key} should be 0")
      assert_equal(value, @sv.send(key.to_sym), "initialized options #{key} should be #{value}")
    end

    @int_100.each do |key, value|
      assert_equal(100, sv.send(key.to_sym), "initialized default #{key} should be 100")
      assert_equal(value, @sv.send(key.to_sym), "initialized options #{key} should be #{value}")
    end
  end

  def test_boolean_attribute_validation
    @boolean_options.each do |key, value|
      assert_raises(ArgumentError, "#{key} must be boolean") { @sv.send(:"#{key}=", 'A') }
      refute_raises { @sv.send(:"#{key}=", value) }
    end
  end

  def test_string_attribute_validation
    @string_options.each do |key, value|
      assert_raises(ArgumentError, "#{key} must be string") { @sv.send(:"#{key}=", :symbol) }
      refute_raises { @sv.send(:"#{key}=", value) }
    end
  end

  def test_symbol_attribute_validation
    @symbol_options.each do |key, value|
      assert_raises(ArgumentError, "#{key} must be symbol") { @sv.send(:"#{key}=", "foo") }
      refute_raises { @sv.send(:"#{key}=", value) }
    end
  end

  def test_integer_attribute_validation
    @integer_options.each do |key, value|
      assert_raises(ArgumentError, "#{key} must be integer") { @sv.send(:"#{key}=", "foo") }
      refute_raises { @sv.send(:"#{key}=", value) }
    end
  end

  def test_color_id
    assert_raises(ArgumentError) { @sv.color_id = "10" }
    refute_raises { @sv.color_id = 2 }
    assert_equal(2, @sv.color_id)
  end

  def test_default_grid_color
    assert_raises(ArgumentError) { @sv.default_grid_color = "foo" }
    refute_raises { @sv.default_grid_color = false }
    assert_false(@sv.default_grid_color)
  end

  def test_right_to_left
    assert_raises(ArgumentError) { @sv.right_to_left = "foo´" }
    refute_raises { @sv.right_to_left = true }
    assert(@sv.right_to_left)
  end

  def test_show_formulas
    assert_raises(ArgumentError) { @sv.show_formulas = 'foo' }
    refute_raises { @sv.show_formulas = false }
    assert_false(@sv.show_formulas)
  end

  def test_show_grid_lines
    assert_raises(ArgumentError) { @sv.show_grid_lines = "foo" }
    refute_raises { @sv.show_grid_lines = false }
    assert_false(@sv.show_grid_lines)
  end

  def test_show_outline_symbols
    assert_raises(ArgumentError) { @sv.show_outline_symbols = 'foo' }
    refute_raises { @sv.show_outline_symbols = true }
    assert(@sv.show_outline_symbols)
  end

  def test_show_row_col_headers
    assert_raises(ArgumentError) { @sv.show_row_col_headers = "foo" }
    refute_raises { @sv.show_row_col_headers = false }
    assert_false(@sv.show_row_col_headers)
  end

  def test_show_ruler
    assert_raises(ArgumentError) { @sv.show_ruler = 'foo' }
    refute_raises { @sv.show_ruler = false }
    assert_false(@sv.show_ruler)
  end

  def test_show_white_space
    assert_raises(ArgumentError) { @sv.show_white_space = 'foo' }
    refute_raises { @sv.show_white_space = false }
    assert_false(@sv.show_white_space)
  end

  def test_show_zeros
    assert_raises(ArgumentError) { @sv.show_zeros = "foo" }
    refute_raises { @sv.show_zeros = false }
    assert_false(@sv.show_zeros)
  end

  def test_tab_selected
    assert_raises(ArgumentError) { @sv.tab_selected = "foo" }
    refute_raises { @sv.tab_selected = false }
    assert_false(@sv.tab_selected)
  end

  def test_top_left_cell
    assert_raises(ArgumentError) { @sv.top_left_cell = :cell_adress }
    refute_raises { @sv.top_left_cell = "A2" }
    assert_equal("A2", @sv.top_left_cell)
  end

  def test_view
    assert_raises(ArgumentError) { @sv.view = 'view' }
    refute_raises { @sv.view = :page_break_preview }
    assert_equal(:page_break_preview, @sv.view)
  end

  def test_window_protection
    assert_raises(ArgumentError) { @sv.window_protection = "foo" }
    refute_raises { @sv.window_protection = false }
    assert_false(@sv.window_protection)
  end

  def test_workbook_view_id
    assert_raises(ArgumentError) { @sv.workbook_view_id = "1" }
    refute_raises { @sv.workbook_view_id = 1 }
    assert_equal(1, @sv.workbook_view_id)
  end

  def test_zoom_scale
    assert_raises(ArgumentError) { @sv.zoom_scale = "50" }
    refute_raises { @sv.zoom_scale = 50 }
    assert_equal(50, @sv.zoom_scale)
  end

  def test_zoom_scale_normal
    assert_raises(ArgumentError) { @sv.zoom_scale_normal = "50" }
    refute_raises { @sv.zoom_scale_normal = 50 }
    assert_equal(50, @sv.zoom_scale_normal)
  end

  def test_zoom_scale_page_layout_view
    assert_raises(ArgumentError) { @sv.zoom_scale_page_layout_view = "50" }
    refute_raises { @sv.zoom_scale_page_layout_view = 50 }
    assert_equal(50, @sv.zoom_scale_page_layout_view)
  end

  def test_zoom_scale_sheet_layout_view
    assert_raises(ArgumentError) { @sv.zoom_scale_sheet_layout_view = "50" }
    refute_raises { @sv.zoom_scale_sheet_layout_view = 50 }
    assert_equal(50, @sv.zoom_scale_sheet_layout_view)
  end

  def test_to_xml
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet name: "sheetview"
    @ws.sheet_view do |vs|
      vs.view = :page_break_preview
    end

    doc = Nokogiri::XML.parse(@ws.sheet_view.to_xml_string)

    assert_equal(1, doc.xpath("//sheetView[@tabSelected=0]").size)

    assert_equal(1, doc.xpath("//sheetView[@tabSelected=0][@showWhiteSpace=0][@showOutlineSymbols=1][@showFormulas=0]
        [@rightToLeft=0][@windowProtection=0][@showZeros=1][@showRuler=1]
        [@showRowColHeaders=1][@showGridLines=1][@defaultGridColor=1]
        [@zoomScale='100'][@workbookViewId='0'][@zoomScaleSheetLayoutView='0'][@zoomScalePageLayoutView='0']
        [@zoomScaleNormal='0'][@view='pageBreakPreview']").size)
  end

  def test_add_selection
    @sv.add_selection(:top_left, active_cell: "A1")

    assert_equal('A1', @sv.selections[:top_left].active_cell)
  end
end
