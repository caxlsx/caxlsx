# encoding: UTF-8
require 'tc_helper.rb'

class TestAxlsxStyler < Test::Unit::TestCase

  def setup
    FileUtils.mkdir_p("tmp/") ### TODO: remove
  end

  def test_to_stream_automatically_performs_apply_styles
    p = Axlsx::Package.new
    wb = p.workbook

    filename = 'to_stream_automatically_performs_apply_styles'
    assert_nil wb.styles_applied
    wb.add_worksheet do |sheet|
      sheet.add_row ['A1', 'B1']
      sheet.add_style 'A1:B1', b: true
    end
    File.open("tmp/#{filename}.xlsx", 'wb') do |f|
      f.write p.to_stream.read
    end
    assert_equal 1, wb.styles.style_index.count
  end

  def test_serialize_automatically_performs_apply_styles
    p = Axlsx::Package.new
    wb = p.workbook

    filename = 'without_apply_styles_serialize'
    assert_nil wb.styles_applied
    wb.add_worksheet do |sheet|
      sheet.add_row ['A1', 'B1']
      sheet.add_style 'A1:B1', b: true
    end
    p.serialize("tmp/#{filename}.xlsx")
    assert_equal true, wb.styles_applied
    assert_equal 1, wb.styles.style_index.count
  end

  def test_merge_styles_1
    p = Axlsx::Package.new
    wb = p.workbook

    filename = 'merge_styles_1'
    bold = wb.styles.add_style b: true

    wb.add_worksheet do |sheet|
      sheet.add_row
      sheet.add_row ['', '1', '2', '3'], style: [nil, bold]
      sheet.add_row ['', '4', '5', '6'], style: bold
      sheet.add_row ['', '7', '8', '9']
      sheet.add_style 'B2:D4', b: true
      sheet.add_border 'B2:D4', { style: :thin, color: '000000' }
    end
    wb.apply_styles
    assert_equal 9, wb.styles.style_index.count
    p.serialize("tmp/#{filename}.xlsx")
    assert_equal true, wb.styles_applied
  end

  def test_merge_styles_2
    p = Axlsx::Package.new
    wb = p.workbook

    filename = 'merge_styles_2'
    bold = wb.styles.add_style b: true

    wb.add_worksheet do |sheet|
      sheet.add_row ['A1', 'B1'], style: [nil, bold]
      sheet.add_row ['A2', 'B2'], style: bold
      sheet.add_row ['A3', 'B3']
      sheet.add_style 'A1:A2', i: true
    end
    wb.apply_styles
    assert_equal 3, wb.styles.style_index.count
    p.serialize("tmp/#{filename}.xlsx")
    assert_equal true, wb.styles_applied
  end

  def test_merge_styles_3
    p = Axlsx::Package.new
    wb = p.workbook

    filename = 'merge_styles_3'
    bold = wb.styles.add_style b: true

    wb.add_worksheet do |sheet|
      sheet.add_row ['A1', 'B1'], style: [nil, bold]
      sheet.add_row ['A2', 'B2']
      sheet.add_style 'B1:B2', bg_color: 'FF0000'
    end
    wb.apply_styles
    assert_equal 3, wb.styles.style_index.count
    p.serialize("tmp/#{filename}.xlsx")
    assert_equal true, wb.styles_applied
  end

  def test_table_with_borders
    p = Axlsx::Package.new
    wb = p.workbook

    filename = 'borders_test'
    wb.add_worksheet do |sheet|
      sheet.add_row
      sheet.add_row ['', 'Product', 'Category',  'Price']
      sheet.add_row ['', 'Butter', 'Dairy',      4.99]
      sheet.add_row ['', 'Bread', 'Baked Goods', 3.45]
      sheet.add_row ['', 'Broccoli', 'Produce',  2.99]
      sheet.add_row ['', 'Pizza', 'Frozen Foods',  4.99]
      sheet.column_widths 5, 20, 20, 20

      sheet.add_style 'B2:D2', b: true
      sheet.add_style 'B2:B6', b: true
      sheet.add_style 'B2:D2', bg_color: '95AFBA'
      sheet.add_style 'B3:D6', bg_color: 'E2F89C'
      sheet.add_style 'D3:D6', alignment: { horizontal: :left }
      sheet.add_border 'B2:D6'
      sheet.add_border 'B3:D3', [:top]
      sheet.add_border 'B3:D3', edges: [:bottom], style: :medium
      sheet.add_border 'B3:D3', edges: [:bottom], style: :medium, color: '32f332'
    end
    p.serialize("tmp/#{filename}.xlsx")
    assert_equal true, wb.styles_applied
    assert_equal 12, wb.styles.style_index.count
    assert_equal 12 + 2, wb.styles.style_index.keys.max
  end

  def test_duplicate_borders
    p = Axlsx::Package.new
    wb = p.workbook

    filename = 'duplicate_borders_test'
    wb.add_worksheet do |sheet|
      sheet.add_row
      sheet.add_row ['', 'B2', 'C2', 'D2']
      sheet.add_row ['', 'B3', 'C3', 'D3']
      sheet.add_row ['', 'B4', 'C4', 'D4']

      sheet.add_border 'B2:D4'
      sheet.add_border 'B2:D4'
    end
    p.serialize("tmp/#{filename}.xlsx")
    assert_equal true, wb.styles_applied
    assert_equal 8, wb.styles.style_index.count
    assert_equal 8, wb.styled_cells.count
  end

  def test_multiple_style_borders_on_same_cells
    p = Axlsx::Package.new
    wb = p.workbook

    filename = 'multiple_style_borders'
    wb.add_worksheet do |sheet|
      sheet.add_row
      sheet.add_row ['', 'B2', 'C2', 'D2']
      sheet.add_row ['', 'B3', 'C3', 'D3']

      sheet.add_border 'B2:D3', :all
      sheet.add_border 'B2:D2', edges: [:bottom], style: :thick, color: 'ff0000'
    end
    p.serialize("tmp/#{filename}.xlsx")
    assert_equal true, wb.styles_applied
    assert_equal 6, wb.styles.style_index.count
    assert_equal 6, wb.styled_cells.count

    b2_cell_style = {
      border: {
        style: :thick,
        color: 'ff0000',
        edges: [:bottom, :left, :top]
      },
      type: :xf,
      name: 'Arial',
      sz: 11,
      family: 1
    }
    assert_equal b2_cell_style, wb.styles.style_index.values.find{|x| x == b2_cell_style}

    d3_cell_style = {
      border: {
        style: :thin,
        color: '000000',
        edges: [:bottom, :right]
      },
      type: :xf,
      name: 'Arial',
      sz: 11,
      family: 1
    }
    assert_equal d3_cell_style, wb.styles.style_index.values.find{|x| x == d3_cell_style}
  end

  def test_mixed_borders_1
    p = Axlsx::Package.new
    wb = p.workbook

    filename = 'mixed_borders_1'
    wb.add_worksheet do |sheet|
      sheet.add_row
      sheet.add_row ['', '1', '2', '3']
      sheet.add_row ['', '4', '5', '6']
      sheet.add_row ['', '7', '8', '9']
      sheet.add_style 'B2:D4', border: { style: :thin, color: '000000' }
      sheet.add_border 'C3:D4', style: :medium
    end
    wb.apply_styles
    assert_equal 9, wb.styled_cells.count
    assert_equal 2, wb.styles.style_index.count
    p.serialize("tmp/#{filename}.xlsx")
    assert_equal true, wb.styles_applied
  end

  def test_mixed_borders_2
    p = Axlsx::Package.new
    wb = p.workbook

    filename = 'mixed_borders_2'
    wb.add_worksheet do |sheet|
      sheet.add_row
      sheet.add_row ['', '1', '2', '3']
      sheet.add_row ['', '4', '5', '6']
      sheet.add_row ['', '7', '8', '9']
      sheet.add_border 'B2:D4', style: :medium
      sheet.add_style 'D2:D4', border: { style: :thin, color: '000000' }
    end
    wb.apply_styles
    assert_equal 8, wb.styled_cells.count
    assert_equal 6, wb.styles.style_index.count
    p.serialize("tmp/#{filename}.xlsx")
    assert_equal true, wb.styles_applied
  end

  def test_dxf_cell
    p = Axlsx::Package.new
    wb = p.workbook

    wb.add_worksheet do |sheet|
      sheet.add_row (1..2).to_a
      sheet.add_style "A1:A1", { bg_color: "AA0000" }

      sheet.add_row (1..2).to_a
      sheet.add_style "B1:B1", { bg_color: "CC0000" }

      sheet.add_row (1..2).to_a
      sheet.add_style "A3:B3", { bg_color: "00FF00" }

      wb.styles.add_style(bg_color: "0000FF", type: :dxf)
    end

    p.serialize("tmp/test_dxf_cell.xlsx")
    assert_equal true, wb.styles_applied

    assert_equal 1, wb.styles.dxfs.count

    assert_equal 6, wb.styles.cellXfs.count
  end

  def test_default_font_with_style_index
    p = Axlsx::Package.new
    wb = p.workbook

    wb.styles.fonts[0].name = 'Pontiac' ### TODO, is this a valid font name in all environments
    wb.styles.fonts[0].sz = 12

    wb.add_worksheet do |sheet|
      sheet.add_row [1,2,3]
      sheet.add_style "A1:C1", { color: "FFFFFF" }
    end

    wb.apply_styles

    assert_equal 1, wb.styles.style_index.size
    
    assert_equal(
      {
        type: :xf, 
        name: "Pontiac", 
        sz: 12, 
        family: 1, 
        color: "FFFFFF",
      }, 
      wb.styles.style_index.values.first
    )
  end

end
