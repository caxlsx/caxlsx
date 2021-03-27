require 'tc_helper.rb'

class TestRow < Test::Unit::TestCase

  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet :name=>"hmmm"
    @row = @ws.add_row
  end

  def test_initialize
    assert(@row.cells.empty?, "no cells by default")
    assert_equal(@row.worksheet, @ws, "has a reference to the worksheet")
    assert_nil(@row.height, "height defaults to nil")
    assert(!@row.custom_height, "no custom height by default")
  end

  def test_initialize_with_fixed_height
    row = @ws.add_row([1,2,3,4,5], :height=>40)
    assert_equal(40, row.height)
    assert(row.custom_height)
  end

  def test_style
    r = @ws.add_row([1,2,3,4,5])
    r.style=1
    r.cells.each { |c| assert_equal(c.style,1) }
  end

  def test_color
    r = @ws.add_row([1,2,3,4,5])
    r.color = "FF00FF00"
    r.cells.each { |c| assert_equal(c.color.rgb, "FF00FF00") }
  end

  def test_index
    assert_equal(@row.row_index, @row.worksheet.rows.index(@row))
  end

  def test_add_cell
    c = @row.add_cell(1)
    assert_equal(@row.cells.last, c)
  end

  def test_add_cell_autowidth_info
    cell = @row.add_cell("this is the cell of cells")
    width = cell.send(:autowidth)
    assert_equal(@ws.column_info.last.width, width)
  end

  def test_array_to_cells
    r = @ws.add_row [1,2,3], :style=>1, :types=>[:integer, :string, :float]
    assert_equal(r.cells.size, 3)
    r.cells.each do |c|
      assert_equal(c.style, 1)
    end
    r = @ws.add_row [1,2,3], :style=>[1]
    assert_equal(r.cells.first.style, 1, "only apply style to cells with at the same index of of the style array")
    assert_equal(r.cells.last.style, 0, "only apply style to cells with at the same index of of the style array")

  end

  def test_array_to_cells_with_escape_formulas
    row = ['=HYPERLINK("http://www.example.com", "CSV Payload")', '=Bar']
    @ws.add_row row, escape_formulas: true

    assert_equal @ws.rows.last.cells[0].escape_formulas, true
    assert_equal @ws.rows.last.cells[1].escape_formulas, true
  end

  def test_array_to_cells_with_escape_formulas_as_an_array
    row = ['=HYPERLINK("http://www.example.com", "CSV Payload")', '+Foo', '-Bar']
    @ws.add_row row, escape_formulas: [true, false, true]

    assert_equal @ws.rows.last.cells.first.escape_formulas, true
    assert_equal @ws.rows.last.cells[1].escape_formulas, false
    assert_equal @ws.rows.last.cells[2].escape_formulas, true
  end

  def test_custom_height
    @row.height = 20
    assert(@row.custom_height)
  end

  def test_height
    assert_raise(ArgumentError) { @row.height = -3 }
    assert_nothing_raised { @row.height = 15 }
    assert_equal(15, @row.height)
  end


  def test_ph
    assert_raise(ArgumentError) { @row.ph = -3 }
    assert_nothing_raised { @row.ph = true }
    assert_equal(true, @row.ph)
  end

  def test_hidden
    assert_raise(ArgumentError) { @row.hidden = -3 }
    assert_nothing_raised { @row.hidden = true }
    assert_equal(true, @row.hidden)
  end

  def test_collapsed
    assert_raise(ArgumentError) { @row.collapsed = -3 }
    assert_nothing_raised { @row.collapsed = true }
    assert_equal(true, @row.collapsed)
  end

  def test_outlineLevel
    assert_raise(ArgumentError) { @row.outlineLevel = -3 }
    assert_nothing_raised { @row.outlineLevel = 2 }
    assert_equal(2, @row.outlineLevel)
  end

  def test_to_xml_without_custom_height
    doc = Nokogiri::XML.parse(@row.to_xml_string(0))
    assert_equal(0, doc.xpath(".//row[@ht]").size)
    assert_equal(0, doc.xpath(".//row[@customHeight]").size)
  end

  def test_to_xml_string
    @row.height = 20
    @row.s = 1
    @row.outlineLevel = 2
    @row.collapsed = true
    @row.hidden = true
    r_s_xml = Nokogiri::XML(@row.to_xml_string(0, ''))
    assert_equal(r_s_xml.xpath(".//row[@r=1]").size, 1)
  end

  def test_to_xml_string_with_custom_height
    @row.add_cell 1
    @row.height = 20
    r_s_xml = Nokogiri::XML(@row.to_xml_string(0, ''))
    assert_equal(r_s_xml.xpath(".//row[@r=1][@ht=20][@customHeight=1]").size, 1)
  end

  def test_offsets
    offset = 3
    values = [1,2,3,4,5]
    r = @ws.add_row(values, offset: offset, style: 1)
    r.cells.each_with_index do |c, index| 
      assert_equal(c.style, index < offset ? 0 : 1)
      assert_equal(c.value, index < offset ? nil : values[index - offset])
    end
  end

  def test_offsets_with_styles
    offset = 3
    values = [1,2,3,4,5]
    styles = (1..5).map{ @ws.workbook.styles.add_style }
    r = @ws.add_row(values, offset: offset, style: styles)
    r.cells.each_with_index do |c, index| 
      assert_equal(c.style, index < offset ? 0 : styles[index-offset])
      assert_equal(c.value, index < offset ? nil : values[index - offset])
    end
  end

end
