require 'tc_helper'

class TestWorksheet < Test::Unit::TestCase
  def setup
    @p = Axlsx::Package.new
    @wb = @p.workbook
    @ws = @wb.add_worksheet
  end

  def test_pn
    assert_equal(@ws.pn, "worksheets/sheet1.xml")
    ws = @ws.workbook.add_worksheet
    assert_equal(ws.pn, "worksheets/sheet2.xml")
  end

  def test_name_is_html_encoded
    @ws.name = '<foo> & <bar>'
    assert_equal(@ws.name, '&lt;foo&gt; &amp; &lt;bar&gt;')
  end

  def test_name_exception_on_invalid_character
    assert_raises(ArgumentError) { @ws.name = 'foo:bar' }
    assert_raises(ArgumentError) { @ws.name = 'foo[bar' }
    assert_raises(ArgumentError) { @ws.name = 'foo]bar' }
    assert_raises(ArgumentError) { @ws.name = 'foo*bar' }
    assert_raises(ArgumentError) { @ws.name = 'foo/bar' }
    assert_raises(ArgumentError) { @ws.name = 'foo\bar' }
    assert_raises(ArgumentError) { @ws.name = 'foo?bar' }
  end

  def test_name_unique
    assert_raise(ArgumentError, "worksheet name must be unique") { n = @ws.name; @ws.workbook.add_worksheet(:name => n) }
  end

  def test_name_unique_only_checks_other_worksheet_names
    assert_nothing_raised { @ws.name = @ws.name }
    assert_nothing_raised { Axlsx::Package.new.workbook.add_worksheet :name => 'Sheet1' }
  end

  def test_exception_if_name_too_long
    assert_nothing_raised { @ws.name = 'x' * 31 }
    assert_raises(ArgumentError) { @ws.name = 'x' * 32 }
  end

  def test_exception_if_name_too_long_because_of_multibyte_characters
    four_characters_for_excel = "\u{1F1EB 1F1F7}" # french flag emoji
    assert_raises(ArgumentError, "name too long!") do
      @ws.name = four_characters_for_excel + "x" * 28
    end
    assert_nothing_raised { @ws.name = "#{four_characters_for_excel}123456789012345678901234567" }
    assert_nothing_raised { @ws.name = "123456789012345678901234567890…" }
    assert_nothing_raised { @ws.name = "123456789012345678901234567890✔" }
  end

  def test_page_margins
    assert(@ws.page_margins.is_a?(Axlsx::PageMargins))
  end

  def test_page_margins_yeild
    @ws.page_margins do |pm|
      assert(pm.is_a?(Axlsx::PageMargins))
      assert(@ws.page_margins == pm)
    end
  end

  def test_page_setup
    assert(@ws.page_setup.is_a?(Axlsx::PageSetup))
  end

  def test_page_setup_yield
    @ws.page_setup do |ps|
      assert(ps.is_a?(Axlsx::PageSetup))
      assert(@ws.page_setup == ps)
    end
  end

  def test_print_options
    assert(@ws.print_options.is_a?(Axlsx::PrintOptions))
  end

  def test_print_options_yield
    @ws.print_options do |po|
      assert(po.is_a?(Axlsx::PrintOptions))
      assert(@ws.print_options == po)
    end
  end

  def test_header_footer
    assert(@ws.header_footer.is_a?(Axlsx::HeaderFooter))
  end

  def test_header_footer_yield
    @ws.header_footer do |hf|
      assert(hf.is_a?(Axlsx::HeaderFooter))
      assert(@ws.header_footer == hf)
    end
  end

  def test_state
    assert_equal(:visible, @ws.state)
  end

  def test_state_validation
    assert_raise(ArgumentError) { @ws.state = :dead }
    assert_nothing_raised { @ws.state = :very_hidden }
  end

  def test_no_autowidth
    @ws.workbook.use_autowidth = false
    @ws.add_row [1, 2, 3, 4]
    assert_equal(@ws.column_info[0].width, nil)
  end

  def test_initialization_options
    page_margins = { :left => 2, :right => 2, :bottom => 2, :top => 2, :header => 2, :footer => 2 }
    page_setup = { :fit_to_height => 1, :fit_to_width => 1, :orientation => :landscape, :paper_width => "210mm", :paper_height => "297mm", :scale => 80 }
    print_options = { :grid_lines => true, :headings => true, :horizontal_centered => true, :vertical_centered => true }
    header_footer = { :different_first => false, :different_odd_even => false, :odd_header => 'Header' }
    optioned = @ws.workbook.add_worksheet(:name => 'bob', :page_margins => page_margins, :page_setup => page_setup, :print_options => print_options, :header_footer => header_footer)
    page_margins.each_key do |key|
      assert_equal(page_margins[key], optioned.page_margins.send(key))
    end
    page_setup.each_key do |key|
      assert_equal(page_setup[key], optioned.page_setup.send(key))
    end
    print_options.each_key do |key|
      assert_equal(print_options[key], optioned.print_options.send(key))
    end
    header_footer.each_key do |key|
      assert_equal(header_footer[key], optioned.header_footer.send(key))
    end
    assert_equal(optioned.name, 'bob')
  end

  # def test_use_gridlines
  #  assert_raise(ArgumentError) { @ws.show_gridlines = -1.1 }
  #  assert_nothing_raised { @ws.show_gridlines = false }
  #  assert_equal(@ws.show_gridlines, false)
  # end

  # def test_selected
  #  assert_raise(ArgumentError) { @ws.selected = -1.1 }
  #  assert_nothing_raised { @ws.selected = true }
  #  assert_equal(@ws.selected, true)
  # end

  def test_rels_pn
    assert_equal(@ws.rels_pn, "worksheets/_rels/sheet1.xml.rels")
    ws = @ws.workbook.add_worksheet
    assert_equal(ws.rels_pn, "worksheets/_rels/sheet2.xml.rels")
  end

  def test_rId
    assert_equal @ws.workbook.relationships.for(@ws).Id, @ws.rId
  end

  def test_index
    assert_equal(@ws.index, @ws.workbook.worksheets.index(@ws))
  end

  def test_dimension
    @ws.add_row [1, 2, 3]
    @ws.add_row [4, 5, 6]
    assert_equal @ws.dimension.sqref, "A1:C2"
  end

  def test_dimension_with_empty_row
    @ws.add_row
    assert_equal "A1:AA200", @ws.dimension.sqref
  end

  def test_referencing
    @ws.add_row [1, 2, 3]
    @ws.add_row [4, 5, 6]
    range = @ws["A1:C2"]
    first_row = @ws[0]
    last_row = @ws[1]
    assert_equal(@ws.rows[0], first_row)
    assert_equal(@ws.rows[1], last_row)
    assert_equal(range.size, 6)
    assert_equal(range.first, @ws.rows.first.cells.first)
    assert_equal(range.last, @ws.rows.last.cells.last)
  end

  def test_add_row
    assert(@ws.rows.empty?, "sheet has no rows by default")
    r = @ws.add_row([1, 2, 3])
    assert_equal(@ws.rows.size, 1, "add_row adds a row")
    assert_equal(@ws.rows.first, r, "the row returned is the row added")
  end

  def test_add_chart
    assert(@ws.workbook.charts.empty?, "the sheet's workbook should not have any charts by default")
    @ws.add_chart Axlsx::Pie3DChart
    assert_equal(@ws.workbook.charts.size, 1, "add_chart adds a chart to the workbook")
  end

  def test_add_page_break_with_string_cell_ref
    assert(@ws.row_breaks.empty?)
    assert(@ws.col_breaks.empty?)
    @ws.add_page_break("B1")
    assert_equal(1, @ws.row_breaks.size)
    assert_equal(1, @ws.col_breaks.size)
  end

  def test_add_page_break_with_cell
    @ws.add_row [1, 2, 3, 4]
    @ws.add_row [1, 2, 3, 4]

    assert(@ws.row_breaks.empty?)
    assert(@ws.col_breaks.empty?)
    @ws.add_page_break(@ws.rows.last.cells[1])
    assert_equal(1, @ws.row_breaks.size)
    assert_equal(1, @ws.col_breaks.size)
  end

  def test_drawing
    assert @ws.drawing.nil?
    @ws.add_chart(Axlsx::Pie3DChart)
    assert @ws.drawing.is_a?(Axlsx::Drawing)
  end

  def test_add_pivot_table
    assert(@ws.workbook.pivot_tables.empty?, "the sheet's workbook should not have any pivot tables by default")
    @ws.add_pivot_table 'G5:G6', 'A1:D:10'
    assert_equal(@ws.workbook.pivot_tables.size, 1, "add_pivot_tables adds a pivot_table to the workbook")
  end

  def test_col_style
    @ws.add_row [1, 2, 3, 4]
    @ws.add_row [1, 2, 3, 4]
    @ws.add_row [1, 2, 3, 4]
    @ws.add_row [1, 2, 3, 4]
    @ws.col_style((1..2), 1, :row_offset => 1)
    @ws.rows[(1..-1)].each do |r|
      assert_equal(r.cells[1].style, 1)
      assert_equal(r.cells[2].style, 1)
    end
    assert_equal(@ws.rows.first.cells[1].style, 0)
    assert_equal(@ws.rows.first.cells[0].style, 0)
  end

  def test_col_style_with_empty_column
    @ws.add_row [1, 2, 3, 4]
    @ws.add_row [1]
    @ws.add_row [1, 2, 3, 4]
    assert_nothing_raised { @ws.col_style(1, 1) }
  end

  def test_cols
    @ws.add_row [1, 2, 3, 4]
    @ws.add_row [1, 2, 3, 4]
    @ws.add_row [1, 2, 3]
    @ws.add_row [1, 2, 3, 4]
    c = @ws.cols[1]
    assert_equal(c.size, 4)
    assert_equal(c[0].value, 2)
  end

  def test_cols_with_block
    @ws.add_row [1, 2, 3]
    @ws.add_row [1]
    cols = @ws.cols { |_row, _column| :foo }
    assert_equal(:foo, cols[1][1])
  end

  def test_row_style
    @ws.add_row [1, 2, 3, 4]
    @ws.add_row [1, 2, 3, 4]
    @ws.add_row [1, 2, 3, 4]
    @ws.add_row [1, 2, 3, 4]
    @ws.row_style 1, 1, :col_offset => 1
    @ws.rows[1].cells[(1..-1)].each do |c|
      assert_equal(c.style, 1)
    end
    assert_equal(@ws.rows[1].cells[0].style, 0)
    assert_equal(@ws.rows[2].cells[1].style, 0)
    @ws.row_style(1..2, 1, :col_offset => 2)
    @ws.rows[(1..2)].each do |r|
      r.cells[(2..-1)].each do |c|
        assert_equal(c.style, 1)
      end
    end
  end

  # def test_to_xml_string_fit_to_page
  #  @ws.page_setup.fit_to_width = 1
  #  doc = Nokogiri::XML(@ws.to_xml_string)
  #  assert_equal(doc.xpath('//xmlns:worksheet/xmlns:sheetPr/xmlns:pageSetUpPr[@fitToPage=1]').size, 1)
  # end

  def test_to_xml_string_dimensions
    @ws.add_row [1, 2, 3]
    doc = Nokogiri::XML(@ws.to_xml_string)
    assert_equal(doc.xpath('//xmlns:worksheet/xmlns:dimension[@ref="A1:C1"]').size, 1)
  end

  # def test_fit_to_page_assignation_does_nothing
  #  @ws.fit_to_page = true
  #  assert_equal(@ws.fit_to_page?, false)
  # end

  # def test_to_xml_string_selected
  #  @ws.selected = true
  #  doc = Nokogiri::XML(@ws.to_xml_string)
  #  assert_equal(doc.xpath('//xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView[@tabSelected=1]').size, 1)
  # end

  # def test_to_xml_string_show_gridlines
  #  @ws.show_gridlines = false
  #  doc = Nokogiri::XML(@ws.to_xml_string)
  #  assert_equal(doc.xpath('//xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView[@showGridLines=0]').size, 1)
  # end

  def test_to_xml_string_auto_fit_data
    @ws.add_row [1, "two"]
    doc = Nokogiri::XML(@ws.to_xml_string)
    assert_equal(doc.xpath('//xmlns:worksheet/xmlns:cols/xmlns:col').size, 2)
  end

  def test_to_xml_string_sheet_data
    @ws.add_row [1, "two"]
    doc = Nokogiri::XML(@ws.to_xml_string)
    assert_equal(doc.xpath('//xmlns:worksheet/xmlns:sheetData/xmlns:row').size, 1)
  end

  def test_to_xml_string_auto_filter
    @ws.add_row [1, "two"]
    @ws.auto_filter.range = "A1:B1"
    doc = Nokogiri::XML(@ws.to_xml_string)
    assert_equal(doc.xpath('//xmlns:worksheet/xmlns:autoFilter[@ref="A1:B1"]').size, 1)
  end

  def test_to_xml_string_merge_cells
    @ws.add_row [1, "two"]
    @ws.merge_cells "A1:D1"
    @ws.merge_cells "E1:F1"
    doc = Nokogiri::XML(@ws.to_xml_string)
    assert_equal(doc.xpath('//xmlns:worksheet/xmlns:mergeCells/xmlns:mergeCell[@ref="A1:D1"]').size, 1)
    assert_equal(doc.xpath('//xmlns:worksheet/xmlns:mergeCells/xmlns:mergeCell[@ref="E1:F1"]').size, 1)
  end

  def test_to_xml_string_merge_cells_row
    row = @ws.add_row [1, "two"]
    @ws.merge_cells row
    doc = Nokogiri::XML(@ws.to_xml_string)
    assert_equal(doc.xpath('//xmlns:worksheet/xmlns:mergeCells/xmlns:mergeCell[@ref="A1:B1"]').size, 1)
  end

  def test_to_xml_string_row_breaks
    @ws.add_page_break("A1")
    doc = Nokogiri::XML(@ws.to_xml_string)
    assert_equal(doc.xpath('//xmlns:worksheet/xmlns:rowBreaks/xmlns:brk[@id="0"]').size, 1)
  end

  def test_to_xml_string_sheet_protection
    @ws.sheet_protection.password = 'fish'
    doc = Nokogiri::XML(@ws.to_xml_string)
    assert(doc.xpath('//xmlns:sheetProtection'))
  end

  def test_to_xml_string_page_margins
    @ws.page_margins do |pm|
      pm.left = 9
      pm.right = 7
    end
    doc = Nokogiri::XML(@ws.to_xml_string)
    assert_equal(doc.xpath('//xmlns:worksheet/xmlns:pageMargins[@left="9"][@right="7"]').size, 1)
  end

  def test_to_xml_string_page_setup
    @ws.page_setup do |ps|
      ps.paper_width = "210mm"
      ps.scale = 80
    end
    doc = Nokogiri::XML(@ws.to_xml_string)
    assert_equal(doc.xpath('//xmlns:worksheet/xmlns:pageSetup[@paperWidth="210mm"][@scale="80"]').size, 1)
  end

  def test_to_xml_string_print_options
    @ws.print_options do |po|
      po.grid_lines = true
      po.horizontal_centered = true
    end
    doc = Nokogiri::XML(@ws.to_xml_string)
    assert_equal(doc.xpath('//xmlns:worksheet/xmlns:printOptions[@gridLines=1][@horizontalCentered=1]').size, 1)
  end

  def test_to_xml_string_header_footer
    @ws.header_footer do |hf|
      hf.different_first = false
      hf.different_odd_even = false
      hf.odd_header = 'Test Header'
    end
    doc = Nokogiri::XML(@ws.to_xml_string)
    assert_equal(doc.xpath('//xmlns:worksheet/xmlns:headerFooter[@differentFirst=0][@differentOddEven=0]').size, 1)
  end

  def test_to_xml_string_drawing
    @ws.add_chart Axlsx::Pie3DChart
    doc = Nokogiri::XML(@ws.to_xml_string)
    assert_equal @ws.send(:worksheet_drawing).relationship.Id, doc.xpath('//xmlns:worksheet/xmlns:drawing').first["r:id"]
  end

  def test_to_xml_string_tables
    @ws.add_row ["one", "two"]
    @ws.add_row [1, 2]
    table = @ws.add_table "A1:B2"
    doc = Nokogiri::XML(@ws.to_xml_string)
    assert_equal(doc.xpath('//xmlns:worksheet/xmlns:tableParts[@count="1"]').size, 1)
    assert_equal table.rId, doc.xpath('//xmlns:worksheet/xmlns:tableParts/xmlns:tablePart').first["r:id"]
  end

  def test_to_xml_string
    schema = Nokogiri::XML::Schema(File.open(Axlsx::SML_XSD))
    doc = Nokogiri::XML(@ws.to_xml_string)
    assert(schema.validate(doc).map { |e| puts e.message; e }.empty?, "error free validation")
  end

  def test_styles
    assert(@ws.styles.is_a?(Axlsx::Styles), 'worksheet provides access to styles')
  end

  def test_to_xml_string_with_illegal_chars
    old = Axlsx::trust_input
    Axlsx::trust_input = false
    nasties = "\v\u2028\u0001\u0002\u0003\u0004\u0005\u0006\u0007\u0008\u001f"
    @ws.add_row [nasties]
    assert_equal(0, @ws.rows.last.cells.last.value.index("\v"))
    assert_equal(nil, @ws.to_xml_string.index("\v"))
    Axlsx::trust_input = old
  end

  def test_to_xml_string_with_newlines
    cell_with_newline = "foo\n\r\nbar"
    @ws.add_row [cell_with_newline]
    assert_equal("foo\n\r\nbar", @ws.rows.last.cells.last.value)
    assert_not_nil(@ws.to_xml_string.index("foo\n\r\nbar"))
  end

  # Make sure the XML for all optional elements (like pageMargins, autoFilter, ...)
  # is generated in correct order.
  def test_valid_with_optional_elements
    @ws.page_margins.set :left => 9
    @ws.page_setup.set :fit_to_width => 1
    @ws.print_options.set :headings => true
    @ws.auto_filter.range = "A1:C3"
    @ws.merge_cells "A4:A5"
    @ws.add_chart Axlsx::Pie3DChart
    @ws.add_table "E1:F3"
    @ws.add_pivot_table 'G5:G6', 'A1:D10'
    schema = Nokogiri::XML::Schema(File.open(Axlsx::SML_XSD))
    doc = Nokogiri::XML(@ws.to_xml_string)
    assert(schema.validate(doc).map { |e| puts e.message; e }.empty?, schema.validate(doc).map(&:message).join('\n'))
  end

  def test_relationships
    @ws.add_row [1, 2, 3]
    assert(@ws.relationships.empty?, "No Drawing relationship until you add a chart")
    @ws.add_chart Axlsx::Pie3DChart
    assert_equal(@ws.relationships.size, 1, "adding a chart creates the relationship")
    @ws.add_chart Axlsx::Pie3DChart
    assert_equal(@ws.relationships.size, 1, "multiple charts still only result in one relationship")
    @ws.add_comment :text => 'builder', :author => 'bob', :ref => @ws.rows.last.cells.last
    assert_equal(@ws.relationships.size, 3, "adding a comment adds 2 relationships")
    @ws.add_comment :text => 'not that is a comment!', :author => 'travis', :ref => "A1"
    assert_equal(@ws.relationships.size, 3, "adding multiple comments in the same worksheet should not add any additional comment relationships")
    @ws.add_pivot_table 'G5:G6', 'A1:D10'
    assert_equal(@ws.relationships.size, 4, "adding a pivot table adds 1 relationship")
  end

  def test_set_fixed_width_column
    @ws.add_row ["mule", "donkey", "horse"], :widths => [20, :ignore, nil]
    assert(@ws.column_info.size == 3, "a data item for each column")
    assert_equal(20, @ws.column_info[0].width, "adding a row with fixed width updates :fixed attribute")
    assert_equal(@ws.column_info[1].width, nil, ":ignore does not set any data")
  end

  def test_fixed_height
    @ws.add_row [1, 2, 3], :height => 40
    assert_equal(40, @ws.rows[-1].height)
  end

  def test_set_column_width
    @ws.add_row ["chasing windmills", "penut"]
    @ws.column_widths nil, 0.5
    assert_equal(@ws.column_info[1].width, 0.5, 'eat my width')
    assert_raise(ArgumentError, 'only accept unsigned ints') { @ws.column_widths 2, 7, -1 }
    assert_raise(ArgumentError, 'only accept Integer or Float') { @ws.column_widths 2, 7, "-1" }
  end

  def test_ranges
    @ws.add_row(["foo", "bar"])
    @ws.add_row(["a", "b", "c", "d", "e"])

    valid_range = "A1:B2"
    @ws[valid_range]

    valid_range = "A1:E2" ### Missing middle cells (C1 - E1), still allowed
    @ws[valid_range]

    assert_raise ArgumentError, "Invalid cell definition" do
      invalid_range_format = "A1:B2:C3"
      @ws[invalid_range_format]
    end

    assert_raise ArgumentError, "Missing cell `Z1` for the specified range." do
      invalid_row_range = "A1:Z1"
      @ws[invalid_row_range]
    end

    assert_raise ArgumentError, "Missing cell `D1` for the specified range." do
      invalid_cell_range = "D1:E2" ### Missing start cell, not allowed
      @ws[invalid_cell_range]
    end

    assert_raise ArgumentError, "Missing cell `A99` for the specified range." do
      invalid_cell_range = "A1:A99" ### Missing end cell, not allowed
      @ws[invalid_cell_range]
    end
  end

  def test_protect_range
    assert(@ws.send(:protected_ranges).is_a?(Axlsx::SimpleTypedList))
    assert_equal(0, @ws.send(:protected_ranges).size)
    @ws.protect_range('A1:A3')
    assert_equal('A1:A3', @ws.send(:protected_ranges).last.sqref)
  end

  def test_protect_range_with_cells
    @ws.add_row [1, 2, 3]
    assert_nothing_raised { @ws.protect_range(@ws.rows.first.cells) }
    assert_equal('A1:C1', @ws.send(:protected_ranges).last.sqref)
  end

  def test_merge_cells
    @ws.add_row [1, 2, 3]
    @ws.add_row [4, 5, 6]
    @ws.add_row [7, 8, 9]
    @ws.merge_cells "A1:A2"
    @ws.merge_cells "B2:C3"
    @ws.merge_cells @ws.rows.last.cells[(0..1)]
    assert_equal(@ws.send(:merged_cells).size, 3)
    assert_equal(@ws.send(:merged_cells).last, "A3:B3")
  end

  def test_merge_cells_sorts_correctly_by_row_when_given_array
    10.times do |i|
      @ws.add_row [i]
    end
    @ws.merge_cells [@ws.rows[8].cells.first, @ws.rows[9].cells.first]
    assert_equal "A9:A10", @ws.send(:merged_cells).first
  end

  def test_auto_filter
    assert(@ws.auto_filter.range.nil?)
    assert(@wb.defined_names.none? { |defined_name| defined_name.name == '_xlnm._FilterDatabase' })
    assert_raise(ArgumentError) { @ws.auto_filter = 123 }
    @ws.auto_filter.range = "A1:D9"
    assert_equal(@ws.auto_filter.range, "A1:D9")
    @ws.to_xml_string
    assert(@wb.defined_names.any? { |defined_name| defined_name.name == '_xlnm._FilterDatabase' })
  end

  def test_auto_filter_assign
    other_ws = @wb.add_worksheet

    assert(@ws.auto_filter.range.nil?)
    assert(other_ws.auto_filter.range.nil?)
    assert(@wb.defined_names.none? { |defined_name| defined_name.name == '_xlnm._FilterDatabase' })
    assert_raise(ArgumentError) { @ws.auto_filter = 123 }

    @ws.auto_filter = "A1:D9"
    assert_equal(@ws.auto_filter.range, "A1:D9")

    other_ws.auto_filter = "A1:D2"
    assert_equal(other_ws.auto_filter.range, "A1:D2")

    @ws.to_xml_string
    other_ws.to_xml_string

    filter_database = @wb.defined_names.select { |defined_name| defined_name.name == '_xlnm._FilterDatabase' }
    assert_equal(2, filter_database.size)
    assert_equal(@ws.index, filter_database[0].local_sheet_id)
    assert_equal(other_ws.index, filter_database[1].local_sheet_id)
  end

  def test_sheet_pr_for_auto_filter
    @ws.auto_filter.range = 'A1:D9'
    @ws.auto_filter.add_column 0, :filters, :filter_items => [1]
    doc = Nokogiri::XML(@ws.to_xml_string)
    assert(doc.xpath('//sheetPr[@filterMode=1]'))
  end

  def test_outline_level_rows
    3.times { @ws.add_row [1, 2, 3] }
    @ws.outline_level_rows 0, 2
    assert_equal(1, @ws.rows[0].outline_level)
    assert_equal(true, @ws.rows[2].hidden)
    assert_equal(true, @ws.sheet_view.show_outline_symbols)
  end

  def test_outline_level_columns
    3.times { @ws.add_row [1, 2, 3] }
    @ws.outline_level_columns 0, 2
    assert_equal(1, @ws.column_info[0].outline_level)
    assert_equal(true, @ws.column_info[2].hidden)
    assert_equal(true, @ws.sheet_view.show_outline_symbols)
  end

  def test_worksheet_does_not_get_added_to_workbook_on_initialize_failure
    assert_equal(1, @wb.worksheets.size)
    assert_raise(ArgumentError) { @wb.add_worksheet(:name => 'Sheet1') }
    assert_equal(1, @wb.worksheets.size)
  end

  def test_worksheet_only_includes_outline_pr_when_set
    doc = Nokogiri::XML(@ws.to_xml_string)
    assert_equal(doc.xpath('//xmlns:worksheet/xmlns:sheetPr/xmlns:outlinePr').size, 0)

    @ws.sheet_pr.outline_pr.summary_below = false
    @ws.sheet_pr.outline_pr.summary_right = true
    doc = Nokogiri::XML(@ws.to_xml_string)
    assert_equal(doc.xpath('//xmlns:worksheet/xmlns:sheetPr/xmlns:outlinePr').size, 1)
    assert_equal(doc.xpath('//xmlns:worksheet/xmlns:sheetPr/xmlns:outlinePr[@summaryBelow=0][@summaryRight=1]').size, 1)
  end

  def test_merge_styles_1
    p = Axlsx::Package.new
    wb = p.workbook

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
  end

  def test_merge_styles_2
    p = Axlsx::Package.new
    wb = p.workbook

    bold = wb.styles.add_style b: true

    wb.add_worksheet do |sheet|
      sheet.add_row ['A1', 'B1'], style: [nil, bold]
      sheet.add_row ['A2', 'B2'], style: bold
      sheet.add_row ['A3', 'B3']
      sheet.add_style 'A1:A2', i: true
    end

    wb.apply_styles

    assert_equal 3, wb.styles.style_index.count
  end

  def test_merge_styles_3
    p = Axlsx::Package.new
    wb = p.workbook

    bold = wb.styles.add_style b: true

    wb.add_worksheet do |sheet|
      sheet.add_row ['A1', 'B1'], style: [nil, bold]
      sheet.add_row ['A2', 'B2']
      sheet.add_style 'B1:B2', bg_color: 'FF0000'
    end

    wb.apply_styles

    assert_equal 3, wb.styles.style_index.count
  end

  def test_table_with_borders
    p = Axlsx::Package.new
    wb = p.workbook

    wb.add_worksheet do |sheet|
      sheet.add_row
      sheet.add_row ['', 'Product', 'Category',  'Price']
      sheet.add_row ['', 'Butter', 'Dairy',      4.99]
      sheet.add_row ['', 'Bread', 'Baked Goods', 3.45]
      sheet.add_row ['', 'Broccoli', 'Produce',  2.99]
      sheet.add_row ['', 'Pizza', 'Frozen Foods', 4.99]
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

    wb.apply_styles

    assert_equal true, wb.styles_applied
    assert_equal 12, wb.styles.style_index.count
    assert_equal 12 + 2, wb.styles.style_index.keys.max
  end

  def test_add_border_arguments
    p = Axlsx::Package.new
    wb = p.workbook

    wb.add_worksheet do |sheet|
      20.times.each do
        sheet.add_row [1, 2, 3, 4, 5]
      end

      sheet.add_border "B2:D5"
      sheet.add_border ["C3:C4", "D3:D4", "A2"], { color: "FF00000" }
      sheet.add_border "B10:E10", [:top, :bottom, :left, :right]
      sheet.add_border "B12:E12", :all
      sheet.add_border "B14:E14", Axlsx::Border::EDGES
      sheet.add_border "B16:E16", { edges: :all, style: :thick }
      sheet.add_border "B18:E18", { edges: [:top], style: :thick, color: "FFFFF00" }
    end

    wb.apply_styles

    assert_equal true, wb.styles_applied
    assert_equal 17, wb.styles.style_index.count
  end

  def test_duplicate_borders
    p = Axlsx::Package.new
    wb = p.workbook

    wb.add_worksheet do |sheet|
      sheet.add_row
      sheet.add_row ['', 'B2', 'C2', 'D2']
      sheet.add_row ['', 'B3', 'C3', 'D3']
      sheet.add_row ['', 'B4', 'C4', 'D4']

      sheet.add_border 'B2:D4'
      sheet.add_border 'B2:D4'
    end

    wb.apply_styles

    assert_equal true, wb.styles_applied
    assert_equal 8, wb.styles.style_index.count
    assert_equal 8, wb.styled_cells.count
  end

  def test_multiple_style_borders_on_same_cells
    p = Axlsx::Package.new
    wb = p.workbook

    wb.add_worksheet do |sheet|
      sheet.add_row
      sheet.add_row ['', 'B2', 'C2', 'D2']
      sheet.add_row ['', 'B3', 'C3', 'D3']

      sheet.add_border 'B2:D3', :all
      sheet.add_border 'B2:D2', edges: [:bottom], style: :thick, color: 'ff0000'
    end

    wb.apply_styles

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
    assert_equal b2_cell_style, wb.styles.style_index.values.find { |x| x == b2_cell_style }

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
    assert_equal d3_cell_style, wb.styles.style_index.values.find { |x| x == d3_cell_style }
  end

  def test_mixed_borders_1
    p = Axlsx::Package.new
    wb = p.workbook

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
  end

  def test_mixed_borders_2
    p = Axlsx::Package.new
    wb = p.workbook

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

    wb.apply_styles

    assert_equal 1, wb.styles.dxfs.count

    assert_equal 6, wb.styles.cellXfs.count
  end

  def test_default_font_with_style_index
    p = Axlsx::Package.new
    wb = p.workbook

    wb.styles.fonts[0].name = 'Times New Roman'
    wb.styles.fonts[0].sz = 12

    wb.add_worksheet do |sheet|
      sheet.add_row [1, 2, 3]
      sheet.add_style "A1:C1", { color: "FFFFFF" }
    end

    wb.apply_styles

    assert_equal 1, wb.styles.style_index.size

    assert_equal(
      {
        type: :xf,
        name: "Times New Roman",
        sz: 12,
        family: 1,
        color: "FFFFFF"
      },
      wb.styles.style_index.values.first
    )
  end

  def test_escape_formulas
    @wb.escape_formulas = false
    @ws = @wb.add_worksheet
    assert_false @ws.escape_formulas
    assert_false @ws.add_row(['']).cells.first.escape_formulas
    assert_false @ws.add_row([''], escape_formulas: false).cells.first.escape_formulas
    assert @ws.add_row([''], escape_formulas: true).cells.first.escape_formulas
    assert_equal [true, false], @ws.add_row(['', ''], escape_formulas: [true, false]).cells.map(&:escape_formulas)

    @ws = Axlsx::Worksheet.new(@wb)
    assert_false @ws.escape_formulas

    @wb.escape_formulas = true
    @ws = @wb.add_worksheet
    assert @ws.escape_formulas
    assert @ws.add_row(['']).cells.first.escape_formulas
    assert_false @ws.add_row([''], escape_formulas: false).cells.first.escape_formulas
    assert @ws.add_row([''], escape_formulas: true).cells.first.escape_formulas
    assert_equal [true, false], @ws.add_row(['', ''], escape_formulas: [true, false]).cells.map(&:escape_formulas)

    @ws = Axlsx::Worksheet.new(@wb)
    assert @ws.escape_formulas

    @ws.escape_formulas = false
    assert_false @ws.escape_formulas
    assert_false @ws.add_row(['']).cells.first.escape_formulas
    assert_false @ws.add_row([''], escape_formulas: false).cells.first.escape_formulas
    assert @ws.add_row([''], escape_formulas: true).cells.first.escape_formulas
    assert_equal [true, false], @ws.add_row(['', ''], escape_formulas: [true, false]).cells.map(&:escape_formulas)

    @ws.escape_formulas = true
    assert @ws.escape_formulas
    assert @ws.add_row(['']).cells.first.escape_formulas
    assert_false @ws.add_row([''], escape_formulas: false).cells.first.escape_formulas
    assert @ws.add_row([''], escape_formulas: true).cells.first.escape_formulas
    assert_equal [true, false], @ws.add_row(['', ''], escape_formulas: [true, false]).cells.map(&:escape_formulas)
  end
end
