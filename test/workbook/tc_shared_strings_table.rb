# frozen_string_literal: true

require 'tc_helper'

class TestSharedStringsTable < Minitest::Test
  def setup
    @p = Axlsx::Package.new use_shared_strings: true

    ws = @p.workbook.add_worksheet
    ws.add_row ['a', 1, 'b']
    ws.add_row ['b', 1, 'c']
    ws.add_row ['c', 1, 'd']
    ws.rows.last.add_cell('b', type: :text)
  end

  def test_workbook_has_shared_strings
    assert_kind_of(Axlsx::SharedStringsTable, @p.workbook.shared_strings, "shared string table was not created")
  end

  def test_count
    sst = @p.workbook.shared_strings

    assert_equal(7, sst.count)
  end

  def test_unique_count
    sst = @p.workbook.shared_strings

    assert_equal(4, sst.unique_count)
  end

  def test_uses_workbook_xml_space
    assert_equal(@p.workbook.xml_space, @p.workbook.shared_strings.xml_space)
    @p.workbook.xml_space = :default

    assert_equal(:default, @p.workbook.shared_strings.xml_space)
  end

  def test_valid_document
    schema = Nokogiri::XML::Schema(File.open(Axlsx::SML_XSD))
    doc = Nokogiri::XML(@p.workbook.shared_strings.to_xml_string)
    errors = schema.validate(doc)

    assert_empty(errors)
  end

  def test_remove_control_characters_in_xml_serialization
    nasties = "hello\x10\x00\x1C\x1Eworld"
    @p.workbook.worksheets[0].add_row [nasties]

    # test that the nasty string was added to the shared strings
    assert @p.workbook.shared_strings.unique_cells.key?(nasties)

    # test that none of the control characters are in the XML output for shared strings
    refute_match(/#{Axlsx::CONTROL_CHARS}/o, @p.workbook.shared_strings.to_xml_string)

    # assert that the shared string was normalized to remove the control characters
    refute_nil @p.workbook.shared_strings.to_xml_string.index("helloworld")
  end
end
