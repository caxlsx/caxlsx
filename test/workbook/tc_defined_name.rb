# frozen_string_literal: true

require 'tc_helper'

class TestDefinedNames < Minitest::Test
  def setup
    @dn = Axlsx::DefinedName.new('Sheet1!A1:A1')
  end

  def test_initialize
    assert_equal('Sheet1!A1:A1', @dn.formula)
  end

  def test_string_attributes
    %w(short_cut_key status_bar help description custom_menu comment).each do |attr|
      assert_raises(ArgumentError, 'only strings allowed in string attributes') { @dn.send(:"#{attr}=", 1) }
      refute_raises { @dn.send(:"#{attr}=", '_xlnm.Sheet_Title') }
    end
  end

  def test_boolean_attributes
    %w(workbook_parameter publish_to_server xlm vb_proceedure function hidden).each do |attr|
      assert_raises(ArgumentError, 'only booleanish allowed in string attributes') { @dn.send(:"#{attr}=", 'foo') }
      refute_raises { @dn.send(:"#{attr}=", 1) }
    end
  end

  def test_local_sheet_id
    assert_raises(ArgumentError, 'local_sheet_id must be an unsigned int') { @dn.local_sheet_id = -1 }
    refute_raises { @dn.local_sheet_id = 1 }
  end

  def test_do_not_camelcase_value_for_name
    @dn.name = '_xlnm._FilterDatabase'
    doc = Nokogiri::XML(@dn.to_xml_string)

    assert_equal(1, doc.xpath("//definedName[@name='_xlnm._FilterDatabase']").size)
    assert_equal('Sheet1!A1:A1', doc.xpath('//definedName').text)
  end

  def test_to_xml_string
    assert_raises(ArgumentError, 'name is required for serialization') { @dn.to_xml_string }
    @dn.name = '_xlnm.Print_Titles'
    @dn.hidden = true
    doc = Nokogiri::XML(@dn.to_xml_string)

    assert_equal(1, doc.xpath("//definedName[@name='_xlnm.Print_Titles']").size)
    assert_equal(1, doc.xpath("//definedName[@hidden='1']").size)
    assert_equal('Sheet1!A1:A1', doc.xpath('//definedName').text)
  end
end
