# frozen_string_literal: true

require 'tc_helper'
require 'support/capture_warnings'

class TestDataValidation < Minitest::Test
  include CaptureWarnings

  def setup
    # inverse defaults
    @boolean_options = { allowBlank: false, hideDropDown: true, showErrorMessage: false, showInputMessage: true }
    @nil_options = { formula1: 'foo', formula2: 'foo', errorTitle: 'foo', operator: :lessThan, prompt: 'foo', promptTitle: 'foo', sqref: 'foo' }
    @type_option = { type: :whole }
    @error_style_option = { errorStyle: :warning }

    @string_options = { formula1: 'foo', formula2: 'foo', error: 'foo', errorTitle: 'foo', prompt: 'foo', promptTitle: 'foo', sqref: 'foo' }
    @symbol_options = { errorStyle: :warning, operator: :lessThan, type: :whole }

    @options = @boolean_options.merge(@nil_options, @type_option, @error_style_option)

    @dv = Axlsx::DataValidation.new(@options)
  end

  def test_initialize
    dv = Axlsx::DataValidation.new

    @boolean_options.each do |key, value|
      assert_equal(!value, dv.send(key.to_sym), "initialized default #{key} should be #{!value}")
      assert_equal(value, @dv.send(key.to_sym), "initialized options #{key} should be #{value}")
    end

    @nil_options.each do |key, value|
      assert_nil(dv.send(key.to_sym), "initialized default #{key} should be nil")
      assert_equal(value, @dv.send(key.to_sym), "initialized options #{key} should be #{value}")
    end

    @type_option.each do |key, value|
      assert_equal(:none, dv.send(key.to_sym), "initialized default #{key} should be :none")
      assert_equal(value, @dv.send(key.to_sym), "initialized options #{key} should be #{value}")
    end

    @error_style_option.each do |key, value|
      assert_equal(:stop, dv.send(key.to_sym), "initialized default #{key} should be :stop")
      assert_equal(value, @dv.send(key.to_sym), "initialized options #{key} should be #{value}")
    end
  end

  def test_boolean_attribute_validation
    @boolean_options.each do |key, value|
      assert_raises(ArgumentError, "#{key} must be boolean") { @dv.send(:"#{key}=", 'A') }
      refute_raises { @dv.send(:"#{key}=", value) }
    end
  end

  def test_string_attribute_validation
    @string_options.each do |key, value|
      assert_raises(ArgumentError, "#{key} must be string") { @dv.send(:"#{key}=", :symbol) }
      refute_raises { @dv.send(:"#{key}=", value) }
    end
  end

  def test_symbol_attribute_validation
    @symbol_options.each do |key, value|
      assert_raises(ArgumentError, "#{key} must be symbol") { @dv.send(:"#{key}=", "foo") }
      refute_raises { @dv.send(:"#{key}=", value) }
    end
  end

  def test_formula1
    assert_raises(ArgumentError) { @dv.formula1 = 10 }
    refute_raises { @dv.formula1 = "=SUM(A1:A1)" }
    assert_equal("=SUM(A1:A1)", @dv.formula1)
  end

  def test_formula2
    assert_raises(ArgumentError) { @dv.formula2 = 10 }
    refute_raises { @dv.formula2 = "=SUM(A1:A1)" }
    assert_equal("=SUM(A1:A1)", @dv.formula2)
  end

  def test_allowBlank
    assert_raises(ArgumentError) { @dv.allowBlank = "foo´" }
    refute_raises { @dv.allowBlank = false }
    assert_false(@dv.allowBlank)
  end

  def test_error
    assert_raises(ArgumentError) { @dv.error = :symbol }
    refute_raises { @dv.error = "This is a error message" }
    assert_equal("This is a error message", @dv.error)
  end

  def test_errorStyle
    assert_raises(ArgumentError) { @dv.errorStyle = "foo" }
    refute_raises { @dv.errorStyle = :information }
    assert_equal(:information, @dv.errorStyle)
  end

  def test_errorTitle
    assert_raises(ArgumentError) { @dv.errorTitle = :symbol }
    refute_raises { @dv.errorTitle = "This is the error title" }
    assert_equal("This is the error title", @dv.errorTitle)
  end

  def test_operator
    assert_raises(ArgumentError) { @dv.operator = "foo" }
    refute_raises { @dv.operator = :greaterThan }
    assert_equal(:greaterThan, @dv.operator)
  end

  def test_prompt
    assert_raises(ArgumentError) { @dv.prompt = :symbol }
    refute_raises { @dv.prompt = "This is a prompt message" }
    assert_equal("This is a prompt message", @dv.prompt)
  end

  def test_promptTitle
    assert_raises(ArgumentError) { @dv.promptTitle = :symbol }
    refute_raises { @dv.promptTitle = "This is the prompt title" }
    assert_equal("This is the prompt title", @dv.promptTitle)
  end

  def test_showDropDown
    warnings = capture_warnings do
      assert_raises(ArgumentError) { @dv.showDropDown = "foo´" }
      refute_raises { @dv.showDropDown = false }
      assert_false(@dv.showDropDown)
    end

    assert_equal 2, warnings.size
    assert_includes warnings.first, 'The `showDropDown` has an inverted logic, false shows the dropdown list! You should use `hideDropDown` instead.'
  end

  def test_hideDropDown
    assert_raises(ArgumentError) { @dv.hideDropDown = "foo´" }
    refute_raises { @dv.hideDropDown = false }
    assert_false(@dv.hideDropDown)
    # As hideDropdown is just an alias for showDropDown, we should test the original value too
    assert_false(@dv.showDropDown)
  end

  def test_showErrorMessage
    assert_raises(ArgumentError) { @dv.showErrorMessage = "foo´" }
    refute_raises { @dv.showErrorMessage = false }
    assert_false(@dv.showErrorMessage)
  end

  def test_showInputMessage
    assert_raises(ArgumentError) { @dv.showInputMessage = "foo´" }
    refute_raises { @dv.showInputMessage = false }
    assert_false(@dv.showInputMessage)
  end

  def test_sqref
    assert_raises(ArgumentError) { @dv.sqref = 10 }
    refute_raises { @dv.sqref = "A1:A1" }
    assert_equal("A1:A1", @dv.sqref)
  end

  def test_type
    assert_raises(ArgumentError) { @dv.type = "foo" }
    refute_raises { @dv.type = :list }
    assert_equal(:list, @dv.type)
  end

  def test_whole_decimal_data_time_textLength_to_xml
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet name: "data_validation"
    @ws.add_data_validation("A1", { type: :whole, operator: :between, formula1: '5', formula2: '10',
                                    showErrorMessage: true, errorTitle: 'Wrong input', error: 'Only values between 5 and 10',
                                    errorStyle: :information, showInputMessage: true, promptTitle: 'Be careful!',
                                    prompt: 'Only values between 5 and 10' })

    doc = Nokogiri::XML.parse(@ws.to_xml_string)

    # test attributes
    assert_equal(1, doc.xpath("//xmlns:worksheet/xmlns:dataValidations[@count='1']/xmlns:dataValidation[@sqref='A1']
      [@promptTitle='Be careful!'][@prompt='Only values between 5 and 10'][@operator='between'][@errorTitle='Wrong input']
      [@error='Only values between 5 and 10'][@showErrorMessage=1][@allowBlank=1][@showInputMessage=1][@type='whole']
      [@errorStyle='information']").size)
    assert doc.xpath("//xmlns:worksheet/xmlns:dataValidations[@count='1']/xmlns:dataValidation[@sqref='A1']
      [@promptTitle='Be careful!'][@prompt='Only values between 5 and 10'][@operator='between'][@errorTitle='Wrong input']
      [@error='Only values between 5 and 10'][@showErrorMessage=1][@allowBlank=1][@showInputMessage=1]
      [@type='whole'][@errorStyle='information']")

    # test forumula1
    assert_equal(1, doc.xpath("//xmlns:worksheet/xmlns:dataValidations/xmlns:dataValidation/xmlns:formula1").size)
    assert doc.xpath("//xmlns:worksheet/xmlns:dataValidations/xmlns:dataValidation/xmlns:formula1='5'")

    # test forumula2
    assert_equal(1, doc.xpath("//xmlns:worksheet/xmlns:dataValidations/xmlns:dataValidation/xmlns:formula2").size)
    assert doc.xpath("//xmlns:worksheet/xmlns:dataValidations/xmlns:dataValidation/xmlns:formula2='10'")
  end

  def test_list_to_xml
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet name: "data_validation"
    @ws.add_data_validation("A1", { type: :list, formula1: 'A1:A5',
                                    showErrorMessage: true, errorTitle: 'Wrong input', error: 'Only values from list',
                                    errorStyle: :stop, showInputMessage: true, promptTitle: 'Be careful!',
                                    prompt: 'Only values from list', hideDropDown: true })

    doc = Nokogiri::XML.parse(@ws.to_xml_string)

    # test attributes
    assert_equal(1, doc.xpath("//xmlns:worksheet/xmlns:dataValidations[@count='1']/xmlns:dataValidation[@sqref='A1']
      [@promptTitle='Be careful!'][@prompt='Only values from list'][@errorTitle='Wrong input'][@error='Only values from list']
      [@showErrorMessage=1][@allowBlank=1][@showInputMessage=1][@showDropDown=1][@type='list']
      [@errorStyle='stop']").size)
    assert doc.xpath("//xmlns:worksheet/xmlns:dataValidations[@count='1']/xmlns:dataValidation[@sqref='A1']
      [@promptTitle='Be careful!'][@prompt='Only values from list'][@errorTitle='Wrong input'][@error='Only values from list']
      [@showErrorMessage=1][@allowBlank=1][@showInputMessage=1][@showDropDown=1][@type='list'][@errorStyle='stop']")

    # test forumula1
    assert_equal(1, doc.xpath("//xmlns:worksheet/xmlns:dataValidations/xmlns:dataValidation/xmlns:formula1").size)
    assert doc.xpath("//xmlns:worksheet/xmlns:dataValidations/xmlns:dataValidation/xmlns:formula1='A1:A5'")
  end

  def test_custom_to_xml
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet name: "data_validation"
    @ws.add_data_validation("A1", { type: :custom, formula1: '=5/2',
                                    showErrorMessage: true, errorTitle: 'Wrong input', error: 'Only values corresponding formula',
                                    errorStyle: :stop, showInputMessage: true, promptTitle: 'Be careful!',
                                    prompt: 'Only values corresponding formula' })

    doc = Nokogiri::XML.parse(@ws.to_xml_string)

    # test attributes
    assert_equal(1, doc.xpath("//xmlns:worksheet/xmlns:dataValidations[@count='1']/xmlns:dataValidation[@sqref='A1'][@promptTitle='Be careful!']
      [@prompt='Only values corresponding formula'][@errorTitle='Wrong input'][@error='Only values corresponding formula'][@showErrorMessage=1]
      [@allowBlank=1][@showInputMessage=1][@type='custom'][@errorStyle='stop']").size)
    assert doc.xpath("//xmlns:worksheet/xmlns:dataValidations[@count='1']/xmlns:dataValidation[@sqref='A1'][@promptTitle='Be careful!']
      [@prompt='Only values corresponding formula'][@errorTitle='Wrong input'][@error='Only values corresponding formula']
      [@showErrorMessage=1][@allowBlank=1][@showInputMessage=1][@type='custom'][@errorStyle='stop']")

    # test forumula1
    assert_equal(1, doc.xpath("//xmlns:worksheet/xmlns:dataValidations/xmlns:dataValidation/xmlns:formula1").size)
    assert doc.xpath("//xmlns:worksheet/xmlns:dataValidations/xmlns:dataValidation/xmlns:formula1='=5/2'")
  end

  def test_none_to_xml
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet name: "data_validation"
    @ws.add_data_validation("A1", { type: :none,
                                    showInputMessage: true, promptTitle: 'Be careful!',
                                    prompt: 'This is a warning to be extra careful editing this cell' })

    doc = Nokogiri::XML.parse(@ws.to_xml_string)

    # test attributes
    assert_equal(1, doc.xpath("//xmlns:worksheet/xmlns:dataValidations[@count='1']/xmlns:dataValidation[@sqref='A1']
      [@promptTitle='Be careful!'][@prompt='This is a warning to be extra careful editing this cell']
      [@allowBlank=1][@showInputMessage=1][@type='none']").size)
    assert doc.xpath("//xmlns:worksheet/xmlns:dataValidations[@count='1']/xmlns:dataValidation[@sqref='A1']
      [@promptTitle='Be careful!'][@prompt='This is a warning to be extra careful editing this cell']
      [@allowBlank=1][@showInputMessage=1][@type='none']")
  end

  def test_multiple_datavalidations_to_xml
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet name: "data_validation"
    @ws.add_data_validation("A1", { type: :whole, operator: :between, formula1: '5', formula2: '10',
                                    showErrorMessage: true, errorTitle: 'Wrong input', error: 'Only values between 5 and 10',
                                    errorStyle: :information, showInputMessage: true, promptTitle: 'Be careful!',
                                    prompt: 'Only values between 5 and 10' })
    @ws.add_data_validation("B1", { type: :list, formula1: 'A1:A5',
                                    showErrorMessage: true, errorTitle: 'Wrong input', error: 'Only values from list',
                                    errorStyle: :stop, showInputMessage: true, promptTitle: 'Be careful!',
                                    prompt: 'Only values from list', hideDropDown: true })

    doc = Nokogiri::XML.parse(@ws.to_xml_string)

    # test attributes
    assert_equal(1, doc.xpath("//xmlns:worksheet/xmlns:dataValidations[@count='2']/xmlns:dataValidation[@sqref='A1']
      [@promptTitle='Be careful!'][@prompt='Only values between 5 and 10'][@operator='between'][@errorTitle='Wrong input']
      [@error='Only values between 5 and 10'][@showErrorMessage=1][@allowBlank=1][@showInputMessage=1][@type='whole']
      [@errorStyle='information']").size)
    assert doc.xpath("//xmlns:worksheet/xmlns:dataValidations[@count='2']/xmlns:dataValidation[@sqref='A1']
      [@promptTitle='Be careful!'][@prompt='Only values between 5 and 10'][@operator='between'][@errorTitle='Wrong input']
      [@error='Only values between 5 and 10'][@showErrorMessage=1][@allowBlank=1][@showInputMessage=1]
      [@type='whole'][@errorStyle='information']")

    # test attributes
    assert_equal(1, doc.xpath("//xmlns:worksheet/xmlns:dataValidations[@count='2']/xmlns:dataValidation[@sqref='B1']
      [@promptTitle='Be careful!'][@prompt='Only values from list'][@errorTitle='Wrong input'][@error='Only values from list']
      [@showErrorMessage=1][@allowBlank=1][@showInputMessage=1][@showDropDown=1][@type='list']
      [@errorStyle='stop']").size)
    assert doc.xpath("//xmlns:worksheet/xmlns:dataValidations[@count='2']/xmlns:dataValidation[@sqref='B1']
      [@promptTitle='Be careful!'][@prompt='Only values from list'][@errorTitle='Wrong input'][@error='Only values from list']
      [@showErrorMessage=1][@allowBlank=1][@showInputMessage=1][@showDropDown=1][@type='list'][@errorStyle='stop']")
  end

  def test_empty_attributes
    v = Axlsx::DataValidation.new

    assert_equal([:allowBlank,
                  :error,
                  :errorStyle,
                  :errorTitle,
                  :prompt,
                  :promptTitle,
                  :showErrorMessage,
                  :showInputMessage,
                  :sqref,
                  :type], v.send(:get_valid_attributes))
  end
end
