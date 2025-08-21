# frozen_string_literal: true

require 'tc_helper'

class TestExcelIntegration < Test::Unit::TestCase
  def setup
    skip_unless_windows_with_excel
    setup_excel_application
    @temp_files = []
  end

  def teardown
    teardown_excel_application
    @temp_files.each do |file|
      FileUtils.rm_f(file)
    end
  end

  def test_simple_excel_file_creation_and_opening
    # Create a basic workbook
    package = Axlsx::Package.new
    workbook = package.workbook
    workbook.add_worksheet(name: 'Basic Test') do |sheet|
      sheet.add_row ['Name', 'Value', 'Category']
      sheet.add_row ['Item 1', 100, 'A']
      sheet.add_row ['Item 2', 200, 'B']
    end

    # Generate file
    test_file = 'test_basic.xlsx'
    package.serialize(test_file)
    @temp_files << test_file

    # Verify file opens in Excel
    assert_excel_file_opens(test_file)
  end

  def test_excel_file_cell_values_readable
    # Create a workbook with specific values we can test
    package = Axlsx::Package.new
    workbook = package.workbook
    workbook.add_worksheet(name: 'Cell Test') do |sheet|
      sheet.add_row ['Hello', 'World', 42]
      sheet.add_row ['Test', 'Value', 3.14]
    end

    # Generate file
    test_file = 'test_cell_values.xlsx'
    package.serialize(test_file)
    @temp_files << test_file

    # Verify file opens and we can read cell values
    assert_excel_cell_values(test_file, {
      'A1' => 'Hello',
      'B1' => 'World',
      'C1' => 42,
      'A2' => 'Test',
      'B2' => 'Value',
      'C2' => 3.14
    })
  end

  def test_multiple_worksheets
    # Create a workbook with multiple sheets
    package = Axlsx::Package.new
    workbook = package.workbook

    ws1 = workbook.add_worksheet(name: 'Sheet1')
    ws1.add_row ['Data', 'Sheet', 1]

    ws2 = workbook.add_worksheet(name: 'Sheet2')
    ws2.add_row ['Second', 'Worksheet', 2]

    # Generate file
    test_file = 'test_multiple_sheets.xlsx'
    package.serialize(test_file)
    @temp_files << test_file

    # Verify file opens and check worksheet contents
    assert_excel_file_opens(test_file)
    
    # Verify Sheet1 contents
    assert_excel_cell_values_by_sheet(test_file, 'Sheet1', {
      'A1' => 'Data',
      'B1' => 'Sheet',
      'C1' => 1
    })
    
    # Verify Sheet2 contents
    assert_excel_cell_values_by_sheet(test_file, 'Sheet2', {
      'A1' => 'Second',
      'B1' => 'Worksheet',
      'C1' => 2
    })
  end

  private

  def skip_unless_windows_with_excel
    unless self.class.windows_platform?
      skip("Excel integration tests only run on Windows")
    end

    return if self.class.excel_windows?

    skip("Excel integration tests require Microsoft Excel to be installed")
  end

  def setup_excel_application
    @excel = WIN32OLE.new('Excel.Application')
    @excel.visible = false
    @excel.displayAlerts = false
  end

  def teardown_excel_application
    @excel&.Quit
  rescue StandardError
    # Ignore errors during cleanup
  ensure
    @excel = nil
  end

  def assert_excel_file_opens(file_path)
    with_workbook(file_path) do |workbook|
      # File opened successfully
    end
    true
  end

  def assert_excel_cell_values(file_path, expected_values)
    with_workbook(file_path) do |workbook|
      worksheet = workbook.Worksheets(1)

      expected_values.each do |cell_address, expected_value|
        actual_value = worksheet.Range(cell_address).Value

        assert_equal expected_value, actual_value,
                     "Cell #{cell_address} should contain '#{expected_value}' but contains '#{actual_value}'"
      end
    end
    true
  end

  def assert_excel_cell_values_by_sheet(file_path, sheet_name, expected_values)
    with_workbook(file_path) do |workbook|
      worksheet = workbook.Worksheets(sheet_name)

      expected_values.each do |cell_address, expected_value|
        actual_value = worksheet.Range(cell_address).Value

        assert_equal expected_value, actual_value,
                     "Cell #{cell_address} in sheet '#{sheet_name}' should contain '#{expected_value}' but contains '#{actual_value}'"
      end
    end
    true
  end

  def with_workbook(file_path)
    absolute_path = File.absolute_path(file_path)
    workbook = @excel.Workbooks.Open(absolute_path)
    yield workbook
  ensure
    workbook&.Close(false)
  end

  def self.windows_platform?
    RUBY_PLATFORM =~ /mswin|mingw|cygwin/
  end

  def self.excel_windows?
    return @excel_windows if defined?(@excel_windows)

    @excel_windows = windows_platform? &&
                     defined?(WIN32OLE) &&
                     begin
                       excel = WIN32OLE.new('Excel.Application')
                       excel.Quit
                       true
                     rescue StandardError
                       false
                     end
  end
end
