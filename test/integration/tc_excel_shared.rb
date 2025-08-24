# frozen_string_literal: true

# Shared cases for Excel integration
module TestExcelShared
  def setup
    require_or_skip!
    setup_excel_application
    @temp_files = []
  end

  def teardown
    teardown_excel_application
    @temp_files&.each do |file|
      FileUtils.rm_f(file)
    end
  end

  def test_simple_excel_file_creation_and_opening
    # Create a basic workbook
    package = Axlsx::Package.new
    workbook = package.workbook
    workbook.add_worksheet(name: 'Basic Test') do |sheet|
      sheet.add_row %w[Name Value Category]
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
end
