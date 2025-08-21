# frozen_string_literal: true

require 'tc_helper'

class TestEncryptionCompatibility < Test::Unit::TestCase
  def setup
    skip_unless_windows_with_excel
    @test_password = 'test123'
    @temp_files = []
  end

  def teardown
    # Clean up any temporary files
    @temp_files.each do |file|
      FileUtils.rm_f(file)
    end
  end

  def test_caxlsx_encrypted_file_opens_in_excel
    # Create a basic workbook with theme
    package = Axlsx::Package.new
    workbook = package.workbook
    workbook.add_worksheet(name: 'Encryption Test') do |sheet|
      sheet.add_row ['Theme', 'Encryption', 'Test']
      sheet.add_row [1, 2, 3]
      sheet.add_row ['Success', 'Expected', 'Result']
    end

    # Generate unencrypted file
    unencrypted_file = 'test_encryption_unencrypted.xlsx'
    package.serialize(unencrypted_file)
    @temp_files << unencrypted_file

    # Verify unencrypted file opens normally
    assert_excel_file_opens(unencrypted_file, nil, "Unencrypted file should open in Excel")

    # Encrypt the file
    encrypted_file = 'test_encryption_encrypted.xlsx'
    OoxmlCrypt.encrypt_file(unencrypted_file, @test_password, encrypted_file)
    @temp_files << encrypted_file

    # Verify encrypted file opens with password
    assert_excel_file_opens(encrypted_file, @test_password, "Encrypted file should open in Excel with password")
  end

  def test_theme_xml_contains_required_elements_for_encryption
    package = Axlsx::Package.new
    theme_xml = package.workbook.theme.to_xml_string

    # These elements are critical for Excel encryption compatibility
    required_elements = [
      '<a:theme',
      '<a:themeElements>',
      '<a:clrScheme',
      '<a:fontScheme',
      '<a:fmtScheme',
      '<a:fillStyleLst>',
      '<a:lnStyleLst>',
      '<a:effectStyleLst>',
      '<a:bgFillStyleLst>',
      '<a:objectDefaults>',
      '<a:scene3d>',
      '<a:sp3d>',
      '<a:extraClrSchemeLst/>'
    ]

    required_elements.each do |element|
      assert_includes theme_xml, element, "Theme XML missing required element: #{element}"
    end
  end

  def test_complex_workbook_encryption_compatibility
    # Create a more complex workbook to test comprehensive compatibility
    package = Axlsx::Package.new
    workbook = package.workbook

    # Add multiple worksheets
    ws1 = workbook.add_worksheet(name: 'Data Sheet')
    ws1.add_row ['Name', 'Value', 'Category']
    10.times do |i|
      ws1.add_row ["Item #{i + 1}", rand(100), ['A', 'B', 'C'].sample]
    end

    ws2 = workbook.add_worksheet(name: 'Charts')
    ws2.add_row ['Month', 'Sales']
    ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'].each_with_index do |month, _i|
      ws2.add_row [month, rand(500..1499)]
    end

    # Add some styling
    ws1.rows[0].cells.each { |cell| cell.style = 1 }

    # Generate and test
    complex_file = 'test_complex_encryption.xlsx'
    package.serialize(complex_file)
    @temp_files << complex_file

    # Encrypt and test
    encrypted_complex = 'test_complex_encrypted.xlsx'
    OoxmlCrypt.encrypt_file(complex_file, @test_password, encrypted_complex)
    @temp_files << encrypted_complex

    assert_excel_file_opens(encrypted_complex, @test_password, "Complex encrypted workbook should open in Excel")
  end

  private

  def skip_unless_windows_with_excel
    unless windows_platform?
      skip("Excel encryption compatibility tests only run on Windows")
    end

    unless excel_available?
      skip("Excel encryption compatibility tests require Microsoft Excel to be installed")
    end

    return if defined?(OoxmlCrypt)

    skip("Excel encryption compatibility tests require ooxml_crypt gem")
  end

  def assert_excel_file_opens(file_path, password = nil, message = nil)
    return true unless excel_available?

    begin
      require 'win32ole'
      excel = WIN32OLE.new('Excel.Application')
      excel.visible = false
      excel.displayAlerts = false

      absolute_path = File.absolute_path(file_path)

      workbook = if password
                   excel.Workbooks.Open(absolute_path, nil, nil, nil, password)
                 else
                   excel.Workbooks.Open(absolute_path)
                 end

      # File opened successfully
      workbook.Close(false)
      excel.Quit
      true
    rescue StandardError => e
      # Clean up Excel process if it's still running
      begin
        excel&.Quit
      rescue StandardError
      end

      # Re-raise the error for test failure
      error_msg = "Excel failed to open file '#{file_path}': #{e.message}"
      error_msg = "#{message}: #{error_msg}" if message

      flunk(error_msg)
    ensure
      # Ensure Excel process is terminated
      begin
        excel&.Quit
      rescue StandardError
      end
    end
  end
end
