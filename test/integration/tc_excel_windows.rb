# frozen_string_literal: true

require 'tc_helper'
require_relative 'tc_excel_shared'

class TestExcelWindows < Minitest::Test
  include TestExcelShared

  private

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
    with_workbook(file_path) {} # File opened successfully
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

  def require_or_skip!
    flunk('This test requires Excel Windows integration') if ENV['CI_EXCEL_WINDOWS'] && !excel_windows_mri?
    skip('Excel Windows integration tests only run on Windows MRI Ruby') unless windows? && mri?
    skip('Excel Windows integration tests require Microsoft Excel to be installed') unless excel_windows_mri?
  end

  def excel_windows_mri?
    windows? && mri? && excel_app_available?
  end

  def excel_app_available?
    return @excel_app_available if defined?(@excel_app_available)

    require 'win32ole'
    WIN32OLE.new('Excel.Application').Quit
    @excel_app_available = true
  rescue StandardError
    @excel_app_available = false
  end
end
