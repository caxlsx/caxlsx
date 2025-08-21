# frozen_string_literal: true

require 'tc_helper'
require_relative 'tc_excel_shared'

class TestExcelMacOS < Minitest::Test
  include TestExcelShared

  private

  def setup_excel_application
    # Test minimal Excel launch to verify it's available and licensed
    script = <<~APPLESCRIPT
      try
        tell application "Microsoft Excel"
          activate
          delay 0.5
          -- Just verify app responds, don't create anything yet
          set app_name to name
          return "Success: " & app_name
        end tell
      on error errMsg
        return "Error: " & errMsg
      end try
    APPLESCRIPT

    result = `osascript -e '#{script.gsub("'", "\\'")}' 2>&1`.strip

    if result.include?('Error') || result.empty?
      raise "Excel not available or not licensed: #{result}"
    end

    @excel_available = true
  end

  def teardown_excel_application
    return unless @excel_available

    # Quit Excel if it's running
    script = <<~APPLESCRIPT
      try
        tell application "Microsoft Excel"
          quit
        end tell
      on error
        -- Excel might already be closed
      end try
    APPLESCRIPT

    `osascript -e '#{script.gsub("'", "\\'")}' 2>/dev/null`
  rescue StandardError
    # Ignore errors during cleanup
  ensure
    @excel_available = false
  end

  def assert_excel_file_opens(file_path)
    with_workbook(file_path) {} # File opened successfully
    true
  end

  def assert_excel_cell_values(file_path, expected_values)
    with_workbook(file_path) do |workbook_name|
      expected_values.each do |cell_address, expected_value|
        actual_value = get_cell_value(workbook_name, 1, cell_address)

        # Handle type conversions for AppleScript
        actual_value = normalize_cell_value(actual_value)
        expected_value = normalize_cell_value(expected_value)

        assert_equal expected_value, actual_value,
                     "Cell #{cell_address} should contain '#{expected_value}' but contains '#{actual_value}'"
      end
    end
    true
  end

  def assert_excel_cell_values_by_sheet(file_path, sheet_name, expected_values)
    with_workbook(file_path) do |workbook_name|
      expected_values.each do |cell_address, expected_value|
        actual_value = get_cell_value_by_sheet(workbook_name, sheet_name, cell_address)

        # Handle type conversions for AppleScript
        actual_value = normalize_cell_value(actual_value)
        expected_value = normalize_cell_value(expected_value)

        assert_equal expected_value, actual_value,
                     "Cell #{cell_address} in sheet '#{sheet_name}' should contain '#{expected_value}' but contains '#{actual_value}'"
      end
    end
    true
  end

  def with_workbook(file_path)
    absolute_path = File.absolute_path(file_path)
    workbook_name = nil

    # Open the workbook
    script = <<~APPLESCRIPT
      try
        tell application "Microsoft Excel"
          activate
          set workbook_ref to (open workbook "#{absolute_path}")
          set workbook_name to name of workbook_ref
          return workbook_name
        end tell
      on error errMsg
        return "Error: " & errMsg
      end try
    APPLESCRIPT

    result = `osascript -e '#{script.gsub("'", "\\'")}' 2>&1`.strip

    if result.include?('Error')
      flunk("Failed to open Excel file '#{file_path}': #{result}")
    end

    workbook_name = result
    yield workbook_name
  ensure
    # Close the workbook
    if workbook_name && !workbook_name.include?('Error')
      close_script = <<~APPLESCRIPT
        try
          tell application "Microsoft Excel"
            close workbook "#{workbook_name}" saving no
          end tell
        on error
          -- Workbook might already be closed
        end try
      APPLESCRIPT

      `osascript -e '#{close_script.gsub("'", "\\'")}' 2>/dev/null`
    end
  end

  def get_cell_value(workbook_name, sheet_index, cell_address)
    script = <<~APPLESCRIPT
      try
        tell application "Microsoft Excel"
          tell workbook "#{workbook_name}"
            tell worksheet #{sheet_index}
              set cell_value to value of range "#{cell_address}"
              return cell_value as string
            end tell
          end tell
        end tell
      on error errMsg
        return "Error: " & errMsg
      end try
    APPLESCRIPT

    result = `osascript -e '#{script.gsub("'", "\\'")}' 2>&1`.strip

    if result.include?('Error')
      flunk("Failed to read cell #{cell_address}: #{result}")
    end

    result
  end

  def get_cell_value_by_sheet(workbook_name, sheet_name, cell_address)
    script = <<~APPLESCRIPT
      try
        tell application "Microsoft Excel"
          tell workbook "#{workbook_name}"
            tell worksheet "#{sheet_name}"
              set cell_value to value of range "#{cell_address}"
              return cell_value as string
            end tell
          end tell
        end tell
      on error errMsg
        return "Error: " & errMsg
      end try
    APPLESCRIPT

    result = `osascript -e '#{script.gsub("'", "\\'")}' 2>&1`.strip

    if result.include?('Error')
      flunk("Failed to read cell #{cell_address} from sheet '#{sheet_name}': #{result}")
    end

    result
  end

  def normalize_cell_value(value)
    case value
    when /^\d+$/
      value.to_i
    when /^\d+\.\d+$/
      value.to_f
    else
      value.to_s
    end
  end

  def require_or_skip!
    flunk('This test requires Excel macOS integration') if ENV['CI_EXCEL_MACOS'] && !excel_macos?
    skip('Excel macOS integration tests only run on macOS') unless macos_platform?
    skip('Excel macOS integration tests require Microsoft Excel to be installed') unless excel_macos?
  end

  def macos_platform?
    RUBY_PLATFORM =~ /darwin/
  end

  def excel_macos?
    return @excel_macos if defined?(@excel_macos)

    @excel_macos = macos_platform? && excel_app_available?
  end

  def excel_app_available?
    # Check if Microsoft Excel app exists
    script = <<~APPLESCRIPT
      try
        tell application "Microsoft Excel"
          set app_name to name
          return "Available: " & app_name
        end tell
      on error errMsg
        return "Error: " & errMsg
      end try
    APPLESCRIPT

    result = `osascript -e '#{script.gsub("'", "\\'")}' 2>&1`.strip
    result.include?('Available:')
  rescue StandardError
    false
  end
end
