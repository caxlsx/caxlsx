# frozen_string_literal: true

require 'open3'
require 'shellwords'

module Axlsx
  # The ZipCommand class supports zipping the Excel file contents using
  # a binary zip program instead of RubyZip's `Zip::OutputStream`.
  #
  # The methods provided here mimic `Zip::OutputStream` so that `ZipCommand` can
  # be used as a drop-in replacement. Note that method signatures are not
  # identical to `Zip::OutputStream`, they are only sufficiently close so that
  # `ZipCommand` and `Zip::OutputStream` can be interchangeably used within
  # `caxlsx`.
  class ZipCommand
    # Raised when the zip command exits with a non-zero status.
    class ZipError < StandardError; end

    def initialize(zip_command)
      @filenames = []
      @zip_command = zip_command
    end

    # Create a temporary directory for writing files to.
    #
    # The directory and its contents are removed at the end of the block.
    def open(write_to_file_at_path)
      Dir.mktmpdir do |dir|
        @dir = dir
        yield(self)
        run_zip_command(write_to_file_at_path)
      end
    end

    # Closes the current entry and opens a new for writing.
    def write_file(filename, modification_time:)
      raise ArgumentError, "@dir not set" unless @dir

      tempfile_path = File.join(@dir, filename)

      FileUtils.mkdir_p(File.dirname(tempfile_path))
      File.open(tempfile_path, "wb") { |fo| yield(fo) }
      File.utime(modification_time, modification_time, tempfile_path)

      @filenames << filename
    end

    private

    def run_zip_command(write_to_file_at_path)
      # Note that the number of files may overflow the shell argument list
      output = Shellwords.shellescape(File.absolute_path(write_to_file_at_path))
      inputs = Shellwords.shelljoin(@filenames)
      escaped_dir = Shellwords.shellescape(@dir)
      command = "cd #{escaped_dir} && #{@zip_command} #{output} #{inputs}"
      stdout_and_stderr, status = Open3.capture2e(command)
      unless status.success?
        raise ZipError, stdout_and_stderr
      end
    end
  end
end
