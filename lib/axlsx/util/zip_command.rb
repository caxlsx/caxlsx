# frozen_string_literal: true

require 'open3'
require 'shellwords'

module Axlsx
  # The ZipCommand class supports zipping the Excel file contents using
  # a binary zip program instead of RubyZip's `Zip::OutputStream`.
  #
  # The methods provided here mimic `Zip::OutputStream` so that `ZipCommand` can
  # be used as a drop-in replacement. Note that method signatures are not
  # identical to `ZipKitOutputStream`, they are only sufficiently close so that
  # `ZipCommand` and `ZipKitOutputStream` can be interchangeably used within
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
    #
    # @param write_to_file_at_path[String] path for the XLSX to write 
    def open(write_to_file_at_path)
      Dir.mktmpdir do |dir|
        @dir = dir
        yield(self)
        run_zip_command(write_to_file_at_path)
      end
    end

    # Adds a file to the ZIP. Yields an IO the file data can be written to.
    # Works the same as `ZipKit::Streamer#write_file`
    # @param filename[String] the name of the file in the ZIP
    # @param modification_time[Time] the mtime to set for the entry in the ZIP
    # @yield [#write] the writable IO for the file' binary data
    #
    # @return [void]
    def write_file(filename, modification_time:, &blk)
      raise ArgumentError, "@dir not set" unless @dir

      tempfile_path = File.join(@dir, filename)

      FileUtils.mkdir_p(File.dirname(tempfile_path))
      File.open(tempfile_path, "wb", &blk)
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
