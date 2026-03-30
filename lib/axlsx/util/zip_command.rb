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
      @current_file = nil
      @files = []
      @zip_command = zip_command
    end

    # Create a temporary directory for writing files to.
    #
    # The directory and its contents are removed at the end of the block.
    def open(output)
      Dir.mktmpdir do |dir|
        @dir = dir
        yield(self)
        write_file
        zip_parts(output)
      end
    end

    # Closes the current entry and opens a new for writing.
    def put_next_entry(entry)
      write_file
      @current_file = "#{@dir}/#{entry.name}"
      @files << entry.name
      FileUtils.mkdir_p(File.dirname(@current_file))
      @io = File.open(@current_file, "wb")
    end

    # Write to a buffer that will be written to the current entry
    def write(content)
      @io << content
    end
    alias << write

    private

    def write_file
      @io.close if @current_file
      @current_file = nil
      @io = nil
    end

    def zip_parts(output)
      output = Shellwords.shellescape(File.absolute_path(output))
      inputs = Shellwords.shelljoin(@files)
      escaped_dir = Shellwords.shellescape(@dir)
      command = "cd #{escaped_dir} && #{@zip_command} #{output} #{inputs}"
      stdout_and_stderr, status = Open3.capture2e(command)
      unless status.success?
        raise ZipError, stdout_and_stderr
      end
    end
  end
end
