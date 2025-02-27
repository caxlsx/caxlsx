# frozen_string_literal: true

require "zip_kit"

module Axlsx
  # Allows generation of a ZIP file.
  class ZipKitOutputStream
    # Sets up a ZipKit::Streamer for writing the XLSX file and yields it to the caller.
    # At the end of the block the central directory and EOCD marker will be written into the output.
    #
    # @yield [ZipKit::Streamer]
    # @param file_path_or_writable[String,#write] Either a path to a file - the ZIP will be written there, or an object that responds to #write -
    #     that object will receive the ZIP binary data.
    def self.open(file_path_or_writable, &blk)
      if file_path_or_writable.is_a?(String)
        # Assume it is a path
        File.open(file_path_or_writable, "wb") { |f| open(f, &blk) }
      else
        # Assume it is something writable, like a Rails `stream` or any other destination
        # for ZipKit
        ZipKit::Streamer.open(file_path_or_writable, &blk)
      end
    end

    # Sets up a ZipKit::Streamer for writing the XLSX file and yields it to the caller.
    # At the end of the block the central directory and EOCD marker will be written into the output.
    #
    # @yield [ZipKit::Streamer]
    # @return [StringIO] containing the ZIP, backed by a binary string, having pointer at 0
    def self.write_buffer(&blk)
      StringIO.new.binmode.tap do |out|
        ZipKit::Streamer.open(out, &blk)
        out.rewind
      end
    end
  end
end
