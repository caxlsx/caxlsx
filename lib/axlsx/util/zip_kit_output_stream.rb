# frozen_string_literal: true

require "zip_kit"

module Axlsx
  # Allows generation of a ZIP file.
  class ZipKitOutputStream
    # Sets up a ZipKit::Streamer for writing the XLSX file and yields it to the caller.
    # At the end of the block the central directory and EOCD marker will be written into the output.
    #
    # @yield [ZipKit::Streamer]
    # @param file_path_or_writable[String,#IO] Either a path to a file - the ZIP will be written there, or an object that responds to #write -
    #     that object will receive the ZIP binary data.
    def self.open(file_path_or_writable)
      if file_path_or_writable.respond_to?(:write) # Assume it is something writable, like a Rails `stream`
        ZipKit::Streamer.open(file_path_or_writable) do |streamer|
          yield(new(streamer))
        end
      elsif file_path_or_writable.is_a?(String) # Assume it is a path
        File.open(file_path_or_writable, "wb") do |f|
          ZipKit::Streamer.open(f) do |streamer|
            yield(streamer)
          end
        end
      else
        raise ArgumentError, "file_path_or_writable muse be a path to file or respond to write()"
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
