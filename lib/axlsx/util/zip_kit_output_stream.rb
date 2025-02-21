# frozen_string_literal: true

require "zip_kit"

module Axlsx
  # The ZipKitOutputStream allows generation of a streaming ZIP file.
  #
  # The methods provided here mimic `Zip::OutputStream` so that this class can be used a drop-in replacement.
  class ZipKitOutputStream
    # Create a temporary file for writing the ZIP out, residing at `file_name`
    #
    # The directory and its contents are removed at the end of the block.
    def self.open(file_path_or_writable)
      if file_path_or_writable.is_a?(String) # Assume it is a path
        File.open(file_path_or_writable, "wb") do |f|
          ZipKit::Streamer.open(f) do |streamer|
            yield(streamer)
          end
        end
      elsif file_path_or_writable.respond_to?(:write)
        ZipKit::Streamer.open(file_path_or_writable) do |streamer|
          yield(new(streamer))
        end
      else
        raise ArgumentError, "file_path_or_writable muse be a path to file or respond to write()"
      end
    end

    def self.write_buffer(&blk)
      StringIO.new.binmode.tap do |out|
        ZipKit::Streamer.open(out, &blk)
        out.rewind
      end
    end
  end
end
