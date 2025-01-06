# frozen_string_literal: true

module Axlsx
  # The BufferedZipOutputStream buffers the output in order to avoid appending many small strings directly to the
  # the `Zip::OutputStream`.
  #
  # The methods provided here mimic `Zip::OutputStream` so that this class can be used a drop-in replacement.
  class BufferedZipOutputStream
    # The 4_096 was chosen somewhat arbitrary, however, it was difficult to see any obvious improvement with larger
    # buffer sizes.
    BUFFER_SIZE = 4_096

    def initialize(zos)
      @zos = zos
      @buffer = String.new(capacity: BUFFER_SIZE * 2)
    end

    # Create a temporary directory for writing files to.
    #
    # The directory and its contents are removed at the end of the block.
    def self.open(file_name, encrypter = nil)
      Zip::OutputStream.open(file_name, encrypter: encrypter) do |zos|
        bzos = new(zos)
        yield(bzos)
      ensure
        bzos.flush if bzos
      end
    end

    def self.write_buffer(io = ::StringIO.new, encrypter = nil)
      Zip::OutputStream.write_buffer(io, encrypter: encrypter) do |zos|
        bzos = new(zos)
        yield(bzos)
      ensure
        bzos.flush if bzos
      end
    end

    # Closes the current entry and opens a new for writing.
    def put_next_entry(entry)
      flush
      @zos.put_next_entry(entry)
    end

    # Write to a buffer that will be written to the current entry
    def write(content)
      @buffer << content.to_s
      flush if @buffer.size > BUFFER_SIZE
      self
    end
    alias << write

    def flush
      return if @buffer.empty?

      @zos << @buffer
      @buffer.clear
    end
  end
end
