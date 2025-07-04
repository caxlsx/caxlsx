# frozen_string_literal: true

module Axlsx
  # A SimpleTypedList is a type restrictive collection that allows some of the methods from Array and supports basic xml serialization.
  # @private
  class SimpleTypedList < Array
    DESTRUCTIVE = [
      'replace', 'insert', 'collect!', 'map!', 'pop', 'delete_if',
      'reverse!', 'shift', 'shuffle!', 'slice!', 'sort!', 'uniq!',
      'unshift', 'zip', 'flatten!', 'fill', 'drop', 'drop_while',
      'clear'
    ].freeze

    DESTRUCTIVE.each do |name|
      undef_method name
    end

    # We often call index(element) on instances of SimpleTypedList. Thus, we do not want to inherit Array
    # implementation of == / eql? which walks the elements calling == / eql?. Instead we want the fast
    # and original versions from BasicObject.
    alias :== :equal?
    alias :eql? :equal?

    # Creates a new typed list
    # @param [Array, Class] type An array of Class objects or a single Class object
    # @param [String] serialize_as The tag name to use in serialization
    # @raise [ArgumentError] if all members of type are not Class objects
    def initialize(type, serialize_as = nil, start_size = 0)
      super(start_size)

      if type.is_a? Array
        type.each { |item| raise ArgumentError, "All members of type must be Class objects" unless item.is_a? Class }
        @allowed_types = type
      else
        raise ArgumentError, "Type must be a Class object or array of Class objects" unless type.is_a? Class

        @allowed_types = [type]
      end
      @serialize_as = serialize_as unless serialize_as.nil?
    end

    # The class constants of allowed types
    # @return [Array]
    attr_reader :allowed_types

    # The index below which items cannot be removed
    # @return [Integer]
    def locked_at
      defined?(@locked_at) ? @locked_at : nil
    end

    # The tag name to use when serializing this object
    # by default the parent node for all items in the list is the classname of the first allowed type with the first letter in lowercase.
    # @return [String]
    attr_reader :serialize_as

    # Transposes the list (without blowing up like ruby does)
    # any non populated cell in the matrix will be a nil value
    def transpose
      return clone if size == 0

      row_count = size
      max_column_count = map { |row| row.cells.size }.max
      result = Array.new(max_column_count) { Array.new(row_count) }
      # yes, I know it is silly, but that warning is really annoying
      row_count.times do |row_index|
        max_column_count.times do |column_index|
          datum = if self[row_index].cells.size >= max_column_count
                    self[row_index].cells[column_index]
                  elsif block_given?
                    yield(column_index, row_index)
                  end
          result[column_index][row_index] = datum
        end
      end
      result
    end

    # Lock this list at the current size
    # @return [self]
    def lock
      @locked_at = size
      self
    end

    # Unlock the list
    # @return [self]
    def unlock
      @locked_at = nil
      self
    end

    # Appends the elements of +others+ to self.
    # @param [Array<Array>] others one or more arrays to join
    # @raise [ArgumentError] if any of the values being joined are not
    # one of the allowed types
    # @return [SimpleTypedList]
    def concat(*others)
      others.each do |other|
        other.each do |item|
          self << item
        end
      end
      self
    end

    # Pushes the given object on to the end of this array and returns the index
    # of the item added.
    # @param [Any] v the data to be added
    # @raise [ArgumentError] if the value being added is not one of the allowed types
    # @return [Integer] returns the index of the item added.
    def <<(v)
      DataTypeValidator.validate :SimpleTypedList_push, @allowed_types, v
      super
      size - 1
    end

    # Pushes the given object(s) on to the end of this array. This expression
    # returns the array itself, so several appends may be chained together.
    # @param [Any] values the data to be added
    # @raise [ArgumentError] if any of the values being joined are not
    # @return [SimpleTypedList]
    def push(*values)
      values.each do |value|
        self << value
      end
      self
    end

    # delete the item from the list
    # @param [Any] v The item to be deleted.
    # @raise [ArgumentError] if the item's index is protected by locking
    # @return [Any] The item deleted
    def delete(v)
      return unless include? v
      raise ArgumentError, "Item is protected and cannot be deleted" if protected? index(v)

      super
    end

    # delete the item from the list at the index position provided
    # @raise [ArgumentError] if the index is protected by locking
    # @return [Any] The item deleted
    def delete_at(index)
      raise ArgumentError, "Item is protected and cannot be deleted" if protected? index

      super
    end

    # positional assignment. Adds the item at the index specified
    # @param [Integer] index
    # @param [Any] v
    # @raise [ArgumentError] if the index is protected by locking
    # @raise [ArgumentError] if the item is not one of the allowed types
    # @return [Any] The item added
    def []=(index, v)
      DataTypeValidator.validate :SimpleTypedList_insert, @allowed_types, v
      raise ArgumentError, "Item is protected and cannot be changed" if protected? index

      super
    end

    # inserts an item at the index specified
    # @param [Integer] index
    # @param [Any] v
    # @raise [ArgumentError] if the index is protected by locking
    # @raise [ArgumentError] if the index is not one of the allowed types
    # @return [Any] The item inserted
    def insert(index, v)
      DataTypeValidator.validate :SimpleTypedList_insert, @allowed_types, v
      raise ArgumentError, "Item is protected and cannot be changed" if protected? index

      super
      v
    end

    # determines if the index is protected
    # @param [Integer] index
    def protected?(index)
      return false unless locked_at.is_a? Integer

      index < locked_at
    end

    def to_xml_string(str = +'')
      el_name = serialize_as.to_s
      str << '<' << el_name << ' count="' << size.to_s << '">'
      each { |item| item.to_xml_string(str) }
      str << '</' << el_name << '>'
    end
  end
end
