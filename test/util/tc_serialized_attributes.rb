# frozen_string_literal: true

require 'tc_helper'

class Funk
  include Axlsx::Accessors
  include Axlsx::SerializedAttributes

  serializable_attributes :camel_symbol, :boolean, :integer

  attr_accessor :camel_symbol, :boolean, :integer
end

class TestSeralizedAttributes < Minitest::Test
  def setup
    @object = Funk.new
  end

  def test_camel_symbol
    @object.camel_symbol = :foo_bar

    assert_equal('camelSymbol="fooBar" ', @object.serialized_attributes)
  end
end
