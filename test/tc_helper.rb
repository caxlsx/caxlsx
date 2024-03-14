# frozen_string_literal: true

$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'simplecov'
SimpleCov.start do
  add_filter "/test/"
  add_filter "/vendor/"
end

require 'minitest/autorun'
require 'timecop'
require 'webmock/minitest'
require 'axlsx'

module Minitest
  class Test
    def assert_false(value)
      assert_equal(false, value)
    end

    def refute_raises
      yield
    rescue StandardError => e
      raise Minitest::Assertion, "Expected no exception, but raised: #{e.class.name} with message '#{e.message}'"
    end
  end
end
