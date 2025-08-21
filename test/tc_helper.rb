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
    private

    def assert_false(value)
      assert_equal(false, value)
    end

    def refute_raises
      yield
    rescue StandardError => e
      raise Minitest::Assertion, "Expected no exception, but raised: #{e.class.name} with message '#{e.message}'"
    end

    def macos_platform?
      RUBY_PLATFORM.include?('darwin')
    end

    def windows_platform?
      RUBY_PLATFORM =~ /mswin|mingw|cygwin/
    end
  end
end
