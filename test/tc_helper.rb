# frozen_string_literal: true

$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"

require 'simplecov'
SimpleCov.start do
  add_filter "/test/"
  add_filter "/vendor/"
end

require 'minitest/autorun'
require 'rspec/mocks/minitest_integration'
require 'timecop'
require 'webmock/minitest'
require 'axlsx'
require 'ooxml_crypt' if RUBY_ENGINE == 'ruby'

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

    def macos?
      RbConfig::CONFIG['host_os'].match?(/darwin/i)
    end

    def windows?
      RbConfig::CONFIG['host_os'].match?(/windows|mswin|mingw|cygwin/i)
    end

    def mri?
      RUBY_ENGINE == 'ruby'
    end

    def jruby?
      RUBY_ENGINE == 'jruby'
    end

    def truffleruby?
      RUBY_ENGINE == 'truffleruby'
    end

    def ooxml_crypt_available?
      defined?(OoxmlCrypt)
    end
  end
end
