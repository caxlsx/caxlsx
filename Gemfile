# frozen_string_literal: true

source 'https://rubygems.org'
gemspec

group :development, :test do
  gem 'kramdown'
  gem 'yard'

  if RUBY_VERSION >= '2.7'
    gem 'rubocop', '1.91.0'
    gem 'rubocop-minitest', '0.40.0'
    gem 'rubocop-packaging', '0.6.0'
    gem 'rubocop-performance', '1.27.0'
  end
end

group :test do
  gem 'rake'
  # There's a bug in simplecov 1.3.0, see: https://github.com/simplecov-ruby/simplecov/issues/1299
  gem 'simplecov', '< 1.3'
  gem 'minitest'
  gem 'timecop'
  gem 'webmock'
  gem 'rspec-mocks'
  gem 'win32ole', platforms: [:mingw, :x64_mingw, :mswin, :mswin64]

  if RUBY_ENGINE == 'ruby'
    gem 'ooxml_crypt'
  end
end

group :profile do
  gem 'memory_profiler'
  gem 'ruby-prof', platforms: :ruby
end
