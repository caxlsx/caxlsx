# frozen_string_literal: true

source 'https://rubygems.org'
gemspec

group :development, :test do
  gem 'kramdown'
  gem 'yard'

  if RUBY_VERSION >= '2.7'
    gem 'rubocop', '1.79.2'
    gem 'rubocop-minitest', '0.38.1'
    gem 'rubocop-packaging', '0.6.0'
    gem 'rubocop-performance', '1.25.0'
  end
end

group :test do
  gem 'rake'
  gem 'simplecov'
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
