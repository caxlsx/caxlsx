# frozen_string_literal: true

source 'https://rubygems.org'
gemspec

group :development, :test do
  gem 'kramdown'
  gem 'yard'

  if RUBY_ENGINE == 'ruby'
    gem 'ooxml_crypt'
  end

  if RUBY_VERSION >= '2.7'
    gem 'rubocop', '1.73.2'
    gem 'rubocop-minitest', '0.37.1'
    gem 'rubocop-packaging', '0.5.2'
    gem 'rubocop-performance', '1.24.0'
  end
end

group :test do
  gem 'rake'
  gem 'simplecov'
  gem 'minitest'
  gem 'timecop'
  gem 'webmock'
  gem 'rspec-mocks'
end

group :profile do
  gem 'memory_profiler'
  gem 'ruby-prof', platforms: :ruby
end
