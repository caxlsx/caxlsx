# frozen_string_literal: true

source 'https://rubygems.org'
gemspec

group :development, :test do
  gem 'rubocop', '~> 1.51.0'
  gem 'rubocop-minitest', '~> 0.31.0'
  gem 'rubocop-performance', '~> 1.18.0'
end

group :test do
  gem 'rake'
  gem 'simplecov', '>= 0.14.1'
  gem 'test-unit'
  gem 'webmock'
end

group :profile do
  gem 'memory_profiler'
  gem 'ruby-prof', platforms: :ruby
end
