# frozen_string_literal: true

source 'https://rubygems.org'
gemspec

if RUBY_VERSION >= '2.7'
  group :development, :test do
    gem 'rubocop', '~> 1.58.0'
    gem 'rubocop-minitest', '~> 0.33.0'
    gem 'rubocop-performance', '~> 1.19.1'
  end
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
