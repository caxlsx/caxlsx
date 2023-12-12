# frozen_string_literal: true

source 'https://rubygems.org'
gemspec

group :development, :test do
  gem 'kramdown'
  gem 'yard'

  if RUBY_VERSION >= '2.7'
    gem 'rubocop', '~> 1.59.0'
    gem 'rubocop-minitest', '~> 0.33.0'
    gem 'rubocop-performance', '~> 1.19.1'
  end
end

group :test do
  gem 'rake'
  gem 'simplecov'
  gem 'test-unit'
  gem 'timecop'
  gem 'webmock'
end

group :profile do
  gem 'memory_profiler'
  gem 'ruby-prof', platforms: :ruby
end
