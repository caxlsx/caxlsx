source "https://rubygems.org"
gemspec

group :test do
  gem 'rake'
  gem "simplecov"
  gem 'test-unit'
end

group :profile do
  gem 'ruby-prof', :platforms => :ruby
end

platforms :rbx do
  gem 'rubysl'
  gem 'rubysl-test-unit'
  gem 'racc'
  gem 'rubinius-coverage', '~> 2.0'
end