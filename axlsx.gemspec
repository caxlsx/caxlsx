# frozen_string_literal: true

require File.expand_path('lib/axlsx/version', __dir__)

Gem::Specification.new do |s|
  s.name        = 'caxlsx'
  s.version     = Axlsx::VERSION
  s.authors     = ["Randy Morgan", "Jurriaan Pruis"]
  s.email       = 'noel@peden.biz'
  s.homepage    = 'https://github.com/caxlsx/caxlsx'
  s.platform    = Gem::Platform::RUBY
  s.summary     = "Excel OOXML (xlsx) with charts, styles, images and autowidth columns."
  s.license     = 'MIT'
  s.description = <<~MSG
    xlsx spreadsheet generation with charts, images, automated column width, customizable styles and full schema validation. Axlsx helps you create beautiful Office Open XML Spreadsheet documents (Excel, Google Spreadsheets, Numbers, LibreOffice) without having to understand the entire ECMA specification. Check out the README for some examples of how easy it is. Best of all, you can validate your xlsx file before serialization so you know for sure that anything generated is going to load on your client's machine.
  MSG
  s.files = Dir.glob("{lib/**/*,examples/**/*.rb,examples/**/*.jpeg}") + %w[LICENSE README.md Rakefile CHANGELOG.md .yardopts .yardopts_guide]

  s.metadata = {
    'bug_tracker_uri' => 'https://github.com/caxlsx/caxlsx/issues',
    'changelog_uri' => 'https://github.com/caxlsx/caxlsx/blob/master/CHANGELOG.md',
    'source_code_uri' => 'https://github.com/caxlsx/caxlsx',
    'rubygems_mfa_required' => 'true'
  }

  s.add_dependency "htmlentities", "~> 4.3", '>= 4.3.4'
  s.add_dependency "marcel", '~> 1.0'
  s.add_dependency 'nokogiri', '~> 1.10', '>= 1.10.4'
  s.add_dependency 'rubyzip', '>= 2.4', '< 4'

  s.required_ruby_version = '>= 2.6'
  s.require_path = 'lib'
end
