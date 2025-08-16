# frozen_string_literal: true

require 'bundler'
Bundler::GemHelper.install_tasks

task :benchmark do
  require File.expand_path("#{File.dirname(__FILE__)}/test/benchmark.rb")
end

task :gendoc do
  system "yardoc"
  system "yard stats --list-undoc"
end

require 'rake/testtask'
Rake::TestTask.new do |t|
  t.libs << 'test'
  t.test_files = FileList['test/**/tc_*.rb']
  t.verbose = false
  t.warning = true
end

task default: :test
