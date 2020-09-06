#!/usr/bin/env ruby

files = if !ARGV.empty?
          ARGV.select { |file| File.exist?(file) }
        else
          Dir['*_example.md']
        end

files.each do |file|
  puts "Executing #{file.split('.')[0].tr('_', ' ')}"
  code = File.read(file).match(/```ruby(?<code>.+)```/m)[:code]
  unless code.nil?
    eval(['$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"', code].join("\n"))
  end
end
