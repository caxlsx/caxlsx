#!/usr/bin/env ruby

files = if ARGV.length > 0
          ARGV.select { |file| File.exists?(file) }
        else
          Dir['*_example.md']
        end

files.each do |file|
  puts "Executing #{file.split('.')[0].tr('_', ' ')}"
  code = File.read(file).match(/```ruby(?<code>.+)```/m)[:code]
  eval(code) unless code.nil?
end
