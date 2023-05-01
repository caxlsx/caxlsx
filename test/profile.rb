#!/usr/bin/env ruby -s
# frozen_string_literal: true

$:.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'
require 'ruby-prof'

row = []
input1 = (32..126).to_a.pack('U*').chars.to_a # these will need to be escaped
input2 = (65..122).to_a.pack('U*').chars.to_a # these do not need to be escaped
10.times { row << input1.shuffle.join }
10.times { row << input2.shuffle.join }

profile = RubyProf.profile do
  p = Axlsx::Package.new
  p.workbook.add_worksheet do |sheet|
    10000.times do
      sheet << row
    end
  end
  p.to_stream
end

printer = RubyProf::FlatPrinter.new(profile)
printer.print(STDOUT, {})
