#!/usr/bin/env ruby -s
# frozen_string_literal: true

$:.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'
require 'ruby-prof'

# Axlsx.trust_input = true
Axlsx.skip_validations = true

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
  p.serialize("example_prof.xlsx", zip_command: 'zip')
end

printer = RubyProf::FlatPrinter.new(profile)
printer.print(STDOUT, {})

File.delete("example_prof.xlsx")